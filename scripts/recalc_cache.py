#!/usr/bin/env python3
"""Пересчёт формул книги и запись кэшированных значений в файл.

Зачем: openpyxl пишет формулы без кэшированных значений. Пока значения не
записаны, любой читатель кэша (pandas, openpyxl с data_only=True, превью в
почте и мессенджерах) видит в формульных ячейках пустоту. Excel и LibreOffice
пересчитают книгу при открытии, но до этого файл выглядит пустым.

Штатный путь — пересчёт через LibreOffice. В песочнице, где заблокированы
AF_UNIX-сокеты, soffice не поднимается, поэтому формулы считаются движком
`formulas` (чистый Python), а значения дописываются в XML рядом с формулами.
Формулы, форматы и оформление остаются на месте.

    python3 scripts/recalc_cache.py output/eco_tiers_unit_economics.xlsx

Выход: JSON со статусом, числом формул и найденными ошибками.
"""

import argparse
import json
import re
import shutil
import sys
import warnings
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

warnings.filterwarnings("ignore")

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"

ERROR_LITERALS = {"#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"}


def scalar(value):
    """Достать одиночное значение из Ranges/ndarray движка formulas."""
    v = getattr(value, "value", value)
    while hasattr(v, "shape") or isinstance(v, (list, tuple)):
        try:
            if hasattr(v, "size") and v.size == 0:
                return None
            v = v[0] if not hasattr(v, "item") or v.shape else v.item()
        except (IndexError, ValueError, AttributeError):
            break
        if not isinstance(v, (list, tuple)) and not hasattr(v, "shape"):
            break
    if hasattr(v, "item"):
        try:
            v = v.item()
        except (ValueError, AttributeError):
            pass
    return v


def compute(path):
    """Посчитать все формулы книги. Возвращает {(ЛИСТ_UPPER, coord): value}."""
    import formulas

    model = formulas.ExcelModel().loads(str(path)).finish()
    sol = model.calculate()

    out = {}
    key_re = re.compile(r"^'\[.+?\](.+?)'!([A-Z]+\d+)$")
    for key, val in sol.items():
        m = key_re.match(key)
        if not m:
            continue
        out[(m.group(1).upper(), m.group(2))] = scalar(val)
    return out


def sheet_parts(zf):
    """Соответствие имя листа → путь к его XML внутри архива."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{PNS}}}Relationship")}

    parts = {}
    for sh in wb.find(f"{{{NS}}}sheets"):
        rid = sh.get(f"{{{RNS}}}id")
        t = target.get(rid, "")
        t = t[1:] if t.startswith("/") else f"xl/{t}"
        parts[sh.get("name")] = t.replace("xl/xl/", "xl/")
    return parts


def inject(path, values):
    """Дописать <v> к каждой формульной ячейке. Возвращает статистику."""
    path = Path(path)
    src = path.with_suffix(".src.tmp")
    shutil.copy2(path, src)

    stats = {"total_formulas": 0, "filled": 0, "empty": 0, "errors": {}}

    with zipfile.ZipFile(src) as zin:
        parts = sheet_parts(zin)
        rewritten = {}

        for name, part in parts.items():
            try:
                root = ET.fromstring(zin.read(part))
            except KeyError:
                continue
            changed = False

            for cell in root.iter(f"{{{NS}}}c"):
                f = cell.find(f"{{{NS}}}f")
                if f is None:
                    continue
                stats["total_formulas"] += 1
                coord = cell.get("r")
                val = values.get((name.upper(), coord))

                # убрать прежнее кэшированное значение, если было
                for old in cell.findall(f"{{{NS}}}v"):
                    cell.remove(old)

                if val is None or (isinstance(val, str) and val == ""):
                    stats["empty"] += 1
                    continue

                if isinstance(val, str) and val.strip() in ERROR_LITERALS:
                    err = val.strip()
                    stats["errors"].setdefault(err, []).append(f"{name}!{coord}")
                    cell.set("t", "e")
                    v = ET.SubElement(cell, f"{{{NS}}}v")
                    v.text = err
                    changed = True
                    continue

                if isinstance(val, bool):
                    cell.set("t", "b")
                    text = "1" if val else "0"
                elif isinstance(val, (int, float)):
                    cell.attrib.pop("t", None)
                    text = repr(float(val)) if isinstance(val, float) else str(val)
                else:
                    cell.set("t", "str")
                    text = str(val)

                v = ET.SubElement(cell, f"{{{NS}}}v")
                v.text = text
                stats["filled"] += 1
                changed = True

            if changed:
                rewritten[part] = ET.tostring(root, encoding="UTF-8",
                                              xml_declaration=True)

        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = rewritten.get(item.filename)
                zout.writestr(item, data if data is not None else zin.read(item.filename))

    src.unlink()
    stats["total_errors"] = sum(len(v) for v in stats["errors"].values())
    return stats


def main():
    p = argparse.ArgumentParser()
    p.add_argument("path")
    a = p.parse_args()

    ET.register_namespace("", NS)
    values = compute(a.path)
    stats = inject(a.path, values)

    stats["status"] = "errors_found" if stats["total_errors"] else "success"
    stats["errors"] = {k: v[:50] for k, v in stats["errors"].items()}
    print(json.dumps(stats, ensure_ascii=False, indent=2))

    if stats["total_errors"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
