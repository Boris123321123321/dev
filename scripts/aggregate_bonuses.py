#!/usr/bin/env python3
"""Агрегация выгрузки эко-бонусов в обезличенный JSON.

Сырая выгрузка содержит номера телефонов (персональные данные) и в репозиторий
не коммитится. Этот скрипт превращает её в агрегаты без ПД, которые уже можно
хранить в git и подавать на вход build_model.py.

Использование:
    python3 scripts/aggregate_bonuses.py <выгрузка.xlsx> \
        --start 2025-06-01 --end 2026-07-31 \
        -o data/bonus_aggregates.json

Ожидаемые колонки: bonuscard, phone, bonusvalue, segment_asr
"""

import argparse
import json
from datetime import date

import pandas as pd

# Квантили, которые сохраняем для описания распределения начислений.
QUANTILES = [0.0, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99, 1.0]


def months_between(start: date, end: date) -> int:
    """Число календарных месяцев в окне, включая крайние."""
    return (end.year - start.year) * 12 + (end.month - start.month) + 1


def aggregate(path: str, start: date, end: date) -> dict:
    df = pd.read_excel(path, sheet_name=0)

    missing = {"bonuscard", "bonusvalue", "segment_asr"} - set(df.columns)
    if missing:
        raise SystemExit(f"В выгрузке нет колонок: {sorted(missing)}")

    df = df.drop(columns=[c for c in ("phone",) if c in df.columns])
    df["bonusvalue"] = pd.to_numeric(df["bonusvalue"], errors="coerce").fillna(0.0)

    months = months_between(start, end)
    cards = int(len(df))
    total = float(df["bonusvalue"].sum())

    # Справочная разбивка по АСР-сегментам. В расчёте НЕ используется:
    # модель перешла на границы валовой выручки. Нужна только чтобы показать,
    # что начисление эко-бонусов почти не зависит от покупательной силы.
    grouped = (
        df.groupby(df["segment_asr"].fillna("(нет сегмента)"))["bonusvalue"]
        .agg(["size", "sum", "mean", "median"])
        .sort_values("sum", ascending=False)
    )
    per_user_month_overall = total / cards / months
    segments = [
        {
            "segment": str(name),
            "cards": int(row["size"]),
            "bonus_total": float(row["sum"]),
            "bonus_per_user_month": float(row["sum"] / row["size"] / months),
            "index_to_average": float(
                (row["sum"] / row["size"] / months) / per_user_month_overall
            ),
        }
        for name, row in grouped.iterrows()
    ]

    quantiles = {
        f"q{int(q * 100)}": float(df["bonusvalue"].quantile(q)) for q in QUANTILES
    }

    return {
        "source_file": path.rsplit("/", 1)[-1],
        "period_start": start.isoformat(),
        "period_end": end.isoformat(),
        "months": months,
        "cards": cards,
        "bonus_total": total,
        "bonus_per_month": total / months,
        "bonus_per_year": total / months * 12,
        "bonus_per_user_month": per_user_month_overall,
        "bonus_zero_cards": int((df["bonusvalue"] == 0).sum()),
        "bonus_negative_cards": int((df["bonusvalue"] < 0).sum()),
        "quantiles_over_period": quantiles,
        "segments_reference_only": segments,
    }


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("path", help="xlsx с выгрузкой бонусов")
    p.add_argument("--start", required=True, help="начало окна, YYYY-MM-DD")
    p.add_argument("--end", required=True, help="конец окна, YYYY-MM-DD")
    p.add_argument("-o", "--out", default="data/bonus_aggregates.json")
    a = p.parse_args()

    result = aggregate(a.path, date.fromisoformat(a.start), date.fromisoformat(a.end))
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=2)

    print(f"{a.out}: {result['cards']} карт, {result['bonus_total']:,.0f} ₽ "
          f"за {result['months']} мес → {result['bonus_per_user_month']:.2f} ₽/юзер/мес")


if __name__ == "__main__":
    main()
