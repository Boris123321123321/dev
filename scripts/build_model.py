#!/usr/bin/env python3
"""Сборка модели «Эко-тиры: unit-экономика по тирам и месяцам».

Что делает модель (и чем отличается от предыдущей версии):
  1. Считает unit-экономику — всё на одного пользователя в месяц, тиры и
     периоды получаются умножением юнита на количество юзеров.
  2. Тир определяется ТОЛЬКО границей валовой выручки за календарный месяц
     (Лес ≥ 35 000 ₽ и т.д.). Привязки к АСР-сегментам ВВ (Амбассадор,
     Средний, Репертуарный) в расчёте больше нет.
  3. База расчёта — валовая выручка; грязная маржа = ВВ × ставка грязной маржи.
  4. Разложение по каждому тиру и по каждому месяцу с 01.2026 (доходы и расходы).
  5. Все окна — в календарных месяцах. Окон «90 дней» и «30 дней» в модели нет.

Использование:
    python3 scripts/build_model.py \
        --bonus-aggregates data/bonus_aggregates.json \
        -o output/eco_tiers_unit_economics.xlsx

    # когда придёт выгрузка ВВ по пользователям с начала 2026 года:
    python3 scripts/build_model.py \
        --bonus-aggregates data/bonus_aggregates.json \
        --revenue выгрузка_ВВ.xlsx \
        -o output/eco_tiers_unit_economics.xlsx
"""

import argparse
import json
from math import erf, exp, log, sqrt

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- оформление

FONT = "Arial"
BLUE = "0000FF"      # захардкоженный ввод
BLACK = "000000"     # формула
GREEN = "008000"     # ссылка на другой лист
GREY = "595959"      # пояснение
YELLOW = "FFFF00"    # ключевое допущение / заполнить
FACT_FILL = "D9D9D9"  # факт, не менять
HDR_FILL = "1F4E79"

RUB = '#,##0 ₽;(#,##0 ₽);-'
RUB2 = '#,##0.00 ₽;(#,##0.00 ₽);-'
CNT = '#,##0;(#,##0);-'
PCT = '0.0%;(0.0%);-'
PCT2 = '0.00%;(0.00%);-'
MULT = '0.00"×"'

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ------------------------------------------------------------------- тиры

# Границы валовой выручки за календарный месяц. Источник: постановка задачи
# (Лес > 35 000 ₽). Множители перенесены из предыдущей версии модели.
TIERS = [
    {"key": "start",  "name": "🌱 Старт",  "low": 0,     "high": 3000,  "mult": 1.00},
    {"key": "sprout", "name": "🌿 Росток", "low": 3000,  "high": 12000, "mult": 1.40},
    {"key": "tree",   "name": "🌳 Дерево", "low": 12000, "high": 35000, "mult": 2.15},
    {"key": "forest", "name": "🌲 Лес",    "low": 35000, "high": None,  "mult": 2.85},
]

MONTHS_2026 = [f"2026-{m:02d}" for m in range(1, 13)]
LAST_FACT_MONTH = 7  # июль 2026 — последний закрытый месяц в выгрузке бонусов

GROSS_MARGIN = 0.30   # ставка грязной маржи, допущение
PLACEHOLDER_SIGMA = 1.2  # log-СКО распределения ВВ для плейсхолдера

SH_READ = "Читать сначала"
SH_ASSUM = "Допущения"
SH_BONUS = "Данные_бонусы"
SH_REV = "Выручка_вход"
SH_UNIT = "Unit-экономика"
SH_GRID = "Тир_Период"
SH_MIGR = "Миграция"
SH_SCEN = "Сценарии"


def q(sheet: str) -> str:
    """Имя листа для ссылки в формуле. Имена с пробелом обязаны быть в кавычках."""
    return f"'{sheet}'" if " " in sheet or "-" in sheet else sheet


def put(ws, coord, value, *, fmt=None, color=BLACK, bold=False, fill=None,
        italic=False, size=10, wrap=False, note=None, box=False, align=None):
    c = ws[coord]
    c.value = value
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if wrap or align:
        c.alignment = Alignment(wrap_text=wrap, vertical="top" if wrap else "center",
                                horizontal=align)
    if box:
        c.border = BOX
    if note:
        c.comment = Comment(note, "Модель")
    return c


def header(ws, row, labels, start_col=1, width=None):
    for i, label in enumerate(labels):
        col = get_column_letter(start_col + i)
        put(ws, f"{col}{row}", label, bold=True, color="FFFFFF", fill=HDR_FILL,
            wrap=True, align="center", box=True)
        ws.row_dimensions[row].height = 30
    if width:
        for i, w in enumerate(width):
            ws.column_dimensions[get_column_letter(start_col + i)].width = w


def title(ws, text, sub=None):
    put(ws, "A1", text, bold=True, size=14)
    if sub:
        put(ws, "A2", sub, italic=True, color=GREY, size=9)


# --------------------------------------------------- плейсхолдер по выручке

def _phi(z):
    return 0.5 * (1.0 + erf(z / sqrt(2.0)))


def lognormal_tier_split(mean, sigma, tiers):
    """Доля юзеров и условная средняя ВВ в каждом тире для логнормального ВВ.

    Нужно только чтобы засеять плейсхолдер до прихода фактической выгрузки ВВ
    по пользователям. Заменяется функцией tier_split_from_revenue().
    """
    mu = log(mean) - sigma * sigma / 2.0

    def F(t):
        return 1.0 if t is None else _phi((log(t) - mu) / sigma)

    def G(t):  # E[X·1{X<=t}] / mean
        return 1.0 if t is None else _phi((log(t) - mu - sigma * sigma) / sigma)

    out = []
    for t in tiers:
        low = max(t["low"], 1e-9)
        p = F(t["high"]) - F(low)
        e = mean * (G(t["high"]) - G(low))
        out.append({"share": p, "mean_revenue": e / p if p > 1e-12 else 0.0})
    return out


def tier_split_from_revenue(path, tiers):
    """Фактическое распределение из выгрузки ВВ по пользователям.

    Ожидается колонка с ВВ за календарный месяц (или суммарная ВВ + число
    месяцев). Ищем первую числовую колонку с подходящим именем.
    """
    df = pd.read_excel(path, sheet_name=0)
    candidates = [c for c in df.columns
                  if any(k in str(c).lower() for k in
                         ("вв", "выручк", "revenue", "gross", "turnover"))]
    if not candidates:
        raise SystemExit(f"В {path} не нашёл колонку с валовой выручкой. "
                         f"Колонки: {list(df.columns)}")
    col = candidates[0]
    rev = pd.to_numeric(df[col], errors="coerce").dropna()
    print(f"  выгрузка ВВ: колонка {col!r}, {len(rev)} строк, "
          f"средняя {rev.mean():,.0f} ₽")

    out = []
    for t in tiers:
        m = rev >= t["low"]
        if t["high"] is not None:
            m &= rev < t["high"]
        sub = rev[m]
        out.append({
            "share": len(sub) / len(rev),
            "mean_revenue": float(sub.mean()) if len(sub) else 0.0,
            "users": int(len(sub)),
        })
    return out, int(len(rev))


# ----------------------------------------------------------------- листы

def sheet_readme(wb, agg, is_placeholder):
    ws = wb.create_sheet(SH_READ)
    ws.column_dimensions["A"].width = 118
    title(ws, "Эко-тиры: unit-экономика прогрессивных бонусов за утилизацию",
          f"Окно бонусов: {agg['period_start']} — {agg['period_end']} "
          f"({agg['months']} мес). Разложение: по тирам и по месяцам с 01.2026.")

    blocks = [
        ("СУТЬ МЕХАНИКИ",
         "Коэффициент начисления эко-бонусов за сдачу вторсырья растёт вместе с валовой выручкой "
         "покупателя: 4 тира. Модель считает полную экономику — доход (грязная маржа с приростом ВВ) "
         "минус расход (все бонусные затраты, включая доплату тем, кто не менял поведение)."),

        ("ЧТО ИЗМЕНИЛОСЬ В ЭТОЙ ВЕРСИИ",
         "1. Расчёт переведён на unit-экономику: базовая единица — один пользователь за один месяц. "
         "Тиры и периоды получаются масштабированием юнита на число пользователей.\n"
         "2. Убрана связка с АСР-сегментами ВВ. Тир определяется только границей валовой выручки "
         "за календарный месяц: Старт < 3 000 ₽, Росток 3 000–12 000 ₽, Дерево 12 000–35 000 ₽, Лес ≥ 35 000 ₽.\n"
         "3. База расчёта — валовая выручка. Грязная маржа = ВВ × ставка грязной маржи (лист «Допущения»).\n"
         "4. Разложение по каждому тиру и по каждому месяцу начиная с 01.2026 — доходы и расходы (лист «Тир_Период»).\n"
         "5. Все окна — в календарных месяцах. Окон «90 дней» и «30 дней» в модели больше нет: "
         "квалификация в тир — за календарный месяц, эмиссия бонусов нормирована на "
         f"{agg['months']} месяцев фактической выгрузки."),

        ("ПОЧЕМУ СНЯТИЕ СВЯЗКИ С СЕГМЕНТАМИ СИЛЬНО МЕНЯЕТ ЦИФРЫ",
         "В предыдущей версии тир присваивался по сегменту: Амбассадор → Лес (×2,85), Средний → Дерево (×2,15). "
         "Но по собственной экономике сегментов из АСР-выгрузки средняя ВВ Амбассадора ≈ 10 330 ₽/мес "
         "(5,0 покупок × 2 066 ₽), то есть до порога Леса в 35 000 ₽ он не доходит — это уровень Ростка. "
         "Средний ≈ 1 212 ₽/мес — это Старт. Связка «сегмент → тир» завышала присвоенные множители, "
         "а вместе с ними и «слепую доплату». По границам выручки доплата получается в несколько раз меньше — "
         "см. лист «Тир_Период»."),

        ("ДАННЫЕ",
         f"• Факт (серое): выгрузка эко-бонусов {agg['period_start']} — {agg['period_end']}: "
         f"{agg['cards']:,} карт, {agg['bonus_total']:,.0f} ₽ начислено за {agg['months']} мес "
         f"= {agg['bonus_per_month']:,.0f} ₽/мес = {agg['bonus_per_user_month']:.2f} ₽/юзер/мес.\n"
         "• Границы тиров и множители — из постановки задачи (жёлтое, лист «Допущения»).\n"
         "• Ставка грязной маржи — допущение, требует подтверждения финансами."
         .replace(",", " ")),

        ("ЧЕГО ПОКА НЕТ — И ЧТО ЭТО ЗНАЧИТ",
         "Выгрузка фактической валовой выручки по пользователям с начала 2026 года запрошена у аналитиков, "
         "но ещё не пришла. Поэтому распределение пользователей по тирам и средняя ВВ внутри тира на листе "
         "«Выручка_вход» — ПЛЕЙСХОЛДЕР (помечен жёлтым), а не факт. "
         "Он засеян логнормальным распределением со средней ВВ, собранной из агрегатов АСР-выгрузки, "
         f"и log-СКО = {PLACEHOLDER_SIGMA} — только чтобы модель считалась от начала до конца и её можно было проверить.\n"
         "Когда выгрузка придёт, весь лист «Выручка_вход» пересобирается одной командой:\n"
         "    python3 scripts/build_model.py --bonus-aggregates data/bonus_aggregates.json --revenue <файл.xlsx>\n"
         "Ни одна формула ниже при этом не меняется — меняются только входные числа."
         if is_placeholder else
         "Распределение пользователей по тирам на листе «Выручка_вход» построено на фактической выгрузке "
         "валовой выручки по пользователям."),

        ("ЧТО НЕ ЯВЛЯЕТСЯ ФАКТОМ В РАСХОДНОЙ ЧАСТИ",
         "Выгрузка даёт суммарное начисление за 14 месяцев на карту, а не начисление по месяцам. "
         "Поэтому месячная эмиссия = среднее за окно × индекс месяца (индексы на листе «Допущения», "
         "по умолчанию 1,00 — то есть без сезонности). Разбивка бонусов по месяцам даст здесь факт.\n"
         "Разбивка базового эко-бонуса по тирам — тоже допущение: индекс 1,00 для всех тиров. "
         "Основание из факта: начисление на юзера почти не зависит от покупательной силы "
         "(по АСР-сегментам индекс к среднему 0,68–1,21 без монотонного тренда — лист «Данные_бонусы»)."),

        ("ЧЕСТНАЯ ОГОВОРКА ПРО ИНКРЕМЕНТАЛЬНОСТЬ",
         "Прирост маржи при переходе в тир — разница средних ВВ тиров; часть переходов случилась бы и без механики. "
         "Реальную конверсию даст только A/B с holdout. Поэтому в модели три сценария конверсии, "
         "а порог безубыточности показан отдельно (лист «Сценарии»)."),

        ("КАК ПОЛЬЗОВАТЬСЯ",
         "Серое — факт, не менять. Жёлтое — допущения и ячейки под заполнение, их можно крутить. "
         "Синий шрифт — введённое руками число. Чёрный — формула. Зелёный — ссылка на другой лист."),
    ]

    row = 4
    for head, body in blocks:
        put(ws, f"A{row}", head, bold=True, size=11)
        row += 1
        put(ws, f"A{row}", body, wrap=True, size=10)
        ws.row_dimensions[row].height = 14 * (body.count("\n") + 1 + len(body) // 118)
        row += 2
    return ws


def sheet_assumptions(wb, agg):
    ws = wb.create_sheet(SH_ASSUM)
    title(ws, "Допущения и границы тиров",
          "Жёлтое — можно менять. Границы — в валовой выручке за КАЛЕНДАРНЫЙ МЕСЯЦ.")

    # --- A. тиры
    put(ws, "A4", "A. Тиры: граница валовой выручки за календарный месяц", bold=True, size=11)
    header(ws, 5, ["Тир", "ВВ/мес от, ₽", "ВВ/мес до, ₽", "Множитель бонуса"],
           width=[26, 16, 16, 18])
    for i, t in enumerate(TIERS):
        r = 6 + i
        put(ws, f"A{r}", t["name"], box=True)
        put(ws, f"B{r}", t["low"], fmt=RUB, color=BLUE, fill=YELLOW, box=True)
        put(ws, f"C{r}", t["high"] if t["high"] is not None else "без ограничения",
            fmt=RUB if t["high"] is not None else None, color=BLUE, fill=YELLOW, box=True)
        put(ws, f"D{r}", t["mult"], fmt=MULT, color=BLUE, fill=YELLOW, box=True)
    put(ws, "F6", "Тир присваивается по валовой выручке за месяц. "
                  "Привязки к АСР-сегментам ВВ в расчёте нет.",
        italic=True, color=GREY, size=9, wrap=True)

    # --- B. экономика
    put(ws, "A12", "B. Экономика", bold=True, size=11)
    header(ws, 13, ["Параметр", "Значение", "Ед.", "Источник / комментарий"],
           width=[26, 16, 16, 62])

    put(ws, "A14", "Ставка грязной маржи", box=True)
    put(ws, "B14", GROSS_MARGIN, fmt=PCT, color=BLUE, fill=YELLOW, box=True,
        note="Грязная маржа = валовая выручка × эта ставка.\n"
             "ТРЕБУЕТ ПОДТВЕРЖДЕНИЯ финансами.\n"
             "В прошлой версии модели поля «Маржа ВВ/мес» были несовместимы с ВВ "
             "тех же сегментов (239% и 149% от выручки), поэтому вывести ставку из них нельзя.")
    put(ws, "C14", "% от ВВ", box=True)
    put(ws, "D14", "ДОПУЩЕНИЕ. Из полей «Маржа ВВ/мес» прошлой версии ставку вывести нельзя: "
                   "для Репертуарного это 239%, для Среднего 149% от их же ВВ.",
        italic=True, color=GREY, size=9, wrap=True, box=True)

    put(ws, "A15", "Базовый эко-бонус", box=True)
    put(ws, "B15", agg["bonus_per_user_month"], fmt=RUB2, color=BLUE, fill=FACT_FILL, box=True)
    put(ws, "C15", "₽/юзер/мес", box=True)
    put(ws, "D15", f"ФАКТ: {agg['bonus_total']:,.0f} ₽ / {agg['cards']:,} карт / "
                   f"{agg['months']} мес. Источник: {agg['source_file']}."
        .replace(",", " "), italic=True, color=GREY, size=9, wrap=True, box=True)

    put(ws, "A16", "Окно квалификации в тир", box=True)
    put(ws, "B16", 1, fmt='0" мес"', color=BLUE, fill=YELLOW, box=True)
    put(ws, "C16", "мес", box=True)
    put(ws, "D16", "Заменяет прежнее условие «за 30 дн покупок». Все окна модели — в месяцах.",
        italic=True, color=GREY, size=9, wrap=True, box=True)

    put(ws, "A17", "Окно выгрузки бонусов", box=True)
    put(ws, "B17", "=(YEAR(B19)-YEAR(B18))*12+MONTH(B19)-MONTH(B18)+1",
        fmt='0" мес"', box=True)
    put(ws, "C17", "мес", box=True)
    put(ws, "D17", "Число месяцев — формула от дат окна (B18, B19), а не константа. "
                   "Заменяет прежнюю нормировку «×365/90».",
        italic=True, color=GREY, size=9, wrap=True, box=True)

    put(ws, "A18", "Начало окна выгрузки", box=True)
    put(ws, "B18", pd.Timestamp(agg["period_start"]).date(), fmt="DD.MM.YYYY",
        color=BLUE, fill=FACT_FILL, box=True)
    put(ws, "C18", "дата", box=True)
    put(ws, "D18", "ФАКТ из выгрузки бонусов.", italic=True, color=GREY, size=9, box=True)

    put(ws, "A19", "Конец окна выгрузки", box=True)
    put(ws, "B19", pd.Timestamp(agg["period_end"]).date(), fmt="DD.MM.YYYY",
        color=BLUE, fill=FACT_FILL, box=True)
    put(ws, "C19", "дата", box=True)
    put(ws, "D19", "ФАКТ из выгрузки бонусов.", italic=True, color=GREY, size=9, box=True)

    put(ws, "A20", "Последний закрытый месяц факта 2026", box=True)
    put(ws, "B20", LAST_FACT_MONTH, fmt='0', color=BLUE, fill=FACT_FILL, box=True)
    put(ws, "C20", "номер мес", box=True)
    put(ws, "D20", "Месяцы 2026 года до этого включительно попадают в окно выгрузки бонусов, "
                   "дальше — прогноз.", italic=True, color=GREY, size=9, wrap=True, box=True)

    # --- C. индекс базового бонуса по тирам
    put(ws, "A22", "C. Индекс базового эко-бонуса по тирам", bold=True, size=11)
    put(ws, "A23", "Во сколько раз начисление на юзера в тире отличается от среднего. "
                   "По умолчанию 1,00 — начисление не зависит от тира.",
        italic=True, color=GREY, size=9)
    header(ws, 24, ["Тир", "Индекс", "Базовый бонус, ₽/юзер/мес", "Основание"],
           width=[26, 16, 16, 62])
    for i, t in enumerate(TIERS):
        r = 25 + i
        put(ws, f"A{r}", t["name"], box=True)
        put(ws, f"B{r}", 1.00, fmt='0.00', color=BLUE, fill=YELLOW, box=True)
        put(ws, f"C{r}", f"=$B$15*B{r}", fmt=RUB2, box=True)
        if i == 0:
            put(ws, f"D{r}", "ФАКТ-основание: по АСР-сегментам начисление на юзера "
                             "почти не зависит от покупательной силы — индекс 0,68–1,21 "
                             "без монотонного тренда (лист «Данные_бонусы»). "
                             "Точную разбивку даст выгрузка бонусов в разрезе тиров.",
                italic=True, color=GREY, size=9, wrap=True, box=True)
        else:
            put(ws, f"D{r}", "", box=True)

    # --- D. индексы месяцев
    put(ws, "A31", "D. Индекс месяца для эмиссии бонусов", bold=True, size=11)
    put(ws, "A32", "1,00 = эмиссия равна среднему за окно выгрузки. Правится, когда придёт "
                   "разбивка начислений по месяцам.", italic=True, color=GREY, size=9)
    put(ws, "A33", "Месяц", bold=True, color="FFFFFF", fill=HDR_FILL, box=True, align="center")
    put(ws, "A34", "Индекс", bold=True, color="FFFFFF", fill=HDR_FILL, box=True, align="center")
    for i, m in enumerate(MONTHS_2026):
        col = get_column_letter(2 + i)
        put(ws, f"{col}33", m, bold=True, color="FFFFFF", fill=HDR_FILL, box=True, align="center")
        put(ws, f"{col}34", 1.00, fmt='0.00', color=BLUE, fill=YELLOW, box=True)
        ws.column_dimensions[col].width = 10
    return ws


def sheet_bonus(wb, agg):
    ws = wb.create_sheet(SH_BONUS)
    title(ws, "Факт: выгрузка эко-бонусов",
          f"{agg['source_file']} · {agg['period_start']} — {agg['period_end']} · "
          f"серое = факт, не менять")

    put(ws, "A4", "A. Итог по выгрузке", bold=True, size=11)
    header(ws, 5, ["Показатель", "Значение", "Ед.", "Как получено"], width=[34, 18, 16, 60])
    rows = [
        ("Карт участников", agg["cards"], CNT, "шт", "Строк в выгрузке (дублей карт нет)"),
        ("Начислено за окно", agg["bonus_total"], RUB, "₽", "Сумма bonusvalue"),
        ("Месяцев в окне", None, '0" мес"', "мес", f"={q(SH_ASSUM)}!B17"),
        ("Эмиссия в месяц", None, RUB, "₽/мес", "= начислено / месяцев в окне"),
        ("Эмиссия в год", None, RUB, "₽/год", "= эмиссия в месяц × 12"),
        ("Базовый бонус на юзера", None, RUB2, "₽/юзер/мес",
         "= эмиссия в месяц / карт  →  используется как база unit-экономики"),
    ]
    for i, (label, val, fmt, unit, how) in enumerate(rows):
        r = 6 + i
        put(ws, f"A{r}", label, box=True)
        if val is not None:
            put(ws, f"B{r}", val, fmt=fmt, color=BLUE, fill=FACT_FILL, box=True)
        put(ws, f"C{r}", unit, box=True)
        put(ws, f"D{r}", how if not how.startswith("=") else "Ссылка на «Допущения»",
            italic=True, color=GREY, size=9, wrap=True, box=True)
    put(ws, "B8", f"={q(SH_ASSUM)}!B17", fmt='0" мес"', color=GREEN, box=True)
    put(ws, "B9", "=B7/B8", fmt=RUB, box=True)
    put(ws, "B10", "=B9*12", fmt=RUB, box=True)
    put(ws, "B11", "=B9/B6", fmt=RUB2, box=True)

    put(ws, "A14", "B. Распределение начислений на карту за всё окно", bold=True, size=11)
    put(ws, "A15", "Распределение крайне скошенное: медиана 50 ₽ за 14 мес против средних "
                   "224,69 ₽ — четверть карт получила меньше 20 ₽. Это про расходную часть: "
                   "множитель к почти нулевой базе почти ничего не стоит.",
        italic=True, color=GREY, size=9, wrap=True)
    ws.row_dimensions[15].height = 28
    header(ws, 16, ["Квантиль", "₽ за окно", "₽/юзер/мес"], width=[34, 18, 16])
    qs = agg["quantiles_over_period"]
    labels = [("q0", "минимум"), ("q10", "10-й перцентиль"), ("q25", "25-й перцентиль"),
              ("q50", "медиана"), ("q75", "75-й перцентиль"), ("q90", "90-й перцентиль"),
              ("q95", "95-й перцентиль"), ("q99", "99-й перцентиль"), ("q100", "максимум")]
    for i, (key, label) in enumerate(labels):
        r = 17 + i
        put(ws, f"A{r}", label, box=True)
        put(ws, f"B{r}", qs[key], fmt=RUB2, color=BLUE, fill=FACT_FILL, box=True)
        put(ws, f"C{r}", f"=B{r}/{q(SH_ASSUM)}!$B$17", fmt=RUB2, box=True)

    put(ws, "A28", "C. Справочно: АСР-сегменты ВВ — В РАСЧЁТЕ НЕ УЧАСТВУЮТ", bold=True, size=11)
    put(ws, "A29", "Оставлено как доказательство одного факта: начисление эко-бонусов на юзера "
                   "почти не зависит от покупательной силы. У Амбассадора оно НИЖЕ, чем у "
                   "Репертуарного. Поэтому индекс базового бонуса по тирам взят 1,00.",
        italic=True, color=GREY, size=9, wrap=True)
    ws.row_dimensions[29].height = 28
    header(ws, 30, ["Сегмент АСР", "Карт", "Начислено за окно, ₽",
                    "₽/юзер/мес", "Индекс к среднему"], width=[34, 18, 20, 16, 18])
    for i, s in enumerate(agg["segments_reference_only"]):
        r = 31 + i
        put(ws, f"A{r}", s["segment"], box=True)
        put(ws, f"B{r}", s["cards"], fmt=CNT, color=BLUE, fill=FACT_FILL, box=True)
        put(ws, f"C{r}", s["bonus_total"], fmt=RUB, color=BLUE, fill=FACT_FILL, box=True)
        put(ws, f"D{r}", f"=C{r}/B{r}/{q(SH_ASSUM)}!$B$17", fmt=RUB2, box=True)
    total_row = 31 + len(agg["segments_reference_only"])
    for i in range(len(agg["segments_reference_only"])):
        put(ws, f"E{31+i}", f"=IFERROR(D{31+i}/$D${total_row},0)", fmt='0.00', box=True)
    r = total_row
    put(ws, f"A{r}", "ИТОГО", bold=True, box=True)
    put(ws, f"B{r}", f"=SUM(B31:B{r-1})", fmt=CNT, bold=True, box=True)
    put(ws, f"C{r}", f"=SUM(C31:C{r-1})", fmt=RUB, bold=True, box=True)
    put(ws, f"D{r}", f"=C{r}/B{r}/{q(SH_ASSUM)}!$B$17", fmt=RUB2, bold=True, box=True)
    put(ws, f"E{r}", f"=IFERROR(D{r}/$D${r},0)", fmt='0.00', bold=True, box=True)
    return ws


def sheet_revenue(wb, split, total_users, is_placeholder):
    ws = wb.create_sheet(SH_REV)
    title(ws, "Вход: валовая выручка по тирам и месяцам 2026 года",
          "ПЛЕЙСХОЛДЕР — заменить фактической выгрузкой ВВ по пользователям"
          if is_placeholder else "Построено на фактической выгрузке ВВ по пользователям")

    if is_placeholder:
        put(ws, "A3",
            "ВНИМАНИЕ: числа ниже — НЕ ФАКТ. Выгрузка фактической ВВ по пользователям с начала "
            "2026 года запрошена у аналитиков и ещё не получена. Плейсхолдер засеян логнормальным "
            f"распределением (средняя ВВ {2614:,} ₽/мес из агрегатов АСР, log-СКО {PLACEHOLDER_SIGMA}) "
            "только чтобы модель считалась и её можно было проверить. Пересборка при получении файла: "
            "scripts/build_model.py --revenue <файл.xlsx>".replace(",", " "),
            bold=True, fill=YELLOW, wrap=True, size=10)
        ws.row_dimensions[3].height = 44

    labels = ["Месяц", "Статус"]
    labels += [f"Юзеров: {t['name']}" for t in TIERS]
    labels += [f"ВВ/юзер/мес: {t['name']}" for t in TIERS]
    labels += ["Всего юзеров", "ВВ когорты, ₽/мес"]
    header(ws, 5, labels, width=[10, 16] + [13] * 4 + [15] * 4 + [15, 20])

    for i, m in enumerate(MONTHS_2026):
        r = 6 + i
        put(ws, f"A{r}", m, box=True, align="center")
        put(ws, f"B{r}", f'=IF({i+1}<={q(SH_ASSUM)}!$B$20,"в окне выгрузки","прогноз")',
            box=True, align="center", size=9, italic=True, color=GREY)
        for j, t in enumerate(TIERS):
            users = round(split[j]["share"] * total_users)
            put(ws, f"{get_column_letter(3+j)}{r}", users, fmt=CNT,
                color=BLUE, fill=YELLOW, box=True)
            put(ws, f"{get_column_letter(7+j)}{r}", round(split[j]["mean_revenue"]),
                fmt=RUB, color=BLUE, fill=YELLOW, box=True)
        put(ws, f"K{r}", f"=SUM(C{r}:F{r})", fmt=CNT, box=True)
        put(ws, f"L{r}", f"=SUMPRODUCT(C{r}:F{r},G{r}:J{r})", fmt=RUB, box=True)

    r = 18
    put(ws, f"A{r}", "ИТОГО", bold=True, box=True)
    put(ws, f"B{r}", "за 2026 год", bold=True, box=True, size=9, italic=True, color=GREY)
    for j in range(4):
        cl = get_column_letter(3 + j)
        put(ws, f"{cl}{r}", f"=AVERAGE({cl}6:{cl}17)", fmt=CNT, bold=True, box=True,
            note="Среднее число юзеров в тире за месяц (не сумма — это один и тот же пул).")
        cl2 = get_column_letter(7 + j)
        put(ws, f"{cl2}{r}", f"=IFERROR(SUMPRODUCT({cl}6:{cl}17,{cl2}6:{cl2}17)/SUM({cl}6:{cl}17),0)",
            fmt=RUB, bold=True, box=True,
            note="Средневзвешенная по числу юзеров ВВ на юзера за месяц.")
    put(ws, f"K{r}", f"=AVERAGE(K6:K17)", fmt=CNT, bold=True, box=True)
    put(ws, f"L{r}", f"=SUM(L6:L17)", fmt=RUB, bold=True, box=True,
        note="Сумма ВВ когорты за 12 месяцев 2026 года.")

    put(ws, "A20", "Легенда: жёлтые ячейки C6:J17 — единственное, что нужно править вручную "
                   "(или пересобрать скриптом). Всё остальное в книге — формулы от них.",
        italic=True, color=GREY, size=9, wrap=True)
    return ws


def sheet_unit(wb):
    """Ядро модели: unit-экономика — один пользователь, один месяц, по тирам."""
    ws = wb.create_sheet(SH_UNIT)
    title(ws, "Unit-экономика: один пользователь за один месяц",
          "Базовая единица модели. Тиры и периоды на листе «Тир_Период» — это юнит × число юзеров.")

    labels = ["Показатель", "Ед."] + [t["name"] for t in TIERS] + ["Средневзв. по когорте"]
    header(ws, 4, labels, width=[42, 14, 15, 15, 15, 15, 20])

    # Ссылки на средневзвешенные значения из «Выручка_вход» (строка 18).
    rows = [
        # Средневзвешенный множитель: делим на итог юзеров из строки 6, а не на
        # собственную ячейку G5 — иначе циклическая ссылка.
        ("Множитель бонуса тира", MULT,
         lambda i: f"={q(SH_ASSUM)}!$D${6+i}", GREEN,
         "=IFERROR(SUMPRODUCT(C6:F6,C5:F5)/G6,0)"),

        ("Юзеров в тире (среднее за месяц 2026)", CNT,
         lambda i: f"={q(SH_REV)}!{get_column_letter(3+i)}$18", GREEN,
         "=SUM(C6:F6)"),

        ("ДОХОД: валовая выручка", RUB,
         lambda i: f"={q(SH_REV)}!{get_column_letter(7+i)}$18", GREEN,
         "=IFERROR(SUMPRODUCT(C7:F7,C6:F6)/G6,0)"),

        ("ДОХОД: грязная маржа", RUB,
         lambda i: f"={get_column_letter(3+i)}7*{q(SH_ASSUM)}!$B$14", BLACK,
         "=IFERROR(SUMPRODUCT(C8:F8,C6:F6)/G6,0)"),

        ("РАСХОД: эко-бонус базовый (без механики)", RUB2,
         lambda i: f"={q(SH_ASSUM)}!$C${25+i}", GREEN,
         "=IFERROR(SUMPRODUCT(C9:F9,C6:F6)/G6,0)"),

        ("РАСХОД: эко-бонус с механикой", RUB2,
         lambda i: f"={get_column_letter(3+i)}9*{get_column_letter(3+i)}5", BLACK,
         "=IFERROR(SUMPRODUCT(C10:F10,C6:F6)/G6,0)"),

        ("РАСХОД: доплата от механики (Δ бонус)", RUB2,
         lambda i: f"={get_column_letter(3+i)}10-{get_column_letter(3+i)}9", BLACK,
         "=IFERROR(SUMPRODUCT(C11:F11,C6:F6)/G6,0)"),

        ("ВКЛАД: грязная маржа − бонус с механикой", RUB,
         lambda i: f"={get_column_letter(3+i)}8-{get_column_letter(3+i)}10", BLACK,
         "=IFERROR(SUMPRODUCT(C12:F12,C6:F6)/G6,0)"),

        ("Бонус с механикой, % от грязной маржи", PCT2,
         lambda i: f"=IFERROR({get_column_letter(3+i)}10/{get_column_letter(3+i)}8,0)", BLACK,
         "=IFERROR(G10/G8,0)"),

        ("Доплата, % от грязной маржи", PCT2,
         lambda i: f"=IFERROR({get_column_letter(3+i)}11/{get_column_letter(3+i)}8,0)", BLACK,
         "=IFERROR(G11/G8,0)"),

        ("Доплата, % от валовой выручки", PCT2,
         lambda i: f"=IFERROR({get_column_letter(3+i)}11/{get_column_letter(3+i)}7,0)", BLACK,
         "=IFERROR(G11/G7,0)"),
    ]

    units = ["×", "чел", "₽/юзер/мес", "₽/юзер/мес", "₽/юзер/мес", "₽/юзер/мес",
             "₽/юзер/мес", "₽/юзер/мес", "%", "%", "%"]

    for i, (label, fmt, fml, color, total_fml) in enumerate(rows):
        r = 5 + i
        bold = "ВКЛАД" in label
        put(ws, f"A{r}", label, box=True, bold=bold)
        put(ws, f"B{r}", units[i], box=True, size=9, color=GREY)
        for j in range(4):
            put(ws, f"{get_column_letter(3+j)}{r}", fml(j), fmt=fmt, color=color,
                box=True, bold=bold)
        put(ws, f"G{r}", total_fml, fmt=fmt, box=True, bold=True)

    put(ws, "A17", "Как читать", bold=True, size=11)
    put(ws, "A18",
        "Строка «ВКЛАД» — сколько один пользователь тира приносит за месяц после всех бонусных затрат. "
        "Строка «Доплата, % от грязной маржи» — цена механики для этого тира: во что обходится "
        "повышенный коэффициент, если пользователь НЕ изменил поведение. Именно эта доплата, "
        "умноженная на число юзеров, даёт «слепые» затраты на листе «Тир_Период».",
        wrap=True, size=10)
    ws.row_dimensions[18].height = 44
    return ws


def sheet_grid(wb):
    """Разложение по каждому тиру × каждому месяцу 2026 года."""
    ws = wb.create_sheet(SH_GRID)
    title(ws, "Разложение по тирам и месяцам 2026 года: доходы и расходы",
          "Один блок на тир, снизу — ИТОГО по когорте. Столбцы — месяцы, N — итог за 2026 год.")

    # столбцы: A метрика, B..M месяцы, N итог
    ws.column_dimensions["A"].width = 44
    for i in range(12):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13
    ws.column_dimensions["N"].width = 17

    def month_header(row):
        put(ws, f"A{row}", "Показатель", bold=True, color="FFFFFF", fill=HDR_FILL,
            box=True, align="center")
        for i, m in enumerate(MONTHS_2026):
            put(ws, f"{get_column_letter(2+i)}{row}", m, bold=True, color="FFFFFF",
                fill=HDR_FILL, box=True, align="center")
        put(ws, f"N{row}", "2026 итого", bold=True, color="FFFFFF", fill=HDR_FILL,
            box=True, align="center")

    # метрики блока тира: (label, fmt, формула(месяц idx, tier idx), агрегат)
    metrics = [
        ("Юзеров в тире", CNT,
         lambda mi, ti: f"={q(SH_REV)}!{get_column_letter(3+ti)}{6+mi}", "avg"),
        ("Индекс месяца (эмиссия)", '0.00',
         lambda mi, ti: f"={q(SH_ASSUM)}!{get_column_letter(2+mi)}$34", "avg"),
        ("ДОХОД: валовая выручка, ₽", RUB,
         lambda mi, ti: f"={get_column_letter(2+mi)}{{r0}}*{q(SH_REV)}!{get_column_letter(7+ti)}{6+mi}", "sum"),
        ("ДОХОД: грязная маржа, ₽", RUB,
         lambda mi, ti: f"={get_column_letter(2+mi)}{{r2}}*{q(SH_ASSUM)}!$B$14", "sum"),
        ("РАСХОД: эко-бонус базовый, ₽", RUB,
         lambda mi, ti: f"={get_column_letter(2+mi)}{{r0}}*{q(SH_ASSUM)}!$C${25+ti}*{get_column_letter(2+mi)}{{r1}}", "sum"),
        ("РАСХОД: эко-бонус с механикой, ₽", RUB,
         lambda mi, ti: f"={get_column_letter(2+mi)}{{r4}}*{q(SH_ASSUM)}!$D${6+ti}", "sum"),
        ("РАСХОД: доплата от механики, ₽", RUB,
         lambda mi, ti: f"={get_column_letter(2+mi)}{{r5}}-{get_column_letter(2+mi)}{{r4}}", "sum"),
        ("ВКЛАД: маржа − бонус с механикой, ₽", RUB,
         lambda mi, ti: f"={get_column_letter(2+mi)}{{r3}}-{get_column_letter(2+mi)}{{r5}}", "sum"),
        ("Доплата, % от грязной маржи", PCT2,
         lambda mi, ti: f"=IFERROR({get_column_letter(2+mi)}{{r6}}/{get_column_letter(2+mi)}{{r3}},0)", "ratio"),
    ]

    tier_rows = []  # запомним, где лежит каждая метрика каждого тира
    row = 4
    for ti, t in enumerate(TIERS):
        put(ws, f"A{row}", f"Тир {t['name']} — ВВ/мес "
                           f"{'≥ ' + format(t['low'], ',').replace(',', ' ') + ' ₽' if t['high'] is None else format(t['low'], ',').replace(',', ' ') + '–' + format(t['high'], ',').replace(',', ' ') + ' ₽'}",
            bold=True, size=11)
        row += 1
        month_header(row)
        row += 1
        base = row
        rowmap = {f"r{i}": base + i for i in range(len(metrics))}
        for i, (label, fmt, fml, agg_kind) in enumerate(metrics):
            r = base + i
            bold = label.startswith("ВКЛАД")
            put(ws, f"A{r}", label, box=True, bold=bold)
            for mi in range(12):
                put(ws, f"{get_column_letter(2+mi)}{r}", fml(mi, ti).format(**rowmap),
                    fmt=fmt, box=True, bold=bold)
            if agg_kind == "sum":
                put(ws, f"N{r}", f"=SUM(B{r}:M{r})", fmt=fmt, bold=True, box=True)
            elif agg_kind == "avg":
                put(ws, f"N{r}", f"=AVERAGE(B{r}:M{r})", fmt=fmt, bold=True, box=True)
            else:
                put(ws, f"N{r}", f"=IFERROR(N{base+6}/N{base+3},0)", fmt=fmt, bold=True, box=True)
        tier_rows.append(rowmap | {"base": base})
        row = base + len(metrics) + 1

    # --- ИТОГО по когорте
    put(ws, f"A{row}", "ИТОГО по когорте (все тиры)", bold=True, size=11)
    row += 1
    month_header(row)
    row += 1
    # В блоке ИТОГО строки «Индекс месяца» нет, поэтому раскладку считаем явно,
    # а не смещениями от tbase — иначе легко сослаться на соседнюю метрику.
    tbase = row
    totals = [m for m in metrics if m[0] != "Индекс месяца (эмиссия)"]
    total_map = {}
    key_by_label = {
        "Юзеров в тире": "users",
        "ДОХОД: валовая выручка, ₽": "revenue",
        "ДОХОД: грязная маржа, ₽": "margin",
        "РАСХОД: эко-бонус базовый, ₽": "bonus_base",
        "РАСХОД: эко-бонус с механикой, ₽": "bonus_mech",
        "РАСХОД: доплата от механики, ₽": "topup",
        "ВКЛАД: маржа − бонус с механикой, ₽": "contrib",
        "Доплата, % от грязной маржи": "topup_pct",
    }
    for i, (label, _fmt, _fml, _agg) in enumerate(totals):
        total_map[key_by_label[label]] = tbase + i

    for label, fmt, _fml, agg_kind in totals:
        r = total_map[key_by_label[label]]
        bold = label.startswith(("ВКЛАД", "ДОХОД", "РАСХОД"))
        put(ws, f"A{r}", label, box=True, bold=bold)
        # индекс метрики в исходном списке metrics — чтобы взять нужную строку
        # каждого блока тира
        mi_idx = [m[0] for m in metrics].index(label)
        for mi in range(12):
            cl = get_column_letter(2 + mi)
            if agg_kind == "ratio":
                put(ws, f"{cl}{r}",
                    f"=IFERROR({cl}{total_map['topup']}/{cl}{total_map['margin']},0)",
                    fmt=fmt, box=True, bold=bold)
            else:
                parts = "+".join(f"{cl}{tr[f'r{mi_idx}']}" for tr in tier_rows)
                put(ws, f"{cl}{r}", f"={parts}", fmt=fmt, box=True, bold=bold)
        if agg_kind == "sum":
            put(ws, f"N{r}", f"=SUM(B{r}:M{r})", fmt=fmt, bold=True, box=True)
        elif agg_kind == "avg":
            put(ws, f"N{r}", f"=AVERAGE(B{r}:M{r})", fmt=fmt, bold=True, box=True)
        else:
            put(ws, f"N{r}",
                f"=IFERROR(N{total_map['topup']}/N{total_map['margin']},0)",
                fmt=fmt, bold=True, box=True)
    row = tbase + len(totals)

    row += 1
    put(ws, f"A{row}", "Ключевой вывод по расходной части", bold=True, size=11)
    row += 1
    put(ws, f"A{row}",
        "Строка «РАСХОД: доплата от механики» в блоке ИТОГО, столбец N — это «слепая» доплата за 2026 год: "
        "сколько мы платим сверх текущего просто за присвоение повышенных коэффициентов, без изменения поведения. "
        "В предыдущей версии модели она составляла 35,4 млн ₽/год, потому что тир присваивался по АСР-сегменту "
        "и 51 тыс. Амбассадоров попадали в Лес (×2,85), а 85 тыс. Средних — в Дерево (×2,15). "
        "По границам валовой выручки в Лес попадают единицы процентов базы, поэтому доплата на порядок меньше.",
        wrap=True, size=10)
    ws.row_dimensions[row].height = 60
    return ws, total_map


def sheet_migration(wb):
    ws = wb.create_sheet(SH_MIGR)
    title(ws, "Unit-экономика перехода между тирами",
          "Сколько приносит один пользователь, поднявший валовую выручку до следующего тира. "
          "Всё — на одного мигранта в месяц.")

    labels = ["Шаг перехода", "Ед.", "① Старт → Росток", "② Росток → Дерево",
              "③ Дерево → Лес"]
    header(ws, 4, labels, width=[44, 14, 18, 18, 18])

    # исходный тир i -> целевой i+1 ; колонки C,D,E
    rows = [
        ("ВВ на юзера: исходный тир", RUB,
         lambda s: f"={q(SH_UNIT)}!{get_column_letter(3+s)}$7", GREEN),
        ("ВВ на юзера: целевой тир", RUB,
         lambda s: f"={q(SH_UNIT)}!{get_column_letter(4+s)}$7", GREEN),
        ("Δ ВВ (прирост выручки)", RUB,
         lambda s: f"={get_column_letter(3+s)}6-{get_column_letter(3+s)}5", BLACK),
        ("Δ грязная маржа (ДОХОД)", RUB,
         lambda s: f"={get_column_letter(3+s)}7*{q(SH_ASSUM)}!$B$14", BLACK),
        ("Множитель: исходный тир", MULT,
         lambda s: f"={q(SH_ASSUM)}!$D${6+s}", GREEN),
        ("Множитель: целевой тир", MULT,
         lambda s: f"={q(SH_ASSUM)}!$D${7+s}", GREEN),
        ("Δ эко-бонус (РАСХОД)", RUB2,
         lambda s: f"={q(SH_ASSUM)}!$C${25+s}*({get_column_letter(3+s)}10-{get_column_letter(3+s)}9)",
         BLACK),
        ("ЧИСТЫЙ ЭФФЕКТ на мигранта", RUB,
         lambda s: f"={get_column_letter(3+s)}8-{get_column_letter(3+s)}11", BLACK),
        ("Δ бонус, % от Δ маржи", PCT2,
         lambda s: f"=IFERROR({get_column_letter(3+s)}11/{get_column_letter(3+s)}8,0)", BLACK),
        ("Пул-источник: юзеров в исходном тире", CNT,
         lambda s: f"={q(SH_UNIT)}!{get_column_letter(3+s)}$6", GREEN),
    ]
    units = ["₽/юзер/мес", "₽/юзер/мес", "₽/юзер/мес", "₽/юзер/мес", "×", "×",
             "₽/юзер/мес", "₽/юзер/мес", "%", "чел"]

    for i, (label, fmt, fml, color) in enumerate(rows):
        r = 5 + i
        bold = label.startswith("ЧИСТЫЙ")
        put(ws, f"A{r}", label, box=True, bold=bold)
        put(ws, f"B{r}", units[i], box=True, size=9, color=GREY)
        for s in range(3):
            put(ws, f"{get_column_letter(3+s)}{r}", fml(s), fmt=fmt, color=color,
                box=True, bold=bold)

    put(ws, "A17", "Почему переход почти всегда выгоден", bold=True, size=11)
    put(ws, "A18",
        "Δ маржа считается в рублях выручки (тысячи ₽), а Δ бонус — в рублях эко-начисления "
        "(единицы ₽), потому что база эко-бонуса ≈ 16 ₽/юзер/мес. Отсюда «Δ бонус, % от Δ маржи» "
        "— доли процента. Риск механики не в стоимости перехода, а в «слепой» доплате тем, "
        "кто уже в высоком тире и поведение не менял (лист «Тир_Период»).",
        wrap=True, size=10)
    ws.row_dimensions[18].height = 44
    return ws


def sheet_scenarios(wb, total_map):
    ws = wb.create_sheet(SH_SCEN)
    title(ws, "Сценарии конверсии и порог безубыточности",
          "Конверсия — доля пула исходного тира, переходящая в целевой тир за год.")

    header(ws, 4, ["Показатель", "Ед.", "Пессимистичный", "Базовый", "Оптимистичный"],
           width=[46, 14, 18, 18, 18])

    conv = [(0.05, 0.10, 0.20), (0.05, 0.10, 0.20), (0.03, 0.05, 0.10)]
    for s in range(3):
        r = 5 + s
        put(ws, f"A{r}", f"Конверсия шага {'①②③'[s]}", box=True)
        put(ws, f"B{r}", "% пула/год", box=True, size=9, color=GREY)
        for k in range(3):
            put(ws, f"{get_column_letter(3+k)}{r}", conv[s][k], fmt=PCT,
                color=BLUE, fill=YELLOW, box=True)

    # Мигранты набираются в течение года, поэтому в первый год каждый работает
    # в среднем половину года — отсюда отдельная строка «месяцев работы».
    put(ws, "A9", "Мигрантов за год", box=True)
    put(ws, "B9", "чел", box=True, size=9, color=GREY)
    put(ws, "A10", "Месяцев работы мигранта в первый год", box=True)
    put(ws, "B10", "мес", box=True, size=9, color=GREY)
    put(ws, "A11", "ДОХОД: прирост грязной маржи за год", box=True, bold=True)
    put(ws, "B11", "₽/год", box=True, size=9, color=GREY)
    put(ws, "A12", "РАСХОД: бонусы мигрантов за год", box=True, bold=True)
    put(ws, "B12", "₽/год", box=True, size=9, color=GREY)
    put(ws, "A13", "РАСХОД: слепая доплата за год", box=True, bold=True)
    put(ws, "B13", "₽/год", box=True, size=9, color=GREY)
    put(ws, "A14", "ЧИСТЫЙ ЭФФЕКТ за год", box=True, bold=True)
    put(ws, "B14", "₽/год", box=True, size=9, color=GREY)
    put(ws, "A15", "Все бонусные затраты, % от прироста маржи", box=True)
    put(ws, "B15", "%", box=True, size=9, color=GREY)

    M = q(SH_MIGR)
    G = q(SH_GRID)
    for k in range(3):
        c = get_column_letter(3 + k)
        # мигранты = сумма по шагам (пул исходного тира × конверсия шага)
        put(ws, f"{c}9",
            f"={M}!$C$14*{c}5+{M}!$D$14*{c}6+{M}!$E$14*{c}7",
            fmt=CNT, box=True)
        put(ws, f"{c}10", 6, fmt='0', color=BLUE, fill=YELLOW, box=True,
            note="Мигранты набираются в течение года, поэтому в первый год каждый "
                 "работает в среднем половину года. 6 мес — консервативно. "
                 "Поставьте 12 для расчёта на полный год в стационарном режиме.")
        put(ws, f"{c}11",
            f"={c}10*({M}!$C$14*{c}5*{M}!$C$8+{M}!$D$14*{c}6*{M}!$D$8"
            f"+{M}!$E$14*{c}7*{M}!$E$8)", fmt=RUB, box=True, bold=True)
        put(ws, f"{c}12",
            f"={c}10*({M}!$C$14*{c}5*{M}!$C$11+{M}!$D$14*{c}6*{M}!$D$11"
            f"+{M}!$E$14*{c}7*{M}!$E$11)", fmt=RUB, box=True, bold=True)
        put(ws, f"{c}13", f"={G}!$N${total_map['topup']}", fmt=RUB, box=True, bold=True,
            color=GREEN)
        put(ws, f"{c}14", f"={c}11-{c}12-{c}13", fmt=RUB, box=True, bold=True)
        put(ws, f"{c}15", f"=IFERROR(({c}12+{c}13)/{c}11,0)", fmt=PCT2, box=True)

    put(ws, "A18", "Порог безубыточности", bold=True, size=11)
    put(ws, "A19", "Единая конверсия по всем трём шагам, при которой чистый эффект = 0", box=True)
    put(ws, "B19", "% пула/год", box=True, size=9, color=GREY)
    put(ws, "C19",
        f"=IFERROR({G}!$N${total_map['topup']}/"
        f"($C$10*({M}!$C$14*({M}!$C$8-{M}!$C$11)+{M}!$D$14*({M}!$D$8-{M}!$D$11)"
        f"+{M}!$E$14*({M}!$E$8-{M}!$E$11))),0)",
        fmt=PCT2, box=True, bold=True, fill=YELLOW)

    put(ws, "A21", "Запас до порога: базовый сценарий / порог", box=True)
    put(ws, "B21", "×", box=True, size=9, color=GREY)
    put(ws, "C21", "=IFERROR(AVERAGE(D5:D7)/$C$19,0)", fmt='0.0"×"', box=True, bold=True)

    put(ws, "A23", "Итог по когорте за 2026 год (из листа «Тир_Период»)", bold=True, size=11)
    items = [("ДОХОД: валовая выручка", "revenue", RUB),
             ("ДОХОД: грязная маржа", "margin", RUB),
             ("РАСХОД: эко-бонус базовый (как сейчас)", "bonus_base", RUB),
             ("РАСХОД: эко-бонус с механикой", "bonus_mech", RUB),
             ("РАСХОД: доплата от механики (слепая)", "topup", RUB),
             ("ВКЛАД до учёта миграции", "contrib", RUB)]
    for i, (label, key, fmt) in enumerate(items):
        r = 24 + i
        put(ws, f"A{r}", label, box=True, bold=label.startswith("ВКЛАД"))
        put(ws, f"B{r}", "₽/год", box=True, size=9, color=GREY)
        put(ws, f"C{r}", f"={G}!$N${total_map[key]}", fmt=fmt, color=GREEN, box=True,
            bold=label.startswith("ВКЛАД"))

    put(ws, "A31",
        "Порог безубыточности — не прогноз, а свойство конструкции: он показывает, "
        "какая доля пула должна подняться в следующий тир, чтобы окупить слепую доплату. "
        "Реальную конверсию даст только A/B с holdout.",
        italic=True, color=GREY, size=9, wrap=True)
    ws.row_dimensions[31].height = 30
    return ws


# ------------------------------------------------------------------- main

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--bonus-aggregates", default="data/bonus_aggregates.json")
    p.add_argument("--revenue", default=None,
                   help="xlsx с фактической ВВ по пользователям (когда придёт)")
    p.add_argument("-o", "--out", default="output/eco_tiers_unit_economics.xlsx")
    a = p.parse_args()

    with open(a.bonus_aggregates, encoding="utf-8") as fh:
        agg = json.load(fh)

    if a.revenue:
        split, total_users = tier_split_from_revenue(a.revenue, TIERS)
        is_placeholder = False
    else:
        # Средняя ВВ на участника, собранная из агрегатов АСР-выгрузки:
        # (70 459×157 + 84 909×1 212 + 51 329×10 330 + 42 253×157) / 248 950.
        mean_rev = 2614
        split = lognormal_tier_split(mean_rev, PLACEHOLDER_SIGMA, TIERS)
        total_users = agg["cards"]
        is_placeholder = True
        print("  ВВ по пользователям не передана → плейсхолдер "
              f"(логнормаль, средняя {mean_rev} ₽/мес, log-СКО {PLACEHOLDER_SIGMA})")
        for t, s in zip(TIERS, split):
            print(f"    {t['name']:<10} {s['share']*100:6.3f}%  "
                  f"{round(s['share']*total_users):>7} чел  "
                  f"ВВ/юзер/мес {s['mean_revenue']:>8,.0f} ₽")

    wb = Workbook()
    wb.remove(wb.active)

    sheet_readme(wb, agg, is_placeholder)
    sheet_assumptions(wb, agg)
    sheet_bonus(wb, agg)
    sheet_revenue(wb, split, total_users, is_placeholder)
    sheet_unit(wb)
    _grid, total_map = sheet_grid(wb)
    sheet_migration(wb)
    sheet_scenarios(wb, total_map)

    wb.save(a.out)
    print(f"  → {a.out}")


if __name__ == "__main__":
    main()
