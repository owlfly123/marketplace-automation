#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WB Order Planner — план отгрузки на региональные склады Wildberries.

Берёт 4 последних недельных отчёта продаж и файл «В пути», делает ABC-анализ
по каждому региону, считает спрос с учётом всплесков и устойчивого роста
и рекомендует количество к поставке. Результат — Excel с вкладкой на регион.

Запуск:
    python wb_order_planner.py --input demo_data
    python wb_order_planner.py --input "C:\\Work" --cover-days 30 --lead-days 12
"""

import argparse
import math
import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# Настройки
# ─────────────────────────────────────────────────────────────────────────────

# Варианты написания региона в отчётах → код вкладки
REGION_MAP = {
    "Южный + Северо-Кавказский": "Yuzhny",
    "Южный и Северо-Кавказский": "Yuzhny",
    "Южный": "Yuzhny",
    "Приволжский": "Privolzhsky",
    "Северо-Западный": "Severo-Zapadny",
    "Уральский": "Uralsky",
}
REGIONS = ["Yuzhny", "Privolzhsky", "Severo-Zapadny", "Uralsky"]

SALES_SHEET = "Детальные данные"
COL_SKU = "Артикул продавца"
COL_NAME = "Название"
COL_REGION = "Регион"
COL_ORDERS = "Итого заказов, шт"
COL_STOCK = "Остатки склад ВБ, шт"

INWAY_SKU_COL = "ID2"          # обычно 4-я колонка
INWAY_QTY_COL = "Количество"   # обычно 7-я колонка
INWAY_SKU_POS, INWAY_QTY_POS = 3, 6

FILL_YELLOW = PatternFill("solid", fgColor="FFF2CC")
FILL_GREEN = PatternFill("solid", fgColor="E2EFDA")

DATE_RE = re.compile(r"(\d{2})\.(\d{2})(?:\.(\d{2,4}))?")


@dataclass
class Params:
    cover_days: int = 30        # запас на складе, дней
    lead_days: int = 12         # логистическое плечо, дней
    season_factor: float = 0.8  # сезонная поправка к базовому спросу
    spike_ratio: float = 3.0    # во сколько раз неделя должна превышать среднее остальных
    growth_cap: float = 1.6     # максимум коэффициента роста
    abc_a: float = 0.80
    abc_b: float = 0.95

    @property
    def horizon(self) -> int:
        return self.cover_days + self.lead_days


# ─────────────────────────────────────────────────────────────────────────────
# Поиск и чтение файлов
# ─────────────────────────────────────────────────────────────────────────────

def dates_in_name(name: str, default_year: int):
    """Все даты вида ДД.ММ или ДД.ММ.ГГГГ из имени файла."""
    out = []
    for d, m, y in DATE_RE.findall(name):
        year = int(y) if y else default_year
        if year < 100:
            year += 2000
        try:
            out.append(date(year, int(m), int(d)))
        except ValueError:
            pass
    return out


def is_inway_file(path: Path) -> bool:
    return "в пути" in path.stem.lower() or "in_way" in path.stem.lower()


def find_files(folder: Path, today: date):
    """Возвращает (4 последних недельных отчёта по возрастанию дат, файл «В пути»)."""
    weekly, inway = [], []
    for p in folder.glob("*.xlsx"):
        if p.name.startswith("~$") or p.name.startswith("WB_order_plan"):
            continue
        ds = dates_in_name(p.stem, today.year)
        if not ds:
            continue
        if is_inway_file(p):
            inway.append((max(ds), p))
            continue
        try:
            sheets = pd.ExcelFile(p).sheet_names
        except Exception:
            continue
        if SALES_SHEET in sheets:
            weekly.append((min(ds), max(ds), p))

    weekly.sort()
    if len(weekly) < 4:
        raise SystemExit(f"Нужно минимум 4 недельных отчёта с листом «{SALES_SHEET}», найдено: {len(weekly)}")
    inway_file = max(inway)[1] if inway else None
    return weekly[-4:], inway_file


def week_label(start: date, end: date) -> str:
    return f"Заказы {start.day:02d}-{end.day:02d}, шт"


def read_week(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SALES_SHEET, dtype={COL_SKU: str})
    df.columns = [str(c).strip().rstrip(".") for c in df.columns]
    df["region"] = df[COL_REGION].astype(str).str.strip().map(REGION_MAP)
    df = df[df["region"].notna()].copy()
    df[COL_SKU] = df[COL_SKU].astype(str).str.strip()
    for c in (COL_ORDERS, COL_STOCK):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


def sheet_region(sheet_name: str):
    s = sheet_name.strip()
    if s in REGIONS:
        return s
    for ru, code in REGION_MAP.items():
        if s.lower() == ru.lower():
            return code
    return None


def read_inway(path):
    """{(регион, SKU): количество}. Если файла или листа нет — в пути = 0."""
    result = {}
    if path is None:
        return result
    for sheet, df in pd.read_excel(path, sheet_name=None, dtype=str).items():
        region = sheet_region(sheet)
        if region is None or df.empty:
            continue
        sku = df[INWAY_SKU_COL] if INWAY_SKU_COL in df else df.iloc[:, INWAY_SKU_POS]
        qty = df[INWAY_QTY_COL] if INWAY_QTY_COL in df else df.iloc[:, INWAY_QTY_POS]
        qty = pd.to_numeric(qty, errors="coerce").fillna(0)
        for s, q in zip(sku.astype(str).str.strip(), qty):
            if s and s != "nan":
                result[(region, s)] = result.get((region, s), 0) + q
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Расчёт
# ─────────────────────────────────────────────────────────────────────────────

def abc_classes(totals: pd.Series, p: Params) -> pd.Series:
    """ABC по кумулятивной доле заказов: A до 80%, B до 95%, C — остальное."""
    s = totals.sort_values(ascending=False)
    grand = s.sum()
    if grand <= 0:
        return pd.Series("C", index=totals.index)
    share_before = (s.cumsum() - s) / grand  # доля ДО этого SKU: первый SKU всегда A
    cls = pd.Series("C", index=s.index)
    cls[share_before < p.abc_b] = "B"
    cls[share_before < p.abc_a] = "A"
    return cls.reindex(totals.index)


def demand(weeks, p: Params):
    """Возвращает (base_daily, april_daily, k, комментарии, spike, growth)."""
    w = list(weeks)
    comments, spike, growth, k = [], False, False, 1.0

    used = w
    m = max(w)
    i = w.index(m)
    rest = w[:i] + w[i + 1:]
    avg_rest = sum(rest) / len(rest)
    if m > 0 and m >= p.spike_ratio * avg_rest:
        spike = True
        used = rest
        comments.append("Исключен краткосрочный всплеск спроса")

    if w[0] > 0 and w[0] < w[1] < w[2] < w[3]:
        growth = True
        k = min(p.growth_cap, math.sqrt(w[3] / w[0]))
        comments.append("Устойчивый недельный рост спроса" if not spike else "устойчивый рост")

    base_daily = (sum(used) / len(used)) / 7
    april_daily = base_daily * p.season_factor * k
    return base_daily, april_daily, k, comments, spike, growth


def build_plan(weeks_data, inway, p: Params):
    """weeks_data: список из 4 DataFrame (старая → новая неделя)."""
    plans = {}
    for region in REGIONS:
        per_week = []
        names = {}
        for df in weeks_data:
            part = df[df["region"] == region]
            per_week.append(part.groupby(COL_SKU)[COL_ORDERS].sum())
            names.update(part.groupby(COL_SKU)[COL_NAME].last().to_dict())
        last = weeks_data[-1]
        stock = last[last["region"] == region].groupby(COL_SKU)[COL_STOCK].sum()

        table = pd.concat(per_week, axis=1).fillna(0)
        table.columns = ["w1", "w2", "w3", "w4"]
        table["total"] = table.sum(axis=1)
        table["abc"] = abc_classes(table["total"], p)

        rows = []
        for sku, r in table[table["abc"] == "A"].iterrows():
            ws = [r.w1, r.w2, r.w3, r.w4]
            base, april, k, comments, spike, growth = demand(ws, p)
            st = float(stock.get(sku, 0))
            iw = float(inway.get((region, sku), 0))
            need = math.ceil(max(0.0, april * p.horizon - st - iw))
            rows.append({
                "sku": sku, "name": names.get(sku, ""), "stock": st, "inway": iw,
                "total": r.total, "need": need, "base": base, "april": april,
                "weeks": ws, "abc": "A", "k": k, "comment": "; ".join(comments),
                "spike": spike, "growth": growth,
            })
        rows.sort(key=lambda x: x["need"], reverse=True)
        plans[region] = rows
    return plans


# ─────────────────────────────────────────────────────────────────────────────
# Выгрузка в Excel
# ─────────────────────────────────────────────────────────────────────────────

def as_int(x):
    return int(x) if float(x).is_integer() else x


def save_plan(plans, labels, out_path: Path):
    headers = ["Артикул продавца", "Название", "Остатки склад ВБ, шт.", "В пути, шт",
               "Итого заказов, шт", "Рекомендованное количество к поставке, шт.",
               "Base daily demand", "April daily demand", *labels, "ABC",
               "Коэффициент роста (k)", "Комментарий"]
    widths = [14, 45, 12, 10, 12, 18, 12, 12, 12, 12, 12, 12, 6, 12, 40]

    wb = Workbook()
    wb.remove(wb.active)
    for region in REGIONS:
        ws = wb.create_sheet(region)
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        for r in plans[region]:
            ws.append([r["sku"], r["name"], as_int(r["stock"]), as_int(r["inway"]),
                       as_int(r["total"]), r["need"], round(r["base"], 2), round(r["april"], 2),
                       *[as_int(x) for x in r["weeks"]], r["abc"], round(r["k"], 2), r["comment"]])
            fill = FILL_YELLOW if r["spike"] else FILL_GREEN if r["growth"] else None
            if fill:
                for cell in ws[ws.max_row]:
                    cell.fill = fill
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "C2"
    wb.save(out_path)


# ─────────────────────────────────────────────────────────────────────────────
# Проверки качества
# ─────────────────────────────────────────────────────────────────────────────

def sanity_check(plans, inway, p: Params):
    msgs = []
    for region in REGIONS:
        rows = plans[region]
        if not rows:
            msgs.append(f"⚠️  {region}: нет SKU группы A")
            continue
        if not any(r["stock"] for r in rows):
            msgs.append(f"⚠️  {region}: все остатки нулевые — проверьте последний отчёт")
        has_inway = any(k[0] == region for k in inway)
        if has_inway and not any(r["inway"] for r in rows):
            msgs.append(f"⚠️  {region}: в файле «В пути» есть данные, но ни один SKU группы A не совпал")
        for r in rows[:2]:  # ручная сверка двух верхних SKU
            check = math.ceil(max(0.0, r["april"] * p.horizon - r["stock"] - r["inway"]))
            status = "OK" if check == r["need"] else "ОШИБКА"
            msgs.append(f"   {region} / {r['sku']}: недели {[as_int(w) for w in r['weeks']]}, "
                        f"остаток {as_int(r['stock'])}, в пути {as_int(r['inway'])}, "
                        f"к поставке {r['need']} — {status}")
    return msgs


# ─────────────────────────────────────────────────────────────────────────────
# Точка входа
# ─────────────────────────────────────────────────────────────────────────────

def main(argv=None):
    ap = argparse.ArgumentParser(description="План отгрузки WB по регионам")
    ap.add_argument("--input", default=".", help="папка с недельными отчётами и файлом «В пути»")
    ap.add_argument("--cover-days", type=int, default=30)
    ap.add_argument("--lead-days", type=int, default=12)
    ap.add_argument("--season-factor", type=float, default=0.8)
    args = ap.parse_args(argv)

    p = Params(cover_days=args.cover_days, lead_days=args.lead_days, season_factor=args.season_factor)
    folder = Path(args.input)
    today = date.today()

    weekly, inway_file = find_files(folder, today)
    print("Недельные отчёты:")
    for start, end, path in weekly:
        print(f"   {start:%d.%m}–{end:%d.%m}: {path.name}")
    print(f"Файл «В пути»: {inway_file.name if inway_file else 'не найден (в пути = 0)'}")

    weeks_data = [read_week(path) for _, _, path in weekly]
    labels = [week_label(s, e) for s, e, _ in weekly]
    inway = read_inway(inway_file)

    plans = build_plan(weeks_data, inway, p)

    print("\nПроверка:")
    for m in sanity_check(plans, inway, p):
        print(m)

    out = folder / f"WB_order_plan_{weekly[-1][1]:%d_%m}_with_in_way.xlsx"
    save_plan(plans, labels, out)
    print(f"\n✅ Сохранено: {out}")
    for region in REGIONS:
        rows = plans[region]
        print(f"   {region}: {len(rows)} SKU, к поставке {sum(r['need'] for r in rows)} шт.")


if __name__ == "__main__":
    sys.exit(main())
