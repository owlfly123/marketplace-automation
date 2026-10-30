#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генерирует вымышленные данные в формате отчётов WB, чтобы планировщик можно
было запустить без реальных файлов. Все товары и цифры придуманы.

    python generate_demo_data.py            # создаст папку demo_data
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)
OUT = Path(__file__).parent / "demo_data"

TITLES = [
    "Сад на подоконнике", "Как перестать откладывать", "Тайны северного маяка",
    "Математика для мечтателей", "Кухня без отходов", "Дневник путешественника",
    "Сказки старого леса", "Привычки, которые работают", "Город под звёздами",
    "Учимся считать с котом", "Короткая история времени года", "Кофе и философия",
    "Велосипедом через Европу", "Тихий час: медитации", "Живопись за 30 дней",
    "Язык растений", "Финансы для начинающих", "Большая книга головоломок",
    "Письма с острова", "Как говорить с подростком", "Атлас неба",
    "Овощи круглый год", "Детектив в библиотеке", "Мир без пластика",
    "Рисуем динозавров", "Энциклопедия вкусов", "Сон и продуктивность",
    "Пёс по имени Вторник", "Архитектура для детей", "Слова, которые лечат",
]
# Регионы в том виде, как они встречаются в отчётах (включая два варианта написания Южного)
REGIONS = ["Южный + Северо-Кавказский", "Южный и Северо-Кавказский",
           "Приволжский", "Северо-Западный", "Уральский", "Центральный"]


def main():
    OUT.mkdir(exist_ok=True)
    skus = [(str(30000 + i * 37), t) for i, t in enumerate(TITLES)]
    base = {(s, r): random.choice([0, 1, 2, 5, 10, 20, 40]) for s, _ in skus for r in REGIONS}

    # Особые случаи, чтобы показать логику
    spike_sku, growth_sku = skus[0][0], skus[1][0]
    base[(spike_sku, "Приволжский")] = 20
    base[(growth_sku, "Уральский")] = 15

    week_start = date(2026, 9, 28)
    for wk in range(4):
        start = week_start + timedelta(weeks=wk)
        end = start + timedelta(days=6)
        rows = []
        for sku, title in skus:
            for region in REGIONS:
                orders = max(0, int(random.gauss(base[(sku, region)], base[(sku, region)] * 0.25)))
                if sku == spike_sku and region == "Приволжский" and wk == 2:
                    orders = 140  # краткосрочный всплеск
                if sku == growth_sku and region == "Уральский":
                    orders = [10, 14, 19, 26][wk]  # устойчивый рост
                if region.startswith("Южный") and wk % 2:
                    # в разные недели WB по-разному пишет название региона
                    region_name = "Южный и Северо-Кавказский"
                else:
                    region_name = region
                if orders == 0 and random.random() < 0.5:
                    continue
                rows.append({
                    "Артикул продавца": sku, "Название": title, "Регион": region_name,
                    "Итого заказов, шт": orders,
                    "Остатки склад ВБ, шт": random.randint(0, 3) * base[(sku, region)],
                })
        name = f"sales_report_{start:%d.%m}-{end:%d.%m.%Y}.xlsx"
        with pd.ExcelWriter(OUT / name) as xw:
            pd.DataFrame({"Отчёт": ["Демо-данные"]}).to_excel(xw, sheet_name="Сводка", index=False)
            pd.DataFrame(rows).to_excel(xw, sheet_name="Детальные данные", index=False)

    # Файл «В пути»: листы названы по-разному, Уральского листа нет вовсе
    sheets = {"Yuzhny": "Южный", "Privolzhsky": "Privolzhsky", "Severo-Zapadny": "Северо-Западный"}
    with pd.ExcelWriter(OUT / f"in_way_{week_start + timedelta(weeks=4):%d.%m}.xlsx") as xw:
        for code, sheet in sheets.items():
            picked = random.sample(skus, 6)
            pd.DataFrame({
                "Поставка": [f"П-{random.randint(1000, 9999)}" for _ in picked],
                "Склад": [code] * len(picked),
                "ID1": ["" for _ in picked],
                "ID2": [s for s, _ in picked],
                "Название": [t for _, t in picked],
                "Дата": ["" for _ in picked],
                "Количество": [random.choice([10, 20, 30]) for _ in picked],
            }).to_excel(xw, sheet_name=sheet, index=False)

    print(f"Демо-данные созданы в {OUT}")


if __name__ == "__main__":
    main()
