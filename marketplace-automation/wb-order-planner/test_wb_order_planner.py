# Запуск: python -m pytest
import math

import pandas as pd

from wb_order_planner import Params, abc_classes, demand

P = Params()


def test_spike_is_excluded():
    # Неделя 140 больше чем в 3 раза выше среднего остальных (18) — исключаем её
    base, april, k, comments, spike, growth = demand([20, 23, 140, 11], P)
    assert spike and not growth
    assert math.isclose(base, 18 / 7)
    assert math.ceil(april * P.horizon) == 87


def test_steady_growth_is_capped():
    # 10 < 14 < 19 < 26: sqrt(26/10) = 1.61 → ограничиваем 1.6
    base, april, k, comments, spike, growth = demand([10, 14, 19, 26], P)
    assert growth and not spike
    assert k == 1.6
    assert math.ceil(april * P.horizon) == 133


def test_zero_sales_is_not_a_spike():
    *_, spike, growth = demand([0, 0, 0, 0], P)
    assert not spike and not growth


def test_abc():
    totals = pd.Series({"a": 70, "b": 15, "c": 10, "d": 5})
    assert abc_classes(totals, P).to_dict() == {"a": "A", "b": "A", "c": "B", "d": "C"}
