"""Unit tests for app.excel.processor (no file I/O)."""
from __future__ import annotations

import pandas as pd
import pytest

from app.excel import processor as P
from app.utils import constants as C


def make_frame() -> pd.DataFrame:
    """Three rows: one per month, two products, two salespersons."""
    return pd.DataFrame({
        C.COL_DATE:        pd.to_datetime(["2026-01-10", "2026-02-15", "2026-02-20"]),
        C.COL_INVOICE:     ["INV-1", "INV-2", "INV-3"],
        C.COL_CUSTOMER:    ["Acme", "Beta", "Acme"],
        C.COL_PRODUCT:     ["Widget", "Gadget", "Widget"],
        C.COL_CATEGORY:    ["Electronics", "Furniture", "Electronics"],
        C.COL_QUANTITY:    [10, 4, 6],
        C.COL_UNIT_PRICE:  [100.0, 50.0, 100.0],
        C.COL_DISCOUNT:    [0.10, 0.0, 0.0],
        C.COL_TAX_RATE:    [0.16, 0.16, 0.16],
        C.COL_SALESPERSON: ["Alice", "Bob", "Alice"],
        C.COL_GROSS:       [0.0, 0.0, 0.0],   # will be recomputed
        C.COL_DISCOUNT_AMT:[0.0, 0.0, 0.0],
        C.COL_TAX_AMT:     [0.0, 0.0, 0.0],
        C.COL_NET:         [0.0, 0.0, 0.0],
    })


def test_recalculate_matches_expected_formulas():
    df = P.recalculate(make_frame())
    # Row 0: gross=1000, disc=100, tax=(1000-100)*0.16=144, net=1044
    assert df.loc[0, C.COL_GROSS] == pytest.approx(1000.0)
    assert df.loc[0, C.COL_DISCOUNT_AMT] == pytest.approx(100.0)
    assert df.loc[0, C.COL_TAX_AMT] == pytest.approx(144.0)
    assert df.loc[0, C.COL_NET] == pytest.approx(1044.0)
    # Row 1: gross=200, disc=0, tax=32, net=232
    assert df.loc[1, C.COL_NET] == pytest.approx(232.0)


def test_consolidate_adds_month_column_and_sorts():
    df = P.recalculate(make_frame())
    df[C.COL_SOURCE] = "f.xlsx"
    out = P.consolidate([df])
    assert list(out[C.COL_MONTH]) == ["January", "February", "February"]
    assert out[C.COL_DATE].is_monotonic_increasing


def test_product_analysis_totals():
    df = P.recalculate(make_frame())
    out = P.product_analysis(df).set_index("Product")
    # Widget appears twice: transactions=2, qty=16, revenue=1044+696=1740
    assert out.loc["Widget", "Transactions"] == 2
    assert out.loc["Widget", "Total Quantity"] == 16
    assert out.loc["Widget", "Total Revenue"] == pytest.approx(1740.0)


def test_salesperson_analysis_average_order_value():
    df = P.recalculate(make_frame())
    out = P.salesperson_analysis(df).set_index("Salesperson")
    # Alice has 2 transactions, revenue 1044 + 696 = 1740
    assert out.loc["Alice", "Transactions"] == 2
    assert out.loc["Alice", "Total Revenue"] == pytest.approx(1740.0)
    assert out.loc["Alice", "Average Order Value"] == pytest.approx(870.0)


def test_monthly_summary_uses_chronological_order():
    df = P.recalculate(make_frame())
    df[C.COL_SOURCE] = "f.xlsx"
    df = P.consolidate([df])
    out = P.monthly_summary(df)
    assert list(out["Month"]) == ["January", "February"]
    assert out.loc[out["Month"] == "February", "Transactions"].iloc[0] == 2


def test_dashboard_metrics_match_raw_sums():
    df = P.recalculate(make_frame())
    m = P.dashboard_metrics(df)
    assert m["Number of Transactions"] == 3
    assert m["Number of Customers"] == 2
    assert m["Number of Products"] == 2
    assert m["Number of Salespersons"] == 2
    assert m["Total Revenue"] == pytest.approx(
        df[C.COL_NET].sum())