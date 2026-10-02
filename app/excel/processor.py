"""Calculations, consolidation and analysis. No file or GUI code lives here."""
from __future__ import annotations

import pandas as pd

from app.utils import constants as C

PRODUCT_COLUMNS = ["Product", "Category", "Transactions", "Total Quantity",
                   "Total Revenue", "Average Unit Price"]
SALES_COLUMNS = ["Salesperson", "Transactions", "Quantity Sold",
                 "Total Revenue", "Average Order Value"]
MONTH_COLUMNS = ["Month", "Transactions", "Quantity", "Revenue",
                 "Discount", "Tax", "Average Order Value"]


def recalculate(df: pd.DataFrame) -> pd.DataFrame:
    """Apply the same formulas the input workbooks use.

    Input files written by scripts have no cached formula results, so the
    engine always recomputes instead of trusting the stored numbers.
    """
    df[C.COL_GROSS] = df[C.COL_QUANTITY] * df[C.COL_UNIT_PRICE]
    df[C.COL_DISCOUNT_AMT] = df[C.COL_GROSS] * df[C.COL_DISCOUNT]
    df[C.COL_TAX_AMT] = (df[C.COL_GROSS] - df[C.COL_DISCOUNT_AMT]) * df[C.COL_TAX_RATE]
    df[C.COL_NET] = df[C.COL_GROSS] - df[C.COL_DISCOUNT_AMT] + df[C.COL_TAX_AMT]
    return df


def to_typed(df: pd.DataFrame) -> pd.DataFrame:
    """Keep the required columns, enforce types, recompute calculated columns."""
    out = df[C.REQUIRED_COLUMNS].copy()
    out[C.COL_DATE] = pd.to_datetime(out[C.COL_DATE], format="mixed").dt.normalize()
    for col in C.NUMERIC_COLUMNS:
        out[col] = pd.to_numeric(out[col])
    for col in C.TEXT_COLUMNS:
        out[col] = out[col].astype(str).str.strip()
    return recalculate(out)


def month_labels(dates: pd.Series) -> pd.Series:
    """'January' for single-year data, 'January 2027' when years are mixed."""
    names = dates.dt.month.map(lambda m: C.MONTH_NAMES[m - 1])
    if dates.dt.year.nunique() > 1:
        return names + " " + dates.dt.year.astype(str)
    return names


def consolidate(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """Combine typed frames (each already has a Source File column)."""
    combined = pd.concat(frames, ignore_index=True)
    combined = combined.sort_values([C.COL_DATE, C.COL_INVOICE], kind="stable")
    combined = combined.reset_index(drop=True)
    combined[C.COL_MONTH] = month_labels(combined[C.COL_DATE])
    return combined[C.REQUIRED_COLUMNS + [C.COL_SOURCE, C.COL_MONTH]]


def product_analysis(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby([C.COL_PRODUCT, C.COL_CATEGORY], as_index=False).agg(**{
        "Transactions": (C.COL_INVOICE, "count"),
        "Total Quantity": (C.COL_QUANTITY, "sum"),
        "Total Revenue": (C.COL_NET, "sum"),
        "Average Unit Price": (C.COL_UNIT_PRICE, "mean"),
    })
    g = g.rename(columns={C.COL_PRODUCT: "Product", C.COL_CATEGORY: "Category"})
    return g.sort_values(["Total Revenue", "Product"], ascending=[False, True])[PRODUCT_COLUMNS] \
            .reset_index(drop=True)


def salesperson_analysis(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(C.COL_SALESPERSON, as_index=False).agg(**{
        "Transactions": (C.COL_INVOICE, "count"),
        "Quantity Sold": (C.COL_QUANTITY, "sum"),
        "Total Revenue": (C.COL_NET, "sum"),
    })
    g = g.rename(columns={C.COL_SALESPERSON: "Salesperson"})
    g["Average Order Value"] = g["Total Revenue"] / g["Transactions"]
    return g.sort_values(["Total Revenue", "Salesperson"], ascending=[False, True])[SALES_COLUMNS] \
            .reset_index(drop=True)


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    # df is sorted by date, so sort=False keeps chronological month order.
    g = df.groupby(C.COL_MONTH, sort=False, as_index=False).agg(**{
        "Transactions": (C.COL_INVOICE, "count"),
        "Quantity": (C.COL_QUANTITY, "sum"),
        "Revenue": (C.COL_NET, "sum"),
        "Discount": (C.COL_DISCOUNT_AMT, "sum"),
        "Tax": (C.COL_TAX_AMT, "sum"),
    })
    g = g.rename(columns={C.COL_MONTH: "Month"})
    g["Average Order Value"] = g["Revenue"] / g["Transactions"]
    return g[MONTH_COLUMNS]


def dashboard_metrics(df: pd.DataFrame) -> dict[str, float | int]:
    """Headline numbers. 'Revenue' throughout = Net Sales (after discount, incl. tax)."""
    n = len(df)
    revenue = float(df[C.COL_NET].sum())
    return {
        "Total Revenue": revenue,
        "Total Quantity": int(df[C.COL_QUANTITY].sum()),
        "Total Tax": float(df[C.COL_TAX_AMT].sum()),
        "Total Discount": float(df[C.COL_DISCOUNT_AMT].sum()),
        "Average Order Value": revenue / n if n else 0.0,
        "Number of Transactions": n,
        "Number of Customers": int(df[C.COL_CUSTOMER].nunique()),
        "Number of Products": int(df[C.COL_PRODUCT].nunique()),
        "Number of Salespersons": int(df[C.COL_SALESPERSON].nunique()),
    }