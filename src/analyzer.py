"""Compute summary statistics from the merged sales DataFrame."""

import pandas as pd


def summarize_by_category(df: pd.DataFrame) -> pd.DataFrame:
    """Group by category: sum revenue/units, count transactions."""
    result = (
        df.groupby("category", as_index=False)
        .agg(
            revenue=("revenue", "sum"),
            units=("units", "sum"),
            transactions=("revenue", "count"),
        )
        .sort_values("revenue", ascending=False)
        .reset_index(drop=True)
    )
    return result


def monthly_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Group by _source file, sum revenue/units, rename source to 'month'."""
    result = (
        df.groupby("_source", as_index=False)
        .agg(
            revenue=("revenue", "sum"),
            units=("units", "sum"),
        )
        .rename(columns={"_source": "month"})
    )
    # Strip extension for cleaner display
    result["month"] = result["month"].str.replace(r"\.[^.]+$", "", regex=True)
    return result


def top_products(df: pd.DataFrame, n: int = 5) -> pd.DataFrame:
    """Group by product, sum revenue, return top n sorted descending."""
    result = (
        df.groupby("product", as_index=False)
        .agg(
            revenue=("revenue", "sum"),
            units=("units", "sum"),
        )
        .sort_values("revenue", ascending=False)
        .head(n)
        .reset_index(drop=True)
    )
    return result


def kpi_summary(df: pd.DataFrame) -> dict:
    """Return a dict of high-level KPIs."""
    total_revenue = df["revenue"].sum()
    total_units = df["units"].sum()
    avg_order_value = total_revenue / len(df) if len(df) > 0 else 0
    top_product = (
        df.groupby("product")["revenue"].sum().idxmax() if not df.empty else "N/A"
    )
    top_region = (
        df.groupby("region")["revenue"].sum().idxmax() if not df.empty else "N/A"
    )
    num_months = df["_source"].nunique() if "_source" in df.columns else 0
    return {
        "total_revenue": round(total_revenue, 2),
        "total_units": int(total_units),
        "avg_order_value": round(avg_order_value, 2),
        "top_product": top_product,
        "top_region": top_region,
        "num_months": num_months,
    }
