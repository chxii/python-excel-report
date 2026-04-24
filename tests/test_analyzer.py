"""Unit tests for analyzer.py"""

import pytest
import pandas as pd

from src.analyzer import summarize_by_category, monthly_totals, top_products, kpi_summary


@pytest.fixture
def sample_df():
    return pd.DataFrame({
        "product": ["Keyboard", "Mouse", "Chair", "Desk", "Hub"],
        "category": ["Electronics", "Electronics", "Furniture", "Furniture", "Electronics"],
        "region": ["North", "East", "West", "North", "South"],
        "revenue": [1000, 500, 2000, 1500, 800],
        "units": [10, 5, 2, 1, 8],
        "_source": ["jan.csv", "jan.csv", "feb.csv", "feb.csv", "jan.csv"],
    })


def test_summarize_by_category(sample_df):
    result = summarize_by_category(sample_df)
    assert "category" in result.columns
    assert "revenue" in result.columns
    assert "units" in result.columns
    assert "transactions" in result.columns
    # Electronics: 1000+500+800=2300, Furniture: 2000+1500=3500
    furn = result[result["category"] == "Furniture"].iloc[0]
    assert furn["revenue"] == 3500
    elec = result[result["category"] == "Electronics"].iloc[0]
    assert elec["transactions"] == 3


def test_monthly_totals(sample_df):
    result = monthly_totals(sample_df)
    assert "month" in result.columns
    assert "revenue" in result.columns
    assert len(result) == 2
    jan = result[result["month"] == "jan"].iloc[0]
    assert jan["revenue"] == 1000 + 500 + 800


def test_top_products(sample_df):
    result = top_products(sample_df, n=3)
    assert len(result) == 3
    assert result.iloc[0]["product"] == "Chair"  # highest revenue 2000


def test_top_products_less_than_n(sample_df):
    result = top_products(sample_df, n=10)
    assert len(result) == 5  # only 5 products exist


def test_kpi_summary(sample_df):
    kpis = kpi_summary(sample_df)
    assert kpis["total_revenue"] == 5800
    assert kpis["total_units"] == 26
    assert kpis["num_months"] == 2
    assert kpis["top_product"] == "Chair"
    assert kpis["top_region"] == "North"  # North: 1000+1500=2500
    assert kpis["avg_order_value"] == round(5800 / 5, 2)
