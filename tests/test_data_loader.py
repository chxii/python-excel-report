"""Unit tests for data_loader.py"""

import pytest
import pandas as pd
from pathlib import Path
import tempfile
import os

from src.data_loader import load_file, load_folder


@pytest.fixture
def sample_csv(tmp_path):
    path = tmp_path / "test.csv"
    path.write_text("product,category,region,revenue,units\nKeyboard,Electronics,North,1000,10\n")
    return path


@pytest.fixture
def sample_folder(tmp_path):
    (tmp_path / "a.csv").write_text("product,category,region,revenue,units\nMouse,Electronics,East,500,5\n")
    (tmp_path / "b.csv").write_text("product,category,region,revenue,units\nChair,Furniture,West,2000,2\n")
    return tmp_path


def test_load_file_csv(sample_csv):
    df = load_file(sample_csv)
    assert isinstance(df, pd.DataFrame)
    assert list(df.columns) == ["product", "category", "region", "revenue", "units"]
    assert len(df) == 1
    assert df.iloc[0]["product"] == "Keyboard"


def test_load_file_unsupported(tmp_path):
    bad = tmp_path / "file.txt"
    bad.write_text("hello")
    with pytest.raises(ValueError, match="Unsupported file type"):
        load_file(bad)


def test_load_folder_source_column(sample_folder):
    df = load_folder(sample_folder)
    assert "_source" in df.columns
    assert len(df) == 2
    assert set(df["_source"]) == {"a.csv", "b.csv"}


def test_load_folder_concat(sample_folder):
    df = load_folder(sample_folder)
    assert len(df) == 2
    products = set(df["product"])
    assert "Mouse" in products
    assert "Chair" in products


def test_load_folder_empty(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_folder(tmp_path)
