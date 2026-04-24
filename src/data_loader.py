"""Load CSV and Excel files into pandas DataFrames."""

from pathlib import Path
import pandas as pd


def load_file(path: str | Path) -> pd.DataFrame:
    """Load a single CSV or xlsx file. Raises ValueError for unsupported types."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    elif suffix in (".xlsx", ".xls"):
        return pd.read_excel(path)
    else:
        raise ValueError(f"Unsupported file type: {suffix!r}. Expected .csv or .xlsx/.xls")


def load_folder(folder: str | Path) -> pd.DataFrame:
    """Load all CSV/xlsx files in a folder, adding a '_source' column with the filename."""
    folder = Path(folder)
    frames = []
    for file in sorted(folder.iterdir()):
        if file.suffix.lower() in (".csv", ".xlsx", ".xls"):
            df = load_file(file)
            df["_source"] = file.name
            frames.append(df)
    if not frames:
        raise FileNotFoundError(f"No CSV or Excel files found in: {folder}")
    return pd.concat(frames, ignore_index=True)
