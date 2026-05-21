"""Tool 04: Fuzzy Matcher — compare two columns for text similarity."""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd
from rapidfuzz import fuzz


def _load(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(file_path, dtype=str).fillna("")
    return pd.read_excel(file_path, dtype=str).fillna("")


def _score(a: str, b: str) -> float:
    """Return best of ratio and token_sort_ratio."""
    return max(fuzz.ratio(a, b), fuzz.token_sort_ratio(a, b))


def _classify(score: float, threshold: int) -> str:
    if score == 100:
        return "EXACT"
    if score >= threshold:
        return "SIMILAR"
    return "NO MATCH"


def _build_result(
    values_a: list[str],
    values_b: list[str],
    threshold: int,
    col_a_name: str = "Value_A",
    col_b_name: str = "Value_B",
) -> pd.DataFrame:
    """Align two value lists row-by-row and compute match info."""
    n = max(len(values_a), len(values_b))
    rows = []
    for i in range(n):
        a = values_a[i] if i < len(values_a) else ""
        b = values_b[i] if i < len(values_b) else ""
        sc = round(_score(a, b), 1)
        rows.append({
            "Value_A": a,
            "Value_B": b,
            "Match_Type": _classify(sc, threshold),
            "Score_Pct": sc,
        })
    df = pd.DataFrame(rows, columns=["Value_A", "Value_B", "Match_Type", "Score_Pct"])
    df = df.rename(columns={"Value_A": col_a_name, "Value_B": col_b_name})
    return df


def fuzzy_match(
    file_path: str,
    col_a: str,
    col_b: str,
    threshold: int = 80,
) -> pd.DataFrame:
    """Compare two columns within the same file.

    Args:
        file_path: Path to Excel/CSV file.
        col_a: Name of the first column.
        col_b: Name of the second column.
        threshold: Minimum score to classify as SIMILAR (0-100).

    Returns:
        DataFrame with columns: Value_A, Value_B, Match_Type, Score_Pct.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If a specified column is not found.
    """
    df = _load(file_path)
    for col in (col_a, col_b):
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")
    return _build_result(df[col_a].tolist(), df[col_b].tolist(), threshold, col_a, col_b)


def fuzzy_match_files(
    file1: str,
    col1: str,
    file2: str,
    col2: str,
    threshold: int = 80,
) -> pd.DataFrame:
    """Compare one column from each of two files row-by-row.

    Args:
        file1: Path to the first file.
        col1: Column name in file1.
        file2: Path to the second file.
        col2: Column name in file2.
        threshold: Minimum score to classify as SIMILAR (0-100).

    Returns:
        DataFrame with columns: Value_A, Value_B, Match_Type, Score_Pct.

    Raises:
        FileNotFoundError: If either file does not exist.
        ValueError: If a specified column is not found.
    """
    df1 = _load(file1)
    df2 = _load(file2)
    for col, df, path in [(col1, df1, file1), (col2, df2, file2)]:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in '{path}'. Available: {list(df.columns)}")
    return _build_result(df1[col1].tolist(), df2[col2].tolist(), threshold, col1, col2)
