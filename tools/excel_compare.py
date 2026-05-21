"""Tool 01: Excel Compare — compare two Excel/CSV files row-by-row."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd


@dataclass
class ComparisonResult:
    added: pd.DataFrame
    removed: pd.DataFrame
    modified: pd.DataFrame
    unchanged: pd.DataFrame
    summary: dict = field(default_factory=dict)

    def __post_init__(self):
        self.summary = {
            "added": len(self.added),
            "removed": len(self.removed),
            "modified": len(self.modified),
            "unchanged": len(self.unchanged),
        }

    def export_report(self, output_path: str) -> str:
        """Write multi-sheet Excel diff report.

        Args:
            output_path: Destination .xlsx path.

        Returns:
            Absolute path to the written file.
        """
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary_df = pd.DataFrame([self.summary])
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            self.added.to_excel(writer, sheet_name="Added", index=False)
            self.removed.to_excel(writer, sheet_name="Removed", index=False)
            self.modified.to_excel(writer, sheet_name="Modified", index=False)
            self.unchanged.to_excel(writer, sheet_name="Unchanged", index=False)
        return str(Path(output_path).resolve())


def _load(file_path: str) -> pd.DataFrame:
    """Load CSV or Excel file into a DataFrame."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(file_path, dtype=str).fillna("")
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(file_path, dtype=str).fillna("")
    raise ValueError(f"Unsupported file format: {suffix}")


def _detect_key_columns(df1: pd.DataFrame, df2: pd.DataFrame) -> list[str]:
    """Return columns common to both frames whose values are unique in df1."""
    common = [c for c in df1.columns if c in df2.columns]
    for col in common:
        if df1[col].is_unique:
            return [col]
    return common[:1] if common else []


def compare_files(
    file1_path: str,
    file2_path: str,
    key_columns: Optional[list[str]] = None,
    auto_key: bool = True,
) -> ComparisonResult:
    """Compare two Excel/CSV files and return a structured diff.

    Args:
        file1_path: Path to the baseline file.
        file2_path: Path to the comparison file.
        key_columns: Column(s) used as the unique row identifier.
        auto_key: When True and key_columns is None, attempt to detect keys.

    Returns:
        ComparisonResult with added/removed/modified/unchanged DataFrames.

    Raises:
        FileNotFoundError: If either file does not exist.
        ValueError: If key columns are not found in both files.
    """
    df1 = _load(file1_path)
    df2 = _load(file2_path)

    if df1.empty and df2.empty:
        empty = pd.DataFrame()
        return ComparisonResult(empty, empty, empty, empty)

    if key_columns is None:
        if auto_key:
            key_columns = _detect_key_columns(df1, df2)
        else:
            key_columns = []

    if key_columns:
        missing1 = [c for c in key_columns if c not in df1.columns]
        missing2 = [c for c in key_columns if c not in df2.columns]
        if missing1 or missing2:
            raise ValueError(
                f"Key columns missing — file1: {missing1}, file2: {missing2}"
            )

    if not key_columns:
        return _compare_without_keys(df1, df2)

    return _compare_with_keys(df1, df2, key_columns)


def _compare_without_keys(df1: pd.DataFrame, df2: pd.DataFrame) -> ComparisonResult:
    """Row-order comparison when no key columns are available."""
    common_cols = [c for c in df1.columns if c in df2.columns]
    df1 = df1[common_cols].reset_index(drop=True)
    df2 = df2[common_cols].reset_index(drop=True)

    max_len = max(len(df1), len(df2))
    added_rows, removed_rows, modified_rows, unchanged_rows = [], [], [], []

    for i in range(max_len):
        if i >= len(df1):
            added_rows.append(df2.iloc[i])
        elif i >= len(df2):
            removed_rows.append(df1.iloc[i])
        elif df1.iloc[i].equals(df2.iloc[i]):
            unchanged_rows.append(df1.iloc[i])
        else:
            row = df2.iloc[i].copy()
            modified_rows.append(row)

    def to_df(rows, ref):
        if not rows:
            return pd.DataFrame(columns=ref.columns)
        return pd.DataFrame(rows, columns=ref.columns).reset_index(drop=True)

    return ComparisonResult(
        added=to_df(added_rows, df2),
        removed=to_df(removed_rows, df1),
        modified=to_df(modified_rows, df2),
        unchanged=to_df(unchanged_rows, df1),
    )


def _compare_with_keys(
    df1: pd.DataFrame, df2: pd.DataFrame, key_columns: list[str]
) -> ComparisonResult:
    """Key-based comparison — handles reordered rows correctly."""
    all_cols = list(dict.fromkeys(list(df1.columns) + list(df2.columns)))
    df1 = df1.reindex(columns=all_cols, fill_value="")
    df2 = df2.reindex(columns=all_cols, fill_value="")

    df1 = df1.set_index(key_columns)
    df2 = df2.set_index(key_columns)

    keys1 = set(df1.index)
    keys2 = set(df2.index)

    added_keys = keys2 - keys1
    removed_keys = keys1 - keys2
    common_keys = keys1 & keys2

    added = df2.loc[list(added_keys)].reset_index() if added_keys else df2.iloc[0:0].reset_index()
    removed = df1.loc[list(removed_keys)].reset_index() if removed_keys else df1.iloc[0:0].reset_index()

    modified_rows = []
    unchanged_rows = []
    for key in common_keys:
        r1 = df1.loc[key]
        r2 = df2.loc[key]
        if isinstance(r1, pd.DataFrame):
            r1 = r1.iloc[0]
        if isinstance(r2, pd.DataFrame):
            r2 = r2.iloc[0]
        if r1.equals(r2):
            unchanged_rows.append(r2)
        else:
            row = r2.copy()
            modified_rows.append(row)

    def rows_to_df(rows, ref_df):
        if not rows:
            return ref_df.iloc[0:0].reset_index()
        return pd.DataFrame(rows, columns=ref_df.columns).reset_index()

    modified = rows_to_df(modified_rows, df2)
    unchanged = rows_to_df(unchanged_rows, df1)

    return ComparisonResult(
        added=added,
        removed=removed,
        modified=modified,
        unchanged=unchanged,
    )
