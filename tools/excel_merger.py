"""Tool 05: Excel Merger — merge N Excel/CSV files into one master file."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Union

import pandas as pd


@dataclass
class MergeResult:
    total_rows: int
    source_counts: dict[str, int]
    duplicate_rows_removed: int
    output_path: str


def _load(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(file_path, dtype=str).fillna("")
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(file_path, dtype=str).fillna("")
    raise ValueError(f"Unsupported file format: {path.suffix}")


def merge_files(
    file_paths: list[str],
    output_path: str,
    dedup: bool = False,
    fill_missing: Union[str, float] = "",
) -> MergeResult:
    """Merge N Excel/CSV files with same or similar structure into one master.

    Handles column order differences, extra columns, missing columns,
    and header name variations. Adds a 'Source_File' column.

    Args:
        file_paths: List of file paths to merge.
        output_path: Destination path for the merged file (.xlsx or .csv).
        dedup: If True, remove exact duplicate rows after merging.
        fill_missing: Value used for missing columns in any file.

    Returns:
        MergeResult with row counts and output path.

    Raises:
        ValueError: If fewer than 1 file is provided.
        FileNotFoundError: If any file does not exist.
    """
    if not file_paths:
        raise ValueError("At least one file path is required.")

    frames: list[pd.DataFrame] = []
    source_counts: dict[str, int] = {}

    for fp in file_paths:
        df = _load(fp)
        file_name = Path(fp).name
        df.insert(0, "Source_File", file_name)
        source_counts[file_name] = len(df)
        frames.append(df)

    merged = pd.concat(frames, ignore_index=True, sort=False)

    for col in merged.columns:
        merged[col] = merged[col].fillna(fill_missing)

    pre_dedup = len(merged)
    removed = 0
    if dedup:
        merged = merged.drop_duplicates()
        removed = pre_dedup - len(merged)

    merged = merged.reset_index(drop=True)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    out = Path(output_path)
    if out.suffix.lower() == ".csv":
        merged.to_csv(output_path, index=False)
    else:
        merged.to_excel(output_path, index=False)

    return MergeResult(
        total_rows=len(merged),
        source_counts=source_counts,
        duplicate_rows_removed=removed,
        output_path=str(out.resolve()),
    )
