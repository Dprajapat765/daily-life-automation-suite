"""Tool 07: Deduplicator — find and handle duplicate rows in a file."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Optional

import pandas as pd


@dataclass
class DedupResult:
    total_rows: int
    duplicate_count: int
    unique_count: int
    output_path: Optional[str]
    duplicates_path: Optional[str]


def _load(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(file_path, dtype=str).fillna("")
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(file_path, dtype=str).fillna("")
    raise ValueError(f"Unsupported format: {path.suffix}")


def _save(df: pd.DataFrame, path: str) -> str:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    if p.suffix.lower() == ".csv":
        df.to_csv(path, index=False)
    else:
        df.to_excel(path, index=False)
    return str(p.resolve())


def find_duplicates(
    file_path: str,
    key_columns: Optional[list[str]] = None,
    action: Literal["flag", "remove_keep_first", "remove_keep_last"] = "flag",
    output_path: Optional[str] = None,
) -> DedupResult:
    """Find and handle duplicate rows in a single file.

    Args:
        file_path: Source file path (.xlsx or .csv).
        key_columns: Columns to check for duplicates. None = all columns.
        action: What to do with duplicates:
            'flag'              — add Duplicate_Status column, keep all rows
            'remove_keep_first' — keep first occurrence, drop the rest
            'remove_keep_last'  — keep last occurrence, drop the rest
        output_path: Destination for the cleaned file. None = derived from source.

    Returns:
        DedupResult with counts and output paths.

    Raises:
        FileNotFoundError: If the source file does not exist.
        ValueError: If key_columns contain columns not in the file, or action is invalid.
    """
    valid_actions = {"flag", "remove_keep_first", "remove_keep_last"}
    if action not in valid_actions:
        raise ValueError(f"action must be one of {valid_actions}")

    df = _load(file_path)

    if key_columns:
        missing = [c for c in key_columns if c not in df.columns]
        if missing:
            raise ValueError(f"Key columns not found: {missing}. Available: {list(df.columns)}")
        subset = key_columns
    else:
        subset = None

    is_dup = df.duplicated(subset=subset, keep=False)
    dup_count = int(is_dup.sum())
    unique_count = len(df) - dup_count

    src = Path(file_path)
    if output_path is None:
        output_path = str(src.parent / f"{src.stem}_deduped{src.suffix}")
    duplicates_path = str(src.parent / f"{src.stem}_duplicates{src.suffix}")

    if action == "flag":
        result_df = df.copy()
        result_df["Duplicate_Status"] = is_dup.map({True: "DUPLICATE", False: "UNIQUE"})
        dup_df = df[is_dup].copy()
    elif action == "remove_keep_first":
        keep_mask = ~df.duplicated(subset=subset, keep="first")
        result_df = df[keep_mask].reset_index(drop=True)
        dup_df = df[~keep_mask].reset_index(drop=True)
    else:
        keep_mask = ~df.duplicated(subset=subset, keep="last")
        result_df = df[keep_mask].reset_index(drop=True)
        dup_df = df[~keep_mask].reset_index(drop=True)

    saved_out = _save(result_df, output_path)
    saved_dup = _save(dup_df, duplicates_path) if not dup_df.empty else None

    return DedupResult(
        total_rows=len(df),
        duplicate_count=dup_count,
        unique_count=unique_count,
        output_path=saved_out,
        duplicates_path=saved_dup,
    )
