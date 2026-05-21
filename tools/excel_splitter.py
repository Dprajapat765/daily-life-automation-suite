"""Tool 06: Excel Splitter — split one file by unique values in a column."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class SplitResult:
    files_created: list[str]
    rows_per_file: dict[str, int]
    output_dir: str


def _safe_filename(value: str) -> str:
    """Convert a cell value to a safe filename stem."""
    safe = re.sub(r'[\\/*?:"<>|]', "_", str(value))
    safe = safe.strip().replace(" ", "_")
    return safe or "unknown"


def _load(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(file_path, dtype=str).fillna("")
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(file_path, dtype=str).fillna("")
    raise ValueError(f"Unsupported file format: {path.suffix}")


def split_by_column(
    file_path: str,
    split_column: str,
    output_dir: str,
    output_format: str = "xlsx",
) -> SplitResult:
    """Split one Excel/CSV into multiple files by unique values in a column.

    Args:
        file_path: Source file path (.xlsx or .csv).
        split_column: Column whose unique values determine the split.
        output_dir: Directory where output files are written.
        output_format: Output file format — 'xlsx' or 'csv'.

    Returns:
        SplitResult with created file paths and row counts per file.

    Raises:
        FileNotFoundError: If source file does not exist.
        ValueError: If split_column is not found or output_format is invalid.
    """
    if output_format not in {"xlsx", "csv"}:
        raise ValueError(f"output_format must be 'xlsx' or 'csv', got '{output_format}'")

    df = _load(file_path)

    if split_column not in df.columns:
        raise ValueError(
            f"Column '{split_column}' not found. Available: {list(df.columns)}"
        )

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    files_created: list[str] = []
    rows_per_file: dict[str, int] = {}

    for value, group in df.groupby(split_column, sort=False):
        stem = _safe_filename(str(value))
        out_name = f"{stem}.{output_format}"
        out_path = Path(output_dir) / out_name

        if output_format == "csv":
            group.to_csv(out_path, index=False)
        else:
            group.to_excel(out_path, index=False)

        files_created.append(str(out_path.resolve()))
        rows_per_file[out_name] = len(group)

    return SplitResult(
        files_created=files_created,
        rows_per_file=rows_per_file,
        output_dir=str(Path(output_dir).resolve()),
    )
