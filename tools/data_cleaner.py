"""Tool 08: Data Cleaner — fix common data quality issues in Excel/CSV files."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd


@dataclass
class CleanResult:
    changes_made: dict[str, int]
    output_path: str
    change_log: list[str]


_NULL_VALUES = {"", "null", "n/a", "na", "none", "nil", "-", "nan", "#n/a"}

_DATE_FORMATS = [
    "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d",
    "%Y-%m-%d", "%d %b %Y", "%d %B %Y", "%b %d, %Y",
]


def _try_parse_date(val: str) -> Optional[str]:
    from datetime import datetime
    val = val.strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(val, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def _clean_number(val: str) -> str:
    cleaned = re.sub(r"[₹$€£,\s]", "", str(val))
    try:
        float(cleaned)
        return cleaned
    except ValueError:
        return val


def _clean_phone(val: str) -> str:
    digits = re.sub(r"[\s\-\(\)\+]", "", str(val))
    if digits.startswith("91") and len(digits) == 12:
        digits = digits[2:]
    return digits


def _load(file_path: str) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(file_path, dtype=str, keep_default_na=False, na_filter=False)
    return pd.read_excel(file_path, dtype=str, keep_default_na=False, na_values=[])


def _save(df: pd.DataFrame, path: str) -> str:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    if p.suffix.lower() == ".csv":
        df.to_csv(path, index=False)
    else:
        df.to_excel(path, index=False)
    return str(p.resolve())


_DEFAULT_CONFIG = {
    "strip_whitespace": True,
    "normalize_case": None,
    "fix_dates": [],
    "clean_numbers": [],
    "fix_phone": [],
    "empty_to_null": True,
}


def clean_file(
    file_path: str,
    config: Optional[dict] = None,
    output_path: Optional[str] = None,
) -> CleanResult:
    """Clean common data quality issues in Excel/CSV files.

    Config keys (all optional, defaults shown):
        strip_whitespace (bool): Trim leading/trailing spaces in all string cols.
        normalize_case (dict): {col: 'title'|'upper'|'lower'} per column.
        fix_dates (list[str]): Column names to normalise to YYYY-MM-DD.
        clean_numbers (list[str]): Column names to strip currency symbols/commas.
        fix_phone (list[str]): Column names with Indian phone numbers to standardise.
        empty_to_null (bool): Replace "", "NULL", "N/A", "none", "-" with "".

    Args:
        file_path: Source file path.
        config: Dict of operation toggles (see above).
        output_path: Destination. Defaults to <stem>_cleaned.<ext>.

    Returns:
        CleanResult with per-operation change counts, output path, and change log.
    """
    df = _load(file_path)
    cfg = {**_DEFAULT_CONFIG, **(config or {})}
    changes: dict[str, int] = {}
    log: list[str] = []

    if cfg.get("strip_whitespace"):
        count = 0
        for col in df.select_dtypes(include="object").columns:
            before = df[col].copy()
            df[col] = df[col].str.strip()
            diff = (before != df[col]).sum()
            count += int(diff)
        changes["strip_whitespace"] = count
        if count:
            log.append(f"strip_whitespace: {count} cells trimmed")

    if cfg.get("normalize_case") and isinstance(cfg["normalize_case"], dict):
        count = 0
        for col, mode in cfg["normalize_case"].items():
            if col not in df.columns:
                continue
            before = df[col].copy()
            if mode == "title":
                df[col] = df[col].str.title()
            elif mode == "upper":
                df[col] = df[col].str.upper()
            elif mode == "lower":
                df[col] = df[col].str.lower()
            diff = (before != df[col]).sum()
            count += int(diff)
        changes["normalize_case"] = count
        if count:
            log.append(f"normalize_case: {count} cells changed")

    if cfg.get("fix_dates"):
        count = 0
        for col in cfg["fix_dates"]:
            if col not in df.columns:
                continue
            for idx, val in df[col].items():
                if pd.isna(val) or str(val).strip() == "":
                    continue
                parsed = _try_parse_date(str(val))
                if parsed and parsed != str(val).strip():
                    df.at[idx, col] = parsed
                    count += 1
        changes["fix_dates"] = count
        if count:
            log.append(f"fix_dates: {count} dates normalised")

    if cfg.get("clean_numbers"):
        count = 0
        for col in cfg["clean_numbers"]:
            if col not in df.columns:
                continue
            for idx, val in df[col].items():
                if pd.isna(val):
                    continue
                cleaned = _clean_number(str(val))
                if cleaned != str(val):
                    df.at[idx, col] = cleaned
                    count += 1
        changes["clean_numbers"] = count
        if count:
            log.append(f"clean_numbers: {count} values cleaned")

    if cfg.get("fix_phone"):
        count = 0
        for col in cfg["fix_phone"]:
            if col not in df.columns:
                continue
            for idx, val in df[col].items():
                if pd.isna(val):
                    continue
                cleaned = _clean_phone(str(val))
                if cleaned != str(val):
                    df.at[idx, col] = cleaned
                    count += 1
        changes["fix_phone"] = count
        if count:
            log.append(f"fix_phone: {count} phones standardised")

    if cfg.get("empty_to_null"):
        count = 0
        for col in df.select_dtypes(include="object").columns:
            mask = df[col].str.lower().str.strip().isin(_NULL_VALUES)
            count += int(mask.sum())
            df.loc[mask, col] = ""
        changes["empty_to_null"] = count
        if count:
            log.append(f"empty_to_null: {count} cells normalised")

    src = Path(file_path)
    if output_path is None:
        output_path = str(src.parent / f"{src.stem}_cleaned{src.suffix}")

    saved = _save(df, output_path)
    return CleanResult(changes_made=changes, output_path=saved, change_log=log)
