"""Tool 09: Format Validator — validate column values against format rules."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd


_VALIDATORS: dict[str, re.Pattern] = {
    "email": re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$"),
    "phone_india": re.compile(r"^(\+91)?[6-9]\d{9}$"),
    "pincode": re.compile(r"^\d{6}$"),
    "pan": re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$"),
    "gst": re.compile(r"^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$"),
    "positive_number": re.compile(r"^[+]?(\d+(\.\d*)?|\.\d+)$"),
    "non_empty": re.compile(r".+"),
    "url": re.compile(r"^https?://[^\s/$.?#].[^\s]*$"),
}

_DATE_PATTERNS = [
    re.compile(r"^\d{4}-\d{2}-\d{2}$"),
    re.compile(r"^\d{2}/\d{2}/\d{4}$"),
    re.compile(r"^\d{2}-\d{2}-\d{4}$"),
    re.compile(r"^\d{2}/\d{2}/\d{4}$"),
]


def _validate_value(value: str, rule: str) -> str:
    """Return 'VALID', 'INVALID', or 'EMPTY'."""
    if value is None or pd.isna(value) or str(value).strip() == "":
        return "EMPTY"
    v = str(value).strip()
    if rule == "date":
        for pat in _DATE_PATTERNS:
            if pat.match(v):
                return "VALID"
        return "INVALID"
    if rule == "positive_number":
        try:
            num = float(v)
            return "VALID" if num > 0 else "INVALID"
        except ValueError:
            return "INVALID"
    if rule == "phone_india":
        digits_only = re.sub(r"[\s\-]", "", v)
        if _VALIDATORS["phone_india"].match(digits_only):
            return "VALID"
        return "INVALID"
    pat = _VALIDATORS.get(rule)
    if pat is None:
        raise ValueError(f"Unknown validator: '{rule}'. Available: {list(_VALIDATORS.keys()) + ['date']}")
    return "VALID" if pat.match(v) else "INVALID"


@dataclass
class ValidationReport:
    per_column_stats: dict[str, dict]
    invalid_rows: list[dict]
    valid_rate_pct: float
    output_path: Optional[str]

    def export_report(self, output_path: str) -> str:
        from openpyxl.styles import PatternFill
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        stats_df = pd.DataFrame([
            {"Column": col, **stats}
            for col, stats in self.per_column_stats.items()
        ])
        invalid_df = pd.DataFrame(self.invalid_rows) if self.invalid_rows else pd.DataFrame()
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            stats_df.to_excel(writer, sheet_name="Summary", index=False)
            invalid_df.to_excel(writer, sheet_name="Invalid_Rows", index=False)
        return str(Path(output_path).resolve())


def validate_formats(
    file_path: str,
    column_rules: dict[str, str],
    output_path: Optional[str] = None,
) -> ValidationReport:
    """Validate column values against format rules.

    Args:
        file_path: Source Excel/CSV path.
        column_rules: Mapping of {column_name: validator_name}.
            Supported: email, phone_india, date, pincode, pan, gst,
                       positive_number, non_empty, url
        output_path: Optional path for the annotated output Excel.

    Returns:
        ValidationReport with per-column stats, invalid rows, and valid rate.

    Raises:
        FileNotFoundError: If source file does not exist.
        ValueError: If a column is not found, or validator name is unknown.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(file_path, dtype=str, keep_default_na=False, na_filter=False)
    else:
        df = pd.read_excel(file_path, dtype=str, keep_default_na=False, na_values=[])

    for col in column_rules:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")

    result_df = df.copy()
    per_column_stats: dict[str, dict] = {}
    total_valid = total_cells = 0

    for col, rule in column_rules.items():
        statuses = []
        for val in df[col]:
            status = _validate_value(val, rule)
            statuses.append(status)

        status_series = pd.Series(statuses, index=df.index)
        result_df[f"{col}_Status"] = status_series

        valid_c = int((status_series == "VALID").sum())
        invalid_c = int((status_series == "INVALID").sum())
        empty_c = int((status_series == "EMPTY").sum())
        total = len(statuses)

        per_column_stats[col] = {
            "rule": rule,
            "valid": valid_c,
            "invalid": invalid_c,
            "empty": empty_c,
            "total": total,
            "valid_pct": round(valid_c / total * 100, 1) if total else 0.0,
        }
        total_valid += valid_c
        total_cells += total

    invalid_rows: list[dict] = []
    if column_rules:
        status_cols = [f"{col}_Status" for col in column_rules]
        invalid_mask = result_df[status_cols].isin(["INVALID"]).any(axis=1)
        invalid_rows = result_df[invalid_mask].to_dict("records")

    valid_rate = (total_valid / total_cells * 100) if total_cells > 0 else 0.0

    saved_path = None
    if output_path:
        result_df.to_excel(output_path, index=False) if output_path.endswith(".xlsx") else result_df.to_csv(output_path, index=False)
        saved_path = str(Path(output_path).resolve())

    return ValidationReport(
        per_column_stats=per_column_stats,
        invalid_rows=invalid_rows,
        valid_rate_pct=round(valid_rate, 2),
        output_path=saved_path,
    )
