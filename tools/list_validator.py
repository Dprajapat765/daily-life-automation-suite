"""Tool 03: List Validator — cross-check two lists better than VLOOKUP."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd
from rapidfuzz import fuzz, process


@dataclass
class ValidationResult:
    exact_matches: list[dict]
    near_matches: list[dict]
    missing: list[str]
    match_rate_pct: float
    source_col: str
    target_col: str

    def export_report(self, output_path: str) -> str:
        """Write colour-coded Excel report.

        Args:
            output_path: Destination .xlsx path.

        Returns:
            Absolute path to the written file.
        """
        from openpyxl.styles import PatternFill
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        exact_df = pd.DataFrame(self.exact_matches) if self.exact_matches else pd.DataFrame(columns=["Source", "Target", "Match_Type", "Score"])
        near_df = pd.DataFrame(self.near_matches) if self.near_matches else pd.DataFrame(columns=["Source", "Target", "Match_Type", "Score"])
        missing_df = pd.DataFrame({"Missing_Value": self.missing}) if self.missing else pd.DataFrame(columns=["Missing_Value"])
        summary_df = pd.DataFrame([{
            "Total_Source_Values": len(self.exact_matches) + len(self.near_matches) + len(self.missing),
            "Exact_Matches": len(self.exact_matches),
            "Near_Matches": len(self.near_matches),
            "Missing": len(self.missing),
            "Match_Rate_Pct": round(self.match_rate_pct, 2),
        }])

        green = "C6EFCE"
        yellow = "FFEB9C"
        red = "FFC7CE"

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            exact_df.to_excel(writer, sheet_name="Exact_Matches", index=False)
            near_df.to_excel(writer, sheet_name="Near_Matches", index=False)
            missing_df.to_excel(writer, sheet_name="Missing", index=False)

            for sheet, color, df in [
                ("Exact_Matches", green, exact_df),
                ("Near_Matches", yellow, near_df),
                ("Missing", red, missing_df),
            ]:
                ws = writer.sheets[sheet]
                fill = PatternFill("solid", fgColor=color)
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                    for cell in row:
                        cell.fill = fill

        return str(Path(output_path).resolve())


def _load_column(file_path: str, col_name: str) -> list[str]:
    """Load a single column from CSV or Excel as a list of strings."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(file_path, dtype=str).fillna("")
    else:
        df = pd.read_excel(file_path, dtype=str).fillna("")
    if col_name not in df.columns:
        raise ValueError(f"Column '{col_name}' not found in {file_path}. Available: {list(df.columns)}")
    return df[col_name].tolist()


def validate_list(
    source_path: str,
    target_path: str,
    source_col: str,
    target_col: str,
    fuzzy_threshold: int = 85,
) -> ValidationResult:
    """Cross-check source list values against target list.

    Each source value is classified as:
    - EXACT MATCH: found verbatim in target
    - NEAR MATCH: fuzzy score >= fuzzy_threshold
    - MISSING: not found at all

    Args:
        source_path: Path to the file containing the source list.
        target_path: Path to the file containing the target/lookup list.
        source_col: Column name in source file.
        target_col: Column name in target file.
        fuzzy_threshold: Minimum score (0-100) to count as a near match.

    Returns:
        ValidationResult with exact_matches, near_matches, missing, and match_rate_pct.
    """
    source_values = _load_column(source_path, source_col)
    target_values = _load_column(target_path, target_col)

    target_set = {v.strip().lower(): v for v in target_values}
    target_list = list(target_set.keys())

    exact: list[dict] = []
    near: list[dict] = []
    missing: list[str] = []

    for src in source_values:
        src_clean = src.strip()
        src_lower = src_clean.lower()

        if src_lower in target_set:
            exact.append({
                "Source": src_clean,
                "Target": target_set[src_lower],
                "Match_Type": "EXACT",
                "Score": 100,
            })
            continue

        if not target_list:
            missing.append(src_clean)
            continue

        best = process.extractOne(
            src_lower,
            target_list,
            scorer=fuzz.token_sort_ratio,
        )

        if best and best[1] >= fuzzy_threshold:
            near.append({
                "Source": src_clean,
                "Target": target_set[best[0]],
                "Match_Type": "NEAR MATCH",
                "Score": round(best[1], 1),
            })
        else:
            missing.append(src_clean)

    total = len(source_values)
    matched = len(exact) + len(near)
    match_rate = (matched / total * 100) if total > 0 else 0.0

    return ValidationResult(
        exact_matches=exact,
        near_matches=near,
        missing=missing,
        match_rate_pct=match_rate,
        source_col=source_col,
        target_col=target_col,
    )
