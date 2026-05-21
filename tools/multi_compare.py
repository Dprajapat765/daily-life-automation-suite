"""Tool 02: Multi-File Compare — compare N Excel/CSV files with a shared key."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd

from tools.excel_compare import _load, _detect_key_columns


@dataclass
class MultiComparisonResult:
    """Holds the results of comparing N files."""
    per_file_summary: dict[str, dict]
    timeline: pd.DataFrame
    overall_diff: pd.DataFrame
    file_names: list[str]

    def export_report(self, output_path: str) -> str:
        """Write multi-sheet Excel report.

        Args:
            output_path: Destination .xlsx path.

        Returns:
            Absolute path to the written file.
        """
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        summary_rows = [
            {"File": fname, **stats}
            for fname, stats in self.per_file_summary.items()
        ]
        summary_df = pd.DataFrame(summary_rows)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            self.timeline.to_excel(writer, sheet_name="Timeline", index=False)
            self.overall_diff.to_excel(writer, sheet_name="Overall_Diff", index=False)

        return str(Path(output_path).resolve())


def compare_multiple(
    file_paths: list[str],
    key_columns: Optional[list[str]] = None,
    auto_key: bool = True,
) -> MultiComparisonResult:
    """Compare N Excel/CSV files using key-based diffing.

    Args:
        file_paths: Ordered list of file paths (baseline is first).
        key_columns: Column(s) used as the unique row identifier.
        auto_key: When True and key_columns is None, detect keys automatically.

    Returns:
        MultiComparisonResult with per-file summaries, timeline, and overall diff.

    Raises:
        ValueError: If fewer than 2 files are provided.
        FileNotFoundError: If any file does not exist.
        ValueError: If specified key columns are missing from a file.
    """
    if len(file_paths) < 2:
        raise ValueError("At least 2 files are required for comparison.")

    frames: list[pd.DataFrame] = [_load(fp) for fp in file_paths]
    file_names = [Path(fp).name for fp in file_paths]

    if key_columns is None and auto_key:
        key_columns = _detect_key_columns(frames[0], frames[1])

    if key_columns:
        for i, (df, name) in enumerate(zip(frames, file_names)):
            missing = [c for c in key_columns if c not in df.columns]
            if missing:
                raise ValueError(f"Key columns {missing} missing in file '{name}'")

    per_file_summary: dict[str, dict] = {}
    diff_rows: list[pd.DataFrame] = []

    baseline = frames[0]
    baseline_name = file_names[0]

    for i in range(1, len(frames)):
        comp = frames[i]
        comp_name = file_names[i]
        summary, diff_df = _diff_pair(baseline, comp, key_columns, baseline_name, comp_name)
        per_file_summary[comp_name] = summary
        diff_rows.append(diff_df)

    overall_diff = pd.concat(diff_rows, ignore_index=True) if diff_rows else pd.DataFrame()

    timeline = _build_timeline(frames, file_names, key_columns)

    return MultiComparisonResult(
        per_file_summary=per_file_summary,
        timeline=timeline,
        overall_diff=overall_diff,
        file_names=file_names,
    )


def _diff_pair(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    key_columns: Optional[list[str]],
    name1: str,
    name2: str,
) -> tuple[dict, pd.DataFrame]:
    """Return (summary_dict, diff_dataframe) for two frames."""
    all_cols = list(dict.fromkeys(list(df1.columns) + list(df2.columns)))
    df1 = df1.reindex(columns=all_cols, fill_value="")
    df2 = df2.reindex(columns=all_cols, fill_value="")

    if not key_columns:
        df1 = df1.reset_index(drop=True)
        df2 = df2.reset_index(drop=True)
        n = max(len(df1), len(df2))
        rows = []
        added = removed = modified = unchanged = 0
        for idx in range(n):
            if idx >= len(df1):
                r = df2.iloc[idx].to_dict()
                r["_change"] = "added"
                r["_source_file"] = name2
                rows.append(r)
                added += 1
            elif idx >= len(df2):
                r = df1.iloc[idx].to_dict()
                r["_change"] = "removed"
                r["_source_file"] = name1
                rows.append(r)
                removed += 1
            elif not df1.iloc[idx].equals(df2.iloc[idx]):
                r = df2.iloc[idx].to_dict()
                r["_change"] = "modified"
                r["_source_file"] = name2
                rows.append(r)
                modified += 1
            else:
                unchanged += 1
        summary = {"added": added, "removed": removed, "modified": modified, "unchanged": unchanged}
        diff_df = pd.DataFrame(rows, columns=all_cols + ["_change", "_source_file"]) if rows else pd.DataFrame()
        return summary, diff_df

    df1_idx = df1.set_index(key_columns)
    df2_idx = df2.set_index(key_columns)
    keys1 = set(df1_idx.index)
    keys2 = set(df2_idx.index)

    added_keys = keys2 - keys1
    removed_keys = keys1 - keys2
    common_keys = keys1 & keys2

    rows = []
    for k in added_keys:
        r = df2_idx.loc[k].to_dict() if isinstance(df2_idx.loc[k], pd.Series) else df2_idx.loc[k].iloc[0].to_dict()
        r["_change"] = "added"
        r["_source_file"] = name2
        rows.append(r)
    for k in removed_keys:
        r = df1_idx.loc[k].to_dict() if isinstance(df1_idx.loc[k], pd.Series) else df1_idx.loc[k].iloc[0].to_dict()
        r["_change"] = "removed"
        r["_source_file"] = name1
        rows.append(r)

    modified = 0
    unchanged = 0
    for k in common_keys:
        r1 = df1_idx.loc[k] if isinstance(df1_idx.loc[k], pd.Series) else df1_idx.loc[k].iloc[0]
        r2 = df2_idx.loc[k] if isinstance(df2_idx.loc[k], pd.Series) else df2_idx.loc[k].iloc[0]
        if r1.equals(r2):
            unchanged += 1
        else:
            row = r2.to_dict()
            row["_change"] = "modified"
            row["_source_file"] = name2
            rows.append(row)
            modified += 1

    summary = {
        "added": len(added_keys),
        "removed": len(removed_keys),
        "modified": modified,
        "unchanged": unchanged,
    }
    diff_df = pd.DataFrame(rows) if rows else pd.DataFrame()
    return summary, diff_df


def _build_timeline(
    frames: list[pd.DataFrame],
    file_names: list[str],
    key_columns: Optional[list[str]],
) -> pd.DataFrame:
    """Build a timeline showing each key's value across all files."""
    if not key_columns:
        tagged = []
        for df, name in zip(frames, file_names):
            tmp = df.copy()
            tmp.insert(0, "_file", name)
            tagged.append(tmp)
        return pd.concat(tagged, ignore_index=True)

    all_keys: set = set()
    for df in frames:
        idx = df.set_index(key_columns).index
        all_keys.update(idx if hasattr(idx, '__iter__') else [idx])

    rows = []
    for key in sorted(all_keys, key=str):
        row: dict = {}
        if isinstance(key_columns, list) and len(key_columns) == 1:
            row[key_columns[0]] = key
        else:
            for col, val in zip(key_columns, key):
                row[col] = val
        for df, name in zip(frames, file_names):
            df_idx = df.set_index(key_columns)
            non_key = [c for c in df.columns if c not in key_columns]
            if key in df_idx.index:
                entry = df_idx.loc[key]
                if isinstance(entry, pd.DataFrame):
                    entry = entry.iloc[0]
                for col in non_key:
                    row[f"{name}::{col}"] = entry.get(col, "")
            else:
                for col in non_key:
                    row[f"{name}::{col}"] = "N/A"
        rows.append(row)

    return pd.DataFrame(rows)
