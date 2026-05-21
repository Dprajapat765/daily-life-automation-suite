"""Tests for Tool 10 — Null Checker."""
import pytest
import pandas as pd
from pathlib import Path
from tools.null_checker import check_nulls, NullReport


def make_excel(tmp_path, df, name="data.xlsx"):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name="data.csv"):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


class TestCheckNulls:
    def test_file_with_no_nulls(self, tmp_path):
        df = pd.DataFrame({"ID": ["1","2","3"], "Name": ["Alice","Bob","Charlie"]})
        f = make_excel(tmp_path, df)
        result = check_nulls(f)
        assert result.health_score_pct == 100.0
        assert result.column_stats["ID"]["null_count"] == 0
        assert result.column_stats["Name"]["null_count"] == 0

    def test_file_with_some_nulls(self, tmp_path):
        df = pd.DataFrame({"ID": ["1", None, "3"], "Name": ["Alice", "Bob", None]})
        f = make_excel(tmp_path, df)
        result = check_nulls(f)
        assert result.column_stats["ID"]["null_count"] >= 1
        assert result.health_score_pct < 100.0

    def test_mandatory_cols_flag_critical_rows(self, tmp_path):
        df = pd.DataFrame({"ID": [None, "2", "3"], "Name": ["Alice", "Bob", "Charlie"]})
        f = make_excel(tmp_path, df)
        result = check_nulls(f, mandatory_columns=["ID"])
        assert len(result.critical_rows) >= 1

    def test_no_mandatory_cols_no_critical(self, tmp_path):
        df = pd.DataFrame({"ID": [None, "2"], "Name": ["Alice", "Bob"]})
        f = make_excel(tmp_path, df)
        result = check_nulls(f, mandatory_columns=[])
        assert len(result.critical_rows) == 0

    def test_null_pct_calculation(self, tmp_path):
        df = pd.DataFrame({"Col": [None, None, "val", None, "val"]})
        f = make_excel(tmp_path, df)
        result = check_nulls(f)
        assert result.column_stats["Col"]["null_pct"] == 60.0

    def test_row_flags_all_rows_present(self, tmp_path):
        df = pd.DataFrame({"A": ["1","2","3"]})
        f = make_excel(tmp_path, df)
        result = check_nulls(f)
        assert len(result.row_flags) == 3

    def test_csv_input(self, tmp_path):
        df = pd.DataFrame({"ID": ["1","2"], "Name": ["Alice","Bob"]})
        f = make_csv(tmp_path, df)
        result = check_nulls(f)
        assert isinstance(result, NullReport)
        assert result.health_score_pct == 100.0

    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            check_nulls("nofile.xlsx")

    def test_export_report_creates_file(self, tmp_path):
        df = pd.DataFrame({"ID": [None, "2"], "Name": ["Alice", "Bob"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "report.xlsx")
        result = check_nulls(f, output_path=out)
        assert Path(out).exists()
        assert result.output_path is not None

    def test_null_like_strings_detected(self, tmp_path):
        df = pd.DataFrame({"Status": ["N/A", "NULL", "-", "Active", "none"]})
        f = make_excel(tmp_path, df)
        result = check_nulls(f)
        assert result.column_stats["Status"]["null_count"] >= 3

    def test_all_mandatory_cols_missing_all_critical(self, tmp_path):
        df = pd.DataFrame({
            "ID": [None, None, None],
            "Name": ["Alice", "Bob", "Charlie"],
        })
        f = make_excel(tmp_path, df)
        result = check_nulls(f, mandatory_columns=["ID"])
        assert len(result.critical_rows) == 3
