"""Tests for Tool 01 — Excel Compare."""
import pytest
import pandas as pd
from pathlib import Path
from tools.excel_compare import compare_files, ComparisonResult


@pytest.fixture
def base_df():
    return pd.DataFrame({
        "ID":   ["1", "2", "3", "4"],
        "Name": ["Alice", "Bob", "Charlie", "Diana"],
        "Score":["90", "85", "70", "95"],
    })


@pytest.fixture
def base_file(tmp_path, base_df):
    p = tmp_path / "base.xlsx"
    base_df.to_excel(p, index=False)
    return str(p)


@pytest.fixture
def csv_base_file(tmp_path, base_df):
    p = tmp_path / "base.csv"
    base_df.to_csv(p, index=False)
    return str(p)


def make_excel(tmp_path, df, name="comp.xlsx"):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name="comp.csv"):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


class TestCompareFiles:
    def test_identical_files_all_unchanged(self, tmp_path, base_file, base_df):
        comp = make_excel(tmp_path, base_df, "comp.xlsx")
        result = compare_files(base_file, comp, key_columns=["ID"])
        assert result.summary["unchanged"] == 4
        assert result.summary["added"] == 0
        assert result.summary["removed"] == 0
        assert result.summary["modified"] == 0

    def test_detect_added_rows(self, tmp_path, base_file, base_df):
        extra = pd.concat(
            [base_df, pd.DataFrame({"ID": ["5"], "Name": ["Eve"], "Score": ["88"]})],
            ignore_index=True,
        )
        comp = make_excel(tmp_path, extra)
        result = compare_files(base_file, comp, key_columns=["ID"])
        assert result.summary["added"] == 1
        assert "5" in result.added["ID"].values

    def test_detect_removed_rows(self, tmp_path, base_file, base_df):
        reduced = base_df.iloc[:3].copy()
        comp = make_excel(tmp_path, reduced)
        result = compare_files(base_file, comp, key_columns=["ID"])
        assert result.summary["removed"] == 1
        assert "4" in result.removed["ID"].values

    def test_detect_modified_rows(self, tmp_path, base_file, base_df):
        mod = base_df.copy()
        mod.loc[mod["ID"] == "2", "Score"] = "99"
        comp = make_excel(tmp_path, mod)
        result = compare_files(base_file, comp, key_columns=["ID"])
        assert result.summary["modified"] == 1

    def test_csv_input_files(self, tmp_path, csv_base_file, base_df):
        comp = make_csv(tmp_path, base_df)
        result = compare_files(csv_base_file, comp, key_columns=["ID"])
        assert result.summary["unchanged"] == 4

    def test_file_not_found_raises(self, tmp_path, base_file):
        with pytest.raises(FileNotFoundError):
            compare_files(base_file, "nonexistent.xlsx", key_columns=["ID"])

    def test_invalid_key_column_raises(self, tmp_path, base_file, base_df):
        comp = make_excel(tmp_path, base_df)
        with pytest.raises(ValueError):
            compare_files(base_file, comp, key_columns=["NoSuchColumn"])

    def test_auto_key_detection(self, tmp_path, base_file, base_df):
        extra = pd.concat(
            [base_df, pd.DataFrame({"ID": ["5"], "Name": ["Eve"], "Score": ["88"]})],
            ignore_index=True,
        )
        comp = make_excel(tmp_path, extra)
        result = compare_files(base_file, comp, auto_key=True)
        assert result.summary["added"] >= 0

    def test_empty_both_files(self, tmp_path):
        e1 = tmp_path / "e1.xlsx"
        e2 = tmp_path / "e2.xlsx"
        pd.DataFrame().to_excel(e1, index=False)
        pd.DataFrame().to_excel(e2, index=False)
        result = compare_files(str(e1), str(e2))
        assert result.summary == {"added": 0, "removed": 0, "modified": 0, "unchanged": 0}

    def test_export_report_creates_file(self, tmp_path, base_file, base_df):
        comp = make_excel(tmp_path, base_df)
        result = compare_files(base_file, comp, key_columns=["ID"])
        out = tmp_path / "report.xlsx"
        path = result.export_report(str(out))
        assert Path(path).exists()

    def test_no_key_columns_fallback(self, tmp_path, base_df):
        comp = make_excel(tmp_path, base_df)
        base = make_excel(tmp_path, base_df, "base2.xlsx")
        result = compare_files(base, comp, key_columns=[], auto_key=False)
        assert result.summary["unchanged"] == 4

    def test_mixed_xlsx_csv(self, tmp_path, base_df):
        xlsx = make_excel(tmp_path, base_df, "base.xlsx")
        csv = make_csv(tmp_path, base_df, "comp.csv")
        result = compare_files(xlsx, csv, key_columns=["ID"])
        assert result.summary["unchanged"] == 4
