"""Tests for Tool 02 — Multi-File Compare."""
import pytest
import pandas as pd
from pathlib import Path
from tools.multi_compare import compare_multiple, MultiComparisonResult


def make_excel(tmp_path, df, name):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


@pytest.fixture
def df1():
    return pd.DataFrame({
        "ID":   ["1", "2", "3"],
        "Name": ["Alice", "Bob", "Charlie"],
        "Dept": ["HR", "IT", "Finance"],
    })


@pytest.fixture
def df2():
    return pd.DataFrame({
        "ID":   ["1", "2", "4"],
        "Name": ["Alice", "Bob_v2", "Diana"],
        "Dept": ["HR", "IT", "Sales"],
    })


@pytest.fixture
def df3():
    return pd.DataFrame({
        "ID":   ["1", "4", "5"],
        "Name": ["Alice", "Diana", "Eve"],
        "Dept": ["HR", "Sales", "Legal"],
    })


class TestCompareMultiple:
    def test_three_files_key_based(self, tmp_path, df1, df2, df3):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        f3 = make_excel(tmp_path, df3, "f3.xlsx")
        result = compare_multiple([f1, f2, f3], key_columns=["ID"])
        assert isinstance(result, MultiComparisonResult)
        assert "f2.xlsx" in result.per_file_summary
        assert "f3.xlsx" in result.per_file_summary
        assert result.per_file_summary["f2.xlsx"]["added"] == 1
        assert result.per_file_summary["f2.xlsx"]["removed"] == 1

    def test_four_files(self, tmp_path, df1, df2, df3):
        df4 = pd.DataFrame({
            "ID":   ["1", "6"],
            "Name": ["Alice", "Frank"],
            "Dept": ["HR", "Ops"],
        })
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        f3 = make_excel(tmp_path, df3, "f3.xlsx")
        f4 = make_excel(tmp_path, df4, "f4.xlsx")
        result = compare_multiple([f1, f2, f3, f4], key_columns=["ID"])
        assert len(result.per_file_summary) == 3

    def test_different_column_order(self, tmp_path, df1):
        df_reordered = df1[["Dept", "Name", "ID"]]
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df_reordered, "f2.xlsx")
        result = compare_multiple([f1, f2], key_columns=["ID"])
        assert result.per_file_summary["f2.xlsx"]["unchanged"] == 3

    def test_mixed_xlsx_and_csv(self, tmp_path, df1, df2):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_csv(tmp_path, df2, "f2.csv")
        result = compare_multiple([f1, f2], key_columns=["ID"])
        assert "f2.csv" in result.per_file_summary

    def test_too_few_files_raises(self, tmp_path, df1):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        with pytest.raises(ValueError, match="At least 2"):
            compare_multiple([f1])

    def test_file_not_found_raises(self, tmp_path, df1):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        with pytest.raises(FileNotFoundError):
            compare_multiple([f1, "no_such_file.xlsx"])

    def test_invalid_key_column_raises(self, tmp_path, df1, df2):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        with pytest.raises(ValueError):
            compare_multiple([f1, f2], key_columns=["NoCol"])

    def test_timeline_contains_all_keys(self, tmp_path, df1, df2):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        result = compare_multiple([f1, f2], key_columns=["ID"])
        all_ids = set(result.timeline["ID"].astype(str))
        assert {"1", "2", "3", "4"}.issubset(all_ids)

    def test_overall_diff_has_source_file_column(self, tmp_path, df1, df2):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        result = compare_multiple([f1, f2], key_columns=["ID"])
        assert "_source_file" in result.overall_diff.columns

    def test_export_report_creates_file(self, tmp_path, df1, df2):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        result = compare_multiple([f1, f2], key_columns=["ID"])
        out = tmp_path / "report.xlsx"
        path = result.export_report(str(out))
        assert Path(path).exists()

    def test_identical_files_all_unchanged(self, tmp_path, df1):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df1, "f2.xlsx")
        result = compare_multiple([f1, f2], key_columns=["ID"])
        assert result.per_file_summary["f2.xlsx"]["unchanged"] == 3
        assert result.per_file_summary["f2.xlsx"]["modified"] == 0

    def test_auto_key_detection(self, tmp_path, df1, df2):
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        result = compare_multiple([f1, f2], auto_key=True)
        assert isinstance(result, MultiComparisonResult)
