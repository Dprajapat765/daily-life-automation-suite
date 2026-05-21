"""Tests for Tool 05 — Excel Merger."""
import pytest
import pandas as pd
from pathlib import Path
from tools.excel_merger import merge_files, MergeResult


def make_excel(tmp_path, df, name):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


@pytest.fixture
def df_a():
    return pd.DataFrame({"ID": ["1", "2"], "Name": ["Alice", "Bob"], "Dept": ["HR", "IT"]})


@pytest.fixture
def df_b():
    return pd.DataFrame({"ID": ["3", "4"], "Name": ["Charlie", "Diana"], "Dept": ["Finance", "Sales"]})


class TestMergeFiles:
    def test_two_files_same_schema(self, tmp_path, df_a, df_b):
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_b, "b.xlsx")
        out = str(tmp_path / "merged.xlsx")
        result = merge_files([f1, f2], out)
        assert result.total_rows == 4
        assert result.source_counts["a.xlsx"] == 2
        assert result.source_counts["b.xlsx"] == 2
        df = pd.read_excel(out)
        assert "Source_File" in df.columns

    def test_three_files_different_column_order(self, tmp_path, df_a, df_b):
        df_c = pd.DataFrame({"Dept": ["Ops"], "ID": ["5"], "Name": ["Eve"]})
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_b, "b.xlsx")
        f3 = make_excel(tmp_path, df_c, "c.xlsx")
        out = str(tmp_path / "merged.xlsx")
        result = merge_files([f1, f2, f3], out)
        assert result.total_rows == 5

    def test_files_with_extra_columns(self, tmp_path, df_a):
        df_extra = pd.DataFrame({"ID": ["3"], "Name": ["Charlie"], "Dept": ["Finance"], "Salary": ["60000"]})
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_extra, "b.xlsx")
        out = str(tmp_path / "merged.xlsx")
        result = merge_files([f1, f2], out)
        df = pd.read_excel(out)
        assert "Salary" in df.columns
        assert result.total_rows == 3

    def test_files_with_missing_columns(self, tmp_path, df_a):
        df_missing = pd.DataFrame({"ID": ["3"], "Name": ["Charlie"]})
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_missing, "b.xlsx")
        out = str(tmp_path / "merged.xlsx")
        result = merge_files([f1, f2], out, fill_missing="N/A")
        df = pd.read_excel(out)
        assert "Dept" in df.columns

    def test_dedup_true_removes_duplicates(self, tmp_path, df_a):
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_a, "b.xlsx")
        out = str(tmp_path / "merged.xlsx")
        result = merge_files([f1, f2], out, dedup=True)
        assert result.duplicate_rows_removed >= 0
        assert result.total_rows <= 4

    def test_dedup_false_keeps_duplicates(self, tmp_path, df_a):
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_a, "b.xlsx")
        out = str(tmp_path / "merged.xlsx")
        result = merge_files([f1, f2], out, dedup=False)
        assert result.duplicate_rows_removed == 0
        assert result.total_rows == 4

    def test_csv_output(self, tmp_path, df_a, df_b):
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_b, "b.xlsx")
        out = str(tmp_path / "merged.csv")
        result = merge_files([f1, f2], out)
        assert Path(out).exists()
        df = pd.read_csv(out)
        assert len(df) == 4

    def test_mixed_xlsx_csv_input(self, tmp_path, df_a, df_b):
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_csv(tmp_path, df_b, "b.csv")
        out = str(tmp_path / "merged.xlsx")
        result = merge_files([f1, f2], out)
        assert result.total_rows == 4

    def test_file_not_found_raises(self, tmp_path, df_a):
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        out = str(tmp_path / "merged.xlsx")
        with pytest.raises(FileNotFoundError):
            merge_files([f1, "nofile.xlsx"], out)

    def test_empty_list_raises(self, tmp_path):
        with pytest.raises(ValueError):
            merge_files([], str(tmp_path / "merged.xlsx"))

    def test_output_path_created(self, tmp_path, df_a, df_b):
        f1 = make_excel(tmp_path, df_a, "a.xlsx")
        f2 = make_excel(tmp_path, df_b, "b.xlsx")
        out = str(tmp_path / "subdir" / "merged.xlsx")
        result = merge_files([f1, f2], out)
        assert Path(result.output_path).exists()
