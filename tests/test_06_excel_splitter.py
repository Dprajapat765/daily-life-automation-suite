"""Tests for Tool 06 — Excel Splitter."""
import pytest
import pandas as pd
from pathlib import Path
from tools.excel_splitter import split_by_column, SplitResult


def make_excel(tmp_path, df, name="data.xlsx"):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name="data.csv"):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


@pytest.fixture
def region_df():
    return pd.DataFrame({
        "ID":     ["1","2","3","4","5","6"],
        "Name":   ["A","B","C","D","E","F"],
        "Region": ["North","South","East","North","South","East"],
    })


class TestSplitByColumn:
    def test_split_3_unique_values(self, tmp_path, region_df):
        f = make_excel(tmp_path, region_df)
        result = split_by_column(f, "Region", str(tmp_path / "out"))
        assert len(result.files_created) == 3
        assert isinstance(result, SplitResult)

    def test_rows_per_file_correct(self, tmp_path, region_df):
        f = make_excel(tmp_path, region_df)
        result = split_by_column(f, "Region", str(tmp_path / "out"))
        total = sum(result.rows_per_file.values())
        assert total == 6

    def test_split_5_unique_values(self, tmp_path):
        df = pd.DataFrame({
            "ID":   list(range(1, 11)),
            "Group": ["A","B","C","D","E","A","B","C","D","E"],
        })
        df = df.astype(str)
        f = make_excel(tmp_path, df)
        result = split_by_column(f, "Group", str(tmp_path / "out"))
        assert len(result.files_created) == 5

    def test_value_with_spaces_safe_filename(self, tmp_path):
        df = pd.DataFrame({"Name": ["Alice Jones", "Bob"], "Dept": ["Human Resources", "IT"]})
        f = make_excel(tmp_path, df)
        result = split_by_column(f, "Dept", str(tmp_path / "out"))
        names = [Path(fp).name for fp in result.files_created]
        assert any("Human_Resources" in n for n in names)

    def test_output_format_csv(self, tmp_path, region_df):
        f = make_excel(tmp_path, region_df)
        result = split_by_column(f, "Region", str(tmp_path / "out"), output_format="csv")
        assert all(fp.endswith(".csv") for fp in result.files_created)

    def test_invalid_column_raises(self, tmp_path, region_df):
        f = make_excel(tmp_path, region_df)
        with pytest.raises(ValueError, match="Column"):
            split_by_column(f, "NoSuchCol", str(tmp_path / "out"))

    def test_invalid_output_format_raises(self, tmp_path, region_df):
        f = make_excel(tmp_path, region_df)
        with pytest.raises(ValueError, match="output_format"):
            split_by_column(f, "Region", str(tmp_path / "out"), output_format="pdf")

    def test_file_not_found_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            split_by_column("nofile.xlsx", "Region", str(tmp_path / "out"))

    def test_single_value_column(self, tmp_path):
        df = pd.DataFrame({"ID": ["1","2","3"], "Type": ["X","X","X"]})
        f = make_excel(tmp_path, df)
        result = split_by_column(f, "Type", str(tmp_path / "out"))
        assert len(result.files_created) == 1
        assert result.rows_per_file["X.xlsx"] == 3

    def test_csv_source_file(self, tmp_path, region_df):
        f = make_csv(tmp_path, region_df)
        result = split_by_column(f, "Region", str(tmp_path / "out"), output_format="csv")
        assert len(result.files_created) == 3

    def test_output_dir_created(self, tmp_path, region_df):
        f = make_excel(tmp_path, region_df)
        out = str(tmp_path / "deep" / "nested" / "out")
        result = split_by_column(f, "Region", out)
        assert Path(result.output_dir).exists()
