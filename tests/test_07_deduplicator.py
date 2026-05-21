"""Tests for Tool 07 — Deduplicator."""
import pytest
import pandas as pd
from pathlib import Path
from tools.deduplicator import find_duplicates, DedupResult


def make_excel(tmp_path, df, name="data.xlsx"):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


@pytest.fixture
def df_with_dups():
    return pd.DataFrame({
        "ID":   ["1","2","2","3","3","3"],
        "Name": ["Alice","Bob","Bob","Charlie","Charlie","Charlie"],
        "Dept": ["HR","IT","IT","Finance","Finance","Finance"],
    })


@pytest.fixture
def df_no_dups():
    return pd.DataFrame({
        "ID":   ["1","2","3"],
        "Name": ["Alice","Bob","Charlie"],
        "Dept": ["HR","IT","Finance"],
    })


class TestFindDuplicates:
    def test_flag_action_all_rows_kept(self, tmp_path, df_with_dups):
        f = make_excel(tmp_path, df_with_dups)
        out = str(tmp_path / "out.xlsx")
        result = find_duplicates(f, action="flag", output_path=out)
        assert result.total_rows == 6
        assert result.duplicate_count > 0
        df = pd.read_excel(out)
        assert "Duplicate_Status" in df.columns
        assert len(df) == 6

    def test_remove_keep_first(self, tmp_path, df_with_dups):
        f = make_excel(tmp_path, df_with_dups)
        out = str(tmp_path / "out.xlsx")
        result = find_duplicates(f, action="remove_keep_first", output_path=out)
        df = pd.read_excel(out)
        assert len(df) == 3
        assert result.duplicate_count > 0

    def test_remove_keep_last(self, tmp_path, df_with_dups):
        f = make_excel(tmp_path, df_with_dups)
        out = str(tmp_path / "out.xlsx")
        result = find_duplicates(f, action="remove_keep_last", output_path=out)
        df = pd.read_excel(out)
        assert len(df) == 3

    def test_no_duplicates_file(self, tmp_path, df_no_dups):
        f = make_excel(tmp_path, df_no_dups)
        out = str(tmp_path / "out.xlsx")
        result = find_duplicates(f, action="flag", output_path=out)
        assert result.duplicate_count == 0
        assert result.duplicates_path is None

    def test_all_duplicates(self, tmp_path):
        df = pd.DataFrame({"A": ["x","x","x"], "B": ["y","y","y"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = find_duplicates(f, action="remove_keep_first", output_path=out)
        assert result.duplicate_count == 3

    def test_key_columns_partial_duplicate(self, tmp_path):
        df = pd.DataFrame({
            "ID":    ["1","1","2"],
            "Name":  ["Alice","Alice_v2","Bob"],
            "Score": ["90","85","70"],
        })
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = find_duplicates(f, key_columns=["ID"], action="remove_keep_first", output_path=out)
        df_out = pd.read_excel(out)
        assert len(df_out) == 2

    def test_invalid_action_raises(self, tmp_path, df_no_dups):
        f = make_excel(tmp_path, df_no_dups)
        with pytest.raises(ValueError, match="action"):
            find_duplicates(f, action="invalid_action")

    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            find_duplicates("nofile.xlsx")

    def test_invalid_key_column_raises(self, tmp_path, df_no_dups):
        f = make_excel(tmp_path, df_no_dups)
        with pytest.raises(ValueError, match="Key columns"):
            find_duplicates(f, key_columns=["NoCol"])

    def test_output_path_created(self, tmp_path, df_with_dups):
        f = make_excel(tmp_path, df_with_dups)
        out = str(tmp_path / "sub" / "out.xlsx")
        result = find_duplicates(f, action="flag", output_path=out)
        assert Path(result.output_path).exists()

    def test_duplicates_file_written(self, tmp_path, df_with_dups):
        f = make_excel(tmp_path, df_with_dups)
        out = str(tmp_path / "out.xlsx")
        result = find_duplicates(f, action="flag", output_path=out)
        assert result.duplicates_path is not None
        assert Path(result.duplicates_path).exists()
