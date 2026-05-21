"""Tests for Tool 03 — List Validator."""
import pytest
import pandas as pd
from pathlib import Path
from tools.list_validator import validate_list, ValidationResult


def make_excel(tmp_path, df, name):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


@pytest.fixture
def source_df():
    return pd.DataFrame({"Name": ["Alice", "Bob", "Charlie", "Daina", "XYZ_Unknown"]})


@pytest.fixture
def target_df():
    return pd.DataFrame({"Employee": ["Alice", "Bob", "Charlie", "Diana", "Eve"]})


class TestValidateList:
    def test_exact_matches(self, tmp_path, source_df, target_df):
        src = make_excel(tmp_path, pd.DataFrame({"Name": ["Alice", "Bob"]}), "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee")
        assert len(result.exact_matches) == 2
        assert result.match_rate_pct == 100.0

    def test_fuzzy_near_matches(self, tmp_path, target_df):
        src = make_excel(tmp_path, pd.DataFrame({"Name": ["Daina"]}), "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee", fuzzy_threshold=70)
        assert len(result.near_matches) >= 1
        assert result.near_matches[0]["Match_Type"] == "NEAR MATCH"

    def test_missing_values(self, tmp_path, target_df):
        src = make_excel(tmp_path, pd.DataFrame({"Name": ["XYZ_Unknown"]}), "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee", fuzzy_threshold=85)
        assert "XYZ_Unknown" in result.missing

    def test_mixed_results(self, tmp_path, source_df, target_df):
        src = make_excel(tmp_path, source_df, "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee", fuzzy_threshold=85)
        total = len(result.exact_matches) + len(result.near_matches) + len(result.missing)
        assert total == 5

    def test_threshold_100_only_exact(self, tmp_path, target_df):
        src = make_excel(tmp_path, pd.DataFrame({"Name": ["Daina", "Alice"]}), "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee", fuzzy_threshold=100)
        assert len(result.exact_matches) == 1
        assert result.exact_matches[0]["Source"] == "Alice"

    def test_threshold_0_all_matched(self, tmp_path, source_df, target_df):
        src = make_excel(tmp_path, source_df, "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee", fuzzy_threshold=0)
        assert result.match_rate_pct == 100.0

    def test_file_not_found_raises(self, tmp_path, target_df):
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        with pytest.raises(FileNotFoundError):
            validate_list("nofile.xlsx", tgt, "Name", "Employee")

    def test_column_not_found_raises(self, tmp_path, source_df, target_df):
        src = make_excel(tmp_path, source_df, "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        with pytest.raises(ValueError, match="Column"):
            validate_list(src, tgt, "BadCol", "Employee")

    def test_csv_inputs(self, tmp_path, source_df, target_df):
        src = make_csv(tmp_path, source_df, "s.csv")
        tgt = make_csv(tmp_path, target_df, "t.csv")
        result = validate_list(src, tgt, "Name", "Employee")
        assert isinstance(result, ValidationResult)

    def test_empty_source(self, tmp_path, target_df):
        src = make_excel(tmp_path, pd.DataFrame({"Name": []}), "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee")
        assert result.match_rate_pct == 0.0

    def test_export_report_creates_file(self, tmp_path, source_df, target_df):
        src = make_excel(tmp_path, source_df, "s.xlsx")
        tgt = make_excel(tmp_path, target_df, "t.xlsx")
        result = validate_list(src, tgt, "Name", "Employee", fuzzy_threshold=70)
        out = tmp_path / "report.xlsx"
        path = result.export_report(str(out))
        assert Path(path).exists()
