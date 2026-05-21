"""Tests for Tool 04 — Fuzzy Matcher."""
import pytest
import pandas as pd
from tools.fuzzy_matcher import fuzzy_match, fuzzy_match_files


def make_excel(tmp_path, df, name="f.xlsx"):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name="f.csv"):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


class TestFuzzyMatch:
    def test_exact_strings_same_file(self, tmp_path):
        df = pd.DataFrame({"A": ["Hello", "World"], "B": ["Hello", "World"]})
        f = make_excel(tmp_path, df)
        result = fuzzy_match(f, "A", "B", threshold=80)
        assert all(result["Match_Type"] == "EXACT")
        assert all(result["Score_Pct"] == 100.0)

    def test_similar_strings_with_typos(self, tmp_path):
        df = pd.DataFrame({"A": ["Alice Smith"], "B": ["Alic Smith"]})
        f = make_excel(tmp_path, df)
        result = fuzzy_match(f, "A", "B", threshold=80)
        assert result.iloc[0]["Match_Type"] == "SIMILAR"
        assert result.iloc[0]["Score_Pct"] >= 80

    def test_completely_different_strings(self, tmp_path):
        df = pd.DataFrame({"A": ["AAAA"], "B": ["ZZZZ"]})
        f = make_excel(tmp_path, df)
        result = fuzzy_match(f, "A", "B", threshold=80)
        assert result.iloc[0]["Match_Type"] == "NO MATCH"

    def test_threshold_100_no_similar(self, tmp_path):
        df = pd.DataFrame({"A": ["Alice Smith"], "B": ["Alic Smith"]})
        f = make_excel(tmp_path, df)
        result = fuzzy_match(f, "A", "B", threshold=100)
        assert result.iloc[0]["Match_Type"] == "NO MATCH"

    def test_threshold_0_all_matched(self, tmp_path):
        df = pd.DataFrame({"A": ["AAAA"], "B": ["ZZZZ"]})
        f = make_excel(tmp_path, df)
        result = fuzzy_match(f, "A", "B", threshold=0)
        assert result.iloc[0]["Match_Type"] in ("EXACT", "SIMILAR")

    def test_column_not_found_raises(self, tmp_path):
        df = pd.DataFrame({"A": ["Hello"], "B": ["Hello"]})
        f = make_excel(tmp_path, df)
        with pytest.raises(ValueError, match="Column"):
            fuzzy_match(f, "A", "NoCol", threshold=80)

    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            fuzzy_match("nofile.xlsx", "A", "B")

    def test_result_columns(self, tmp_path):
        df = pd.DataFrame({"X": ["foo"], "Y": ["bar"]})
        f = make_excel(tmp_path, df)
        result = fuzzy_match(f, "X", "Y", threshold=80)
        assert "X" in result.columns
        assert "Y" in result.columns
        assert "Match_Type" in result.columns
        assert "Score_Pct" in result.columns

    def test_csv_input(self, tmp_path):
        df = pd.DataFrame({"A": ["hello"], "B": ["hello"]})
        f = make_csv(tmp_path, df)
        result = fuzzy_match(f, "A", "B")
        assert result.iloc[0]["Match_Type"] == "EXACT"


class TestFuzzyMatchFiles:
    def test_exact_match_two_files(self, tmp_path):
        df1 = pd.DataFrame({"Name": ["Alice", "Bob"]})
        df2 = pd.DataFrame({"Employee": ["Alice", "Bob"]})
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        result = fuzzy_match_files(f1, "Name", f2, "Employee")
        assert all(result["Match_Type"] == "EXACT")

    def test_near_match_two_files(self, tmp_path):
        df1 = pd.DataFrame({"Name": ["Alic"]})
        df2 = pd.DataFrame({"Employee": ["Alice"]})
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        result = fuzzy_match_files(f1, "Name", f2, "Employee", threshold=75)
        assert result.iloc[0]["Match_Type"] == "SIMILAR"

    def test_different_lengths(self, tmp_path):
        df1 = pd.DataFrame({"Name": ["Alice", "Bob", "Charlie"]})
        df2 = pd.DataFrame({"Employee": ["Alice"]})
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        result = fuzzy_match_files(f1, "Name", f2, "Employee")
        assert len(result) == 3

    def test_file_not_found_raises(self, tmp_path):
        df1 = pd.DataFrame({"Name": ["Alice"]})
        f1 = make_excel(tmp_path, df1)
        with pytest.raises(FileNotFoundError):
            fuzzy_match_files(f1, "Name", "nofile.xlsx", "Employee")

    def test_column_not_found_raises(self, tmp_path):
        df1 = pd.DataFrame({"Name": ["Alice"]})
        df2 = pd.DataFrame({"Employee": ["Alice"]})
        f1 = make_excel(tmp_path, df1, "f1.xlsx")
        f2 = make_excel(tmp_path, df2, "f2.xlsx")
        with pytest.raises(ValueError):
            fuzzy_match_files(f1, "BadCol", f2, "Employee")
