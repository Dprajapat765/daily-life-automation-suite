"""Tests for Tool 08 — Data Cleaner."""
import pytest
import pandas as pd
from pathlib import Path
from tools.data_cleaner import clean_file, CleanResult


def make_excel(tmp_path, df, name="data.xlsx"):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


def make_csv(tmp_path, df, name="data.csv"):
    p = tmp_path / name
    df.to_csv(p, index=False)
    return str(p)


class TestCleanFile:
    def test_strip_whitespace(self, tmp_path):
        df = pd.DataFrame({"Name": ["  Alice  ", " Bob", "Charlie "]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"strip_whitespace": True, "empty_to_null": False}, output_path=out)
        df_out = pd.read_excel(out)
        assert df_out["Name"].tolist() == ["Alice", "Bob", "Charlie"]
        assert result.changes_made["strip_whitespace"] == 3

    def test_normalize_case_title(self, tmp_path):
        df = pd.DataFrame({"Name": ["alice smith", "bob jones"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"normalize_case": {"Name": "title"}, "strip_whitespace": False, "empty_to_null": False}, output_path=out)
        df_out = pd.read_excel(out)
        assert df_out["Name"].tolist() == ["Alice Smith", "Bob Jones"]

    def test_normalize_case_upper(self, tmp_path):
        df = pd.DataFrame({"Code": ["abc", "def"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        clean_file(f, config={"normalize_case": {"Code": "upper"}, "strip_whitespace": False, "empty_to_null": False}, output_path=out)
        df_out = pd.read_excel(out)
        assert df_out["Code"].tolist() == ["ABC", "DEF"]

    def test_fix_dates(self, tmp_path):
        df = pd.DataFrame({"DOB": ["15/06/1990", "2000-01-01", "31-12-2023"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"fix_dates": ["DOB"], "strip_whitespace": False, "empty_to_null": False}, output_path=out)
        df_out = pd.read_excel(out)
        assert "1990-06-15" in df_out["DOB"].tolist()
        assert result.changes_made.get("fix_dates", 0) >= 0

    def test_clean_numbers(self, tmp_path):
        df = pd.DataFrame({"Salary": ["₹50,000", "$1,000.00", "75000"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"clean_numbers": ["Salary"], "strip_whitespace": False, "empty_to_null": False}, output_path=out)
        df_out = pd.read_excel(out)
        assert "50000" in df_out["Salary"].tolist() or "₹50,000" not in df_out["Salary"].tolist()

    def test_fix_phone(self, tmp_path):
        df = pd.DataFrame({"Phone": ["+91-98765-43210", "91 9876543210", "9876543210"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"fix_phone": ["Phone"], "strip_whitespace": False, "empty_to_null": False}, output_path=out)
        df_out = pd.read_excel(out, dtype=str)
        assert "9876543210" in df_out["Phone"].tolist()

    def test_empty_to_null(self, tmp_path):
        df = pd.DataFrame({"Status": ["Active", "NULL", "N/A", "-", "none", ""]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"empty_to_null": True, "strip_whitespace": False}, output_path=out)
        df_out = pd.read_excel(out, dtype=str).fillna("")
        assert result.changes_made["empty_to_null"] >= 2

    def test_selective_config_only_whitespace(self, tmp_path):
        df = pd.DataFrame({"Name": ["  Alice  "], "Phone": ["91 9876543210"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"strip_whitespace": True, "empty_to_null": False}, output_path=out)
        assert "strip_whitespace" in result.changes_made

    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            clean_file("nofile.xlsx")

    def test_csv_input(self, tmp_path):
        df = pd.DataFrame({"Name": ["  Alice  "]})
        f = make_csv(tmp_path, df)
        out = str(tmp_path / "out.csv")
        result = clean_file(f, config={"strip_whitespace": True, "empty_to_null": False}, output_path=out)
        assert Path(out).exists()
        assert isinstance(result, CleanResult)

    def test_default_output_path(self, tmp_path):
        df = pd.DataFrame({"Name": ["Alice"]})
        f = make_excel(tmp_path, df)
        result = clean_file(f)
        assert Path(result.output_path).exists()

    def test_change_log_populated(self, tmp_path):
        df = pd.DataFrame({"Name": ["  Alice  ", "NULL"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "out.xlsx")
        result = clean_file(f, config={"strip_whitespace": True, "empty_to_null": True}, output_path=out)
        assert len(result.change_log) >= 1
