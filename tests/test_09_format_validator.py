"""Tests for Tool 09 — Format Validator."""
import pytest
import pandas as pd
from pathlib import Path
from tools.format_validator import validate_formats, ValidationReport


def make_excel(tmp_path, df, name="data.xlsx"):
    p = tmp_path / name
    df.to_excel(p, index=False)
    return str(p)


class TestValidateFormats:
    def test_email_valid(self, tmp_path):
        df = pd.DataFrame({"Email": ["user@example.com", "test.me+tag@domain.org"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Email": "email"})
        assert result.per_column_stats["Email"]["valid"] == 2
        assert result.per_column_stats["Email"]["invalid"] == 0

    def test_email_invalid(self, tmp_path):
        df = pd.DataFrame({"Email": ["not-an-email", "missing@", "@nodomain"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Email": "email"})
        assert result.per_column_stats["Email"]["invalid"] == 3

    def test_phone_india_valid(self, tmp_path):
        df = pd.DataFrame({"Phone": ["9876543210", "+919876543210", "8765432109"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Phone": "phone_india"})
        assert result.per_column_stats["Phone"]["valid"] == 3

    def test_pincode_valid_and_invalid(self, tmp_path):
        df = pd.DataFrame({"PIN": ["110001", "12345", "ABCDEF", "560001"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"PIN": "pincode"})
        assert result.per_column_stats["PIN"]["valid"] == 2
        assert result.per_column_stats["PIN"]["invalid"] == 2

    def test_pan_valid(self, tmp_path):
        df = pd.DataFrame({"PAN": ["ABCDE1234F", "ZZZZZ9999Z"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"PAN": "pan"})
        assert result.per_column_stats["PAN"]["valid"] == 2

    def test_date_multiple_formats(self, tmp_path):
        df = pd.DataFrame({"DOB": ["2000-01-15", "15/06/1990", "31-12-2023", "not-a-date"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"DOB": "date"})
        assert result.per_column_stats["DOB"]["valid"] == 3
        assert result.per_column_stats["DOB"]["invalid"] == 1

    def test_positive_number(self, tmp_path):
        df = pd.DataFrame({"Amount": ["100", "0", "-50", "abc", "3.14"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Amount": "positive_number"})
        assert result.per_column_stats["Amount"]["valid"] == 2

    def test_non_empty(self, tmp_path):
        df = pd.DataFrame({"Name": [None, "Alice", None, "Bob"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Name": "non_empty"})
        assert result.per_column_stats["Name"]["empty"] >= 1
        assert result.per_column_stats["Name"]["valid"] >= 2

    def test_multiple_columns_validated(self, tmp_path):
        df = pd.DataFrame({
            "Email": ["user@example.com", "bad"],
            "PIN":   ["110001", "12345"],
        })
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Email": "email", "PIN": "pincode"})
        assert "Email" in result.per_column_stats
        assert "PIN" in result.per_column_stats

    def test_all_valid_rate_100(self, tmp_path):
        df = pd.DataFrame({"Email": ["a@b.com", "x@y.org"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Email": "email"})
        assert result.valid_rate_pct == 100.0

    def test_all_invalid_rate_0(self, tmp_path):
        df = pd.DataFrame({"Email": ["bad", "also-bad"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Email": "email"})
        assert result.valid_rate_pct == 0.0

    def test_column_not_found_raises(self, tmp_path):
        df = pd.DataFrame({"Email": ["a@b.com"]})
        f = make_excel(tmp_path, df)
        with pytest.raises(ValueError, match="Column"):
            validate_formats(f, {"NoCol": "email"})

    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            validate_formats("nofile.xlsx", {"Email": "email"})

    def test_invalid_rows_list(self, tmp_path):
        df = pd.DataFrame({"Email": ["good@example.com", "bad-email"]})
        f = make_excel(tmp_path, df)
        result = validate_formats(f, {"Email": "email"})
        assert len(result.invalid_rows) == 1

    def test_output_file_created(self, tmp_path):
        df = pd.DataFrame({"Email": ["a@b.com"]})
        f = make_excel(tmp_path, df)
        out = str(tmp_path / "validated.xlsx")
        result = validate_formats(f, {"Email": "email"}, output_path=out)
        assert Path(out).exists()
        assert isinstance(result, ValidationReport)
