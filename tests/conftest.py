import pytest
import pandas as pd
from pathlib import Path

TEST_DATA_DIR = Path(__file__).parent.parent / "test_data"
OUTPUTS_DIR   = Path(__file__).parent.parent / "outputs"

@pytest.fixture(autouse=True)
def ensure_dirs():
    TEST_DATA_DIR.mkdir(exist_ok=True)
    OUTPUTS_DIR.mkdir(exist_ok=True)

@pytest.fixture
def sample_employees_df():
    return pd.DataFrame({
        "Employee ID": ["E001","E002","E003","E004","E005"],
        "Name":        ["Alice","Bob","Charlie","Diana","Eve"],
        "Department":  ["HR","IT","Finance","IT","HR"],
        "Salary":      ["50000","60000","55000","70000","52000"],
        "Location":    ["Delhi","Mumbai","Pune","Bangalore","Delhi"],
    })

@pytest.fixture
def sample_excel_file(tmp_path, sample_employees_df):
    path = tmp_path / "employees.xlsx"
    sample_employees_df.to_excel(path, index=False)
    return str(path)
