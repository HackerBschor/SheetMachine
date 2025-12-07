import os
import platform
import subprocess
import tempfile

from xlsxwriter import Workbook

from SheetMachine import Field, Creator


if __name__ == '__main__':
    # Create test fields & data
    test_fields = [
        Field(name="a", title="A", group_by=True, hidden=True),
        Field(name="b", title="B", group_by=True, hidden=True),
        Field(name="c", title="C", total_summary = Field.AggregationType.SUM, group_summary=Field.AggregationType.AVG),
        Field(name="d", title="D", total_summary = Field.AggregationType.SUM, group_summary=Field.AggregationType.SUM, excel_format_def={'num_format': '#,##0.00€'}),
        Field(name="e", title="E"),
        Field(name="f", title="F", formular="=$?c?$<i>+$?d?$<i>", total_summary = Field.AggregationType.SUM, group_summary=Field.AggregationType.SUM),
    ]

    test_data = [
        {"a": "A", "b": "a", "c": 0.1, "d": 1, "e": "1990-01-01"},
        {"a": "A", "b": "b", "c": 0.2, "d": 2, "e": "1990-01-02"},
        {"a": "A", "b": "a", "c": 0.1, "d": 1, "e": "1990-01-01"},
        {"a": "A", "b": "b", "c": 0.2, "d": 2, "e": "1990-01-02"},
        {"a": "C", "b": "a", "c": 0.3, "d": 3, "e": "1990-01-03"},
        {"a": "D", "b": "a", "c": 0.4, "d": 4, "e": "1990-01-04"},
        {"a": "E", "b": "a", "c": 0.5, "d": 5, "e": "1990-01-05"},
    ]

    # Create test file
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, "test_file.xlsx")

    # Create test workbook
    test_workbook = Workbook(file_path)
    test_worksheet = test_workbook.add_worksheet()

    # Generate Excel file
    creator = Creator(test_workbook, test_worksheet, fields=test_fields, group_formatter = lambda row: f'{row["a"]} + {row["b"]}')
    creator.create(test_data)
    test_workbook.close()

    # Open Excel file
    system = platform.system()

    if system == "Windows":
        os.startfile(file_path)
    elif system == "Darwin":  # macOS
        subprocess.call(["open", file_path])
    else:  # Linux / others
        subprocess.call(["xdg-open", file_path])
