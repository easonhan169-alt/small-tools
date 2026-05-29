from io import BytesIO

import pytest

from src.excel_diff_app import build_char_diff_html, compare_workbooks, create_replacement_workbook_bytes, get_column_letter


class FakeCell:
    def __init__(self, value):
        self.value = value


class FakeSheet:
    def __init__(self, title, values):
        self.title = title
        self.values = values
        self.max_row = max((row for row, _ in values), default=1)
        self.max_column = max((col for _, col in values), default=1)

    def cell(self, row, column):
        return FakeCell(self.values.get((row, column)))


class FakeWorkbook:
    def __init__(self, sheets):
        self._sheets = {sheet.title: sheet for sheet in sheets}
        self.sheetnames = list(self._sheets)

    def __getitem__(self, sheet_name):
        return self._sheets[sheet_name]


def _fake_workbook(sheet_name="Sheet", value="old"):
    return FakeWorkbook([FakeSheet(sheet_name, {(1, 1): value})])


def test_compare_workbooks_reports_cell_and_character_diffs():
    left = _fake_workbook(value="客户名称A")
    right = _fake_workbook(value="客户名称B")

    sheet_diffs, cell_diffs = compare_workbooks(left, right)

    assert sheet_diffs == []
    assert len(cell_diffs) == 1
    assert cell_diffs[0].sheet == "Sheet"
    assert cell_diffs[0].cell == "A1"
    assert cell_diffs[0].diff_type == "值不同"
    assert "<del>A</del><ins>B</ins>" in cell_diffs[0].char_diff_html


def test_compare_workbooks_reports_structure_diffs():
    left = _fake_workbook(sheet_name="OnlyLeft")
    right = _fake_workbook(sheet_name="OnlyRight")

    sheet_diffs, cell_diffs = compare_workbooks(left, right)

    assert cell_diffs == []
    assert {diff.sheet for diff in sheet_diffs} == {"OnlyLeft", "OnlyRight"}


def test_get_column_letter_supports_multi_letter_columns():
    assert get_column_letter(1) == "A"
    assert get_column_letter(26) == "Z"
    assert get_column_letter(27) == "AA"


def test_build_char_diff_html_handles_insertions_and_deletions():
    assert build_char_diff_html("abc", "axbc") == "a<ins>x</ins>bc"
    assert build_char_diff_html("abc", "ac") == "a<del>b</del>c"


def test_create_replacement_workbook_bytes_uses_replacement_content_when_openpyxl_is_available():
    openpyxl = pytest.importorskip("openpyxl")
    source = openpyxl.Workbook()
    source.active["A1"] = "old"
    replacement = openpyxl.Workbook()
    replacement.active["A1"] = "new"

    output_bytes = create_replacement_workbook_bytes(source, replacement)
    output = openpyxl.load_workbook(BytesIO(output_bytes))

    assert output["Sheet"]["A1"].value == "new"
