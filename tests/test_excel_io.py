from dataclasses import replace
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from sinta_shift.excel_io import (
    DRAFT, REQUIRED_FIELDS, export_result, export_template,
    inspect_workbook, preview_sheet, read_input,
)
from sinta_shift.models import Employee, InputError, SolverStatus
from sinta_shift.sample import sample_data
from sinta_shift.solver import solve_schedule as solve


def _example(tmp_path):
    return export_template(tmp_path)


def _change(path, change):
    wb = load_workbook(path)
    change(wb)
    wb.save(path)
    wb.close()


def test_template_round_trip_and_preview(tmp_path):
    path = _example(tmp_path)
    assert inspect_workbook(path) == list(REQUIRED_FIELDS)
    actual = read_input(path)
    expected = sample_data()
    assert [(e.id, e.name, e.work_days) for e in actual.employees] == [(e.id, e.name, e.work_days) for e in expected.employees]
    assert [(r.request_id, r.employee_id, r.day, r.priority) for r in actual.requests] == [(r.request_id, r.employee_id, r.day, r.priority) for r in expected.requests]
    assert not actual.fixed_off
    frame = preview_sheet(path, "職員")
    assert list(frame.columns) == list(REQUIRED_FIELDS["職員"])
    assert frame.index.tolist() == [2, 3, 4]
    assert actual.employees[0].source.file == path.name
    assert actual.employees[0].source.sheet == "職員"
    assert actual.employees[0].source.row == 2


def test_formula_rejected_with_source_and_unused_formula_ignored(tmp_path):
    path = _example(tmp_path)
    _change(path, lambda wb: setattr(wb["職員"]["F2"], "value", "=1+1"))
    assert len(read_input(path).employees) == 3
    _change(path, lambda wb: setattr(wb["職員"]["C2"], "value", "=1+1"))
    with pytest.raises(InputError) as caught:
        read_input(path)
    issue = caught.value.issues[0]
    assert issue.kind == "数式"
    assert (issue.source.file, issue.source.sheet, issue.source.row, issue.source.column) == (path.name, "職員", 2, "C2")
    assert "値を貼り付け" in issue.message


def test_string_identifiers_na_zero_and_numeric_id(tmp_path):
    path = _example(tmp_path)
    def change(wb):
        wb["職員"]["A2"] = "001"
        wb["職員"]["A3"] = "NA"
        wb["職員"]["A4"] = 12
        wb["職員"]["A4"].number_format = "0000"
    _change(path, change)
    data = read_input(path)
    assert [e.id for e in data.employees] == ["001", "NA", "12"]
    frame = preview_sheet(path, "職員")
    assert frame.iloc[0, 0] == "001"
    assert frame.iloc[1, 0] == "NA"


def test_column_sheet_header_mapping(tmp_path):
    path = _example(tmp_path)
    def change(wb):
        ws = wb["職員"]
        ws.title = "メンバー"
        ws.insert_rows(1)
        ws["A2"] = "社員コード"
        ws["B2"] = "名前"
        ws["C2"] = "出勤日数"
    _change(path, change)
    mappings = {"職員": {"sheet": "メンバー", "header_row": 2, "columns": {"職員ID": "社員コード", "氏名": "名前", "対象期間の勤務日数": "出勤日数"}}}
    assert read_input(path, mappings).employees[0].source.row == 3
    assert preview_sheet(path, "メンバー", 2).index.tolist() == [3, 4, 5]
    with pytest.raises(InputError):
        read_input(path)


@pytest.mark.parametrize("value", [None, "", "abc", 1.5, True, "NaN", "Infinity"])
def test_bad_integer_is_not_zero(tmp_path, value):
    path = _example(tmp_path)
    _change(path, lambda wb: setattr(wb["職員"]["C2"], "value", value))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(i.source.row == 2 and i.source.column == "C2" for i in caught.value.issues)


@pytest.mark.parametrize("value", ["11-07", "2026-02-30", 46333, None])
def test_full_dates_required(tmp_path, value):
    path = _example(tmp_path)
    def change(wb):
        ws = wb["必要人数"]
        ws["A2"] = value
        ws["A2"].number_format = "General"
    _change(path, change)
    with pytest.raises(InputError):
        read_input(path)


def test_duplicate_headers_and_missing_sheet_errors(tmp_path):
    path = _example(tmp_path)
    _change(path, lambda wb: setattr(wb["職員"]["B1"], "value", "職員ID"))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any("重複" in i.message for i in caught.value.issues)
    _change(path, lambda wb: wb.remove(wb["確定休"]))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(i.kind == "シート" and i.source.sheet == "確定休" for i in caught.value.issues)


def test_result_workbook_counts_safe_text_and_unique_save(tmp_path):
    data = sample_data()
    data = replace(data, employees=(replace(data.employees[0], name="=1+1"), *data.employees[1:]))
    result = solve(data)
    assert result.status == SolverStatus.OPTIMAL
    first = export_result(data, result.schedule, result, tmp_path)
    second = export_result(data, result.schedule, result, tmp_path)
    assert first != second and first.exists() and second.exists()
    wb = load_workbook(first, data_only=False)
    assert wb.sheetnames == ["勤務表", "日別人数", "希望休の採用結果", "設定条件"]
    for ws in wb:
        assert ws["A1"].value == DRAFT
        assert ws.freeze_panes and ws.auto_filter.ref
        assert any(r.min_row == 1 and r.max_col == ws.max_column for r in ws.merged_cells.ranges)
        assert ws["A1"].alignment.wrap_text
        assert all(c.data_type != "f" for row in ws for c in row)
    assert "A11:G11" in {str(r) for r in wb["勤務表"].merged_cells.ranges}
    assert wb["勤務表"]["A11"].alignment.wrap_text
    assert wb["勤務表"].row_dimensions[11].height >= 32
    assert wb["勤務表"]["B4"].value == "=1+1"
    assert wb["勤務表"]["B4"].data_type == "s"
    assert [wb["日別人数"].cell(row, 3).value for row in range(4, 7)] == [2, 2, 2]
    assert any("第2希望" in str(cell.value) for row in wb["希望休の採用結果"] for cell in row)
    wb.close()


def test_export_blocks_violations_and_original_folder(tmp_path):
    path = _example(tmp_path)
    data = read_input(path)
    result = solve(data)
    with pytest.raises(ValueError, match="別のフォルダ"):
        export_result(data, result.schedule, result, tmp_path, input_path=path)
    broken = dict(result.schedule)
    broken[("A", date(2026, 11, 7))] = True
    with pytest.raises(InputError):
        export_result(data, broken, result, tmp_path)
    assert not list(tmp_path.glob("sinta_shift_result*"))


def test_edited_output_does_not_claim_old_optimality(tmp_path):
    data = sample_data()
    result = solve(data)
    path = export_result(data, result.schedule, result, tmp_path, edited=True)
    wb = load_workbook(path)
    settings = {row[0]: row[1] for row in wb["設定条件"].iter_rows(min_row=4, values_only=True)}
    assert settings["採用件数の最適性"] == "未確認"
    assert "手修正" in settings["計算状態"]
    wb.close()


def test_empty_requests_supported_and_preview_limited(tmp_path):
    path = _example(tmp_path)
    def change(wb):
        wb["希望休"].delete_rows(2, 100)
        for index in range(30):
            wb["職員"].append([f"ID{index:03}", f"名前{index}", 2])
    _change(path, change)
    assert read_input(path).requests == ()
    assert len(preview_sheet(path, "職員")) == 20


def test_other_formats_and_corrupt_files_rejected(tmp_path):
    with pytest.raises(InputError, match="xlsx"):
        read_input(tmp_path / "old.xls")
    broken = tmp_path / "broken.xlsx"
    broken.write_text("not excel")
    with pytest.raises(InputError, match="開けません"):
        read_input(broken)
