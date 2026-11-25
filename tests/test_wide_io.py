"""希望日1/2/3を独立した日付申請として取り込む受入テスト。"""
from datetime import date

import pytest
from openpyxl import load_workbook

from sinta_shift.excel_io import (
    WIDE_REQUEST_FIELDS, export_independent_sample, export_result, export_template, read_input,
)
from sinta_shift.models import InputError, PUBLIC_REQUEST, PAID_REQUEST, PUBLIC_OFF, PAID_LEAVE, SolveResult, SolverStatus
from sinta_shift.planning import confirm_paid_requests, configure_input
from sinta_shift.solver import solve_schedule
from sinta_shift.evaluation import evaluate_schedule


def _edit(path, action):
    wb = load_workbook(path)
    try:
        action(wb)
        wb.save(path)
    finally:
        wb.close()


def _sample(tmp_path):
    return export_independent_sample(tmp_path)


def _confirm(data):
    return confirm_paid_requests(data, [item.request_id for item in data.independent_requests if item.leave_type == PAID_REQUEST])


def test_wide_sample_has_nine_independent_days_and_no_silent_paid_conversion(tmp_path):
    path = _sample(tmp_path)
    data = read_input(path)
    assert len(data.employees) == 3 and len(data.days) == 30
    assert data.settings.default_public_holidays == 9
    assert all(employee.public_holidays == 9 for employee in data.employees)
    assert len(data.requests) == 6
    assert data.request_format == "independent"
    assert len(data.independent_requests) == 9
    assert len({item.request_id for item in data.independent_requests}) == 9
    assert sum(item.leave_type == PUBLIC_REQUEST for item in data.independent_requests) == 6
    assert sum(item.leave_type == PAID_REQUEST for item in data.independent_requests) == 3
    assert all(item.priority == 1 for item in data.requests)
    assert data.paid_leaves == ()
    assert data.confirmed_paid_request_ids == frozenset()
    assert data.independent_requests[0].source.file == path.name
    assert data.independent_requests[0].source.sheet == "希望休"
    assert data.independent_requests[0].source.column == "B2"
    wb = load_workbook(path)
    assert tuple(cell.value for cell in wb["希望休"][1][:7]) == WIDE_REQUEST_FIELDS
    assert wb["有給"].max_row == 1
    assert len(wb["希望休"].data_validations.dataValidation) == 3
    assert wb["希望休"]["A2"].value == "001"
    assert wb["希望休"]["A2"].data_type == "s"
    assert wb["希望休"]["B2"].number_format == "yyyy-mm-dd"
    assert "独立" in wb["希望休"]["I1"].value
    wb.close()


def test_unconfirmed_wide_paid_requests_stop_calculation_and_export(tmp_path):
    data = read_input(_sample(tmp_path))
    result = solve_schedule(data)
    assert result.schedule is None
    assert any("確認" in issue.message for issue in result.issues) or "確認" in result.message
    with pytest.raises(InputError) as caught:
        export_result(data, {}, SolveResult(SolverStatus.FEASIBLE), tmp_path)
    assert any("確認" in issue.message for issue in caught.value.issues)
    assert not list(tmp_path.glob("sinta_shift_result*"))


def test_confirmed_sample_calculates_all_three_days_and_exports_nine_rows(tmp_path):
    data = configure_input(_confirm(read_input(_sample(tmp_path))))
    assert len(data.paid_leaves) == 3
    assert len(data.confirmed_paid_request_ids) == 3
    assert all(employee.work_days == 20 for employee in data.employees)
    result = solve_schedule(data)
    assert result.status in (SolverStatus.OPTIMAL, SolverStatus.FEASIBLE)
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid
    assert evaluation.adopted_count == 6
    assert len(evaluation.paid_request_results) == 3
    for item in data.independent_requests:
        assert result.schedule[item.employee_id, item.day] == (PUBLIC_OFF if item.leave_type == PUBLIC_REQUEST else PAID_LEAVE)
    output = export_result(data, result.schedule, result, tmp_path)
    wb = load_workbook(output, data_only=False)
    assert wb.sheetnames == ["勤務表", "日別人数", "希望休の採用結果", "設定条件"]
    ws = wb["希望休の採用結果"]
    assert tuple(cell.value for cell in ws[3]) == ("申請ID", "職員ID", "氏名", "希望日", "結果日", "日数", "結果", "区分")
    assert not any("順位" in str(cell.value) or "候補" in str(cell.value) for cell in ws[3])
    rows = list(ws.iter_rows(min_row=4, values_only=True))
    assert len(rows) == 9
    assert sum(row[7] == PUBLIC_REQUEST for row in rows) == 6
    assert all(row[5] == 1 for row in rows)
    assert sum(row[7] == PAID_REQUEST and row[6] == "有給確定" for row in rows) == 3
    assert sum(row[7] == PUBLIC_REQUEST and row[6] == "採用" for row in rows) == 6
    assert not any("第1希望" in str(value) for row in rows for value in row)
    assert "公休希望採用 6/6" in ws["A2"].value
    assert "有給確定 3/3" in ws["A2"].value
    assert all(row[3] and row[4] for row in rows)
    metadata = dict(wb["設定条件"].iter_rows(min_row=4, values_only=True))
    assert metadata["独立申請件数"] == 9
    assert metadata["独立有給確定件数"] == 3
    assert not any("順位" in str(key) for key in metadata)
    assert "独立" in metadata["希望条件"]
    assert all(request_id in metadata["確認済み有給申請ID"] for request_id in data.confirmed_paid_request_ids)
    assert all(cell.data_type != "f" for sheet in wb for row in sheet for cell in row)
    wb.close()


def test_explicit_wide_template_keeps_legacy_default_unchanged(tmp_path):
    old_path = export_template(tmp_path)
    old = read_input(old_path)
    assert len(old.requests) == 4
    assert old.independent_requests == ()
    old_result = solve_schedule(old)
    output = export_result(old, old_result.schedule, old_result, tmp_path)
    wb = load_workbook(output)
    assert tuple(cell.value for cell in wb["希望休の採用結果"][3]) == ("申請ID", "職員ID", "氏名", "候補日・順位", "採用日", "採用順位", "採用状況")
    assert any(row[0] == "R2" and row[5] == 2 and "第2希望" in row[6] for row in wb["希望休の採用結果"].iter_rows(min_row=4, values_only=True))
    wb.close()
    wide = read_input(export_template(tmp_path, request_format="independent"))
    assert wide.independent_requests


def test_unadopted_wide_public_request_still_means_one_day_without_rank_text(tmp_path):
    path = _sample(tmp_path)
    _edit(path, lambda wb: setattr(wb["希望休"]["B3"], "value", wb["希望休"]["B2"].value))
    data = configure_input(_confirm(read_input(path)))
    result = solve_schedule(data)
    assert result.schedule is not None
    wb = load_workbook(export_result(data, result.schedule, result, tmp_path))
    rows = list(wb["希望休の採用結果"].iter_rows(min_row=4, values_only=True))
    assert all(row[5] == 1 for row in rows)
    assert any(row[7] == PUBLIC_REQUEST and row[6] == "未採用" and row[4] is None for row in rows)
    assert not any("第1希望" in str(value) for row in rows for value in row)
    wb.close()


@pytest.mark.parametrize("field,cell,value,reason", (
    ("missing_kind", "C2", None, "両方"),
    ("missing_day", "B2", None, "両方"),
    ("bad_kind", "C2", "休み", "区分"),
    ("bad_date", "B2", "11/07", "年を含む"),
    ("outside", "B2", "2026-12-01", "期間外"),
    ("unknown", "A2", "UNKNOWN", "未知"),
))
def test_invalid_wide_values_show_exact_original_cell(tmp_path, field, cell, value, reason):
    path = _sample(tmp_path)
    _edit(path, lambda wb: setattr(wb["希望休"][cell], "value", value))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.file == path.name and issue.source.sheet == "希望休" and issue.source.row == 2 and issue.source.column == cell and reason in issue.message for issue in caught.value.issues)


def test_both_empty_means_no_application_without_zero_or_date_guess(tmp_path):
    path = _sample(tmp_path)
    def edit(wb):
        wb["希望休"]["B2"] = None
        wb["希望休"]["C2"] = None
    _edit(path, edit)
    data = read_input(path)
    assert len(data.independent_requests) == 8
    assert len(data.requests) == 5


def test_zero_wide_applications_still_keep_eight_output_columns_and_format_audit(tmp_path):
    path = _sample(tmp_path)
    def edit(wb):
        ws = wb["希望休"]
        for row in ws.iter_rows(min_row=2, max_row=4, min_col=2, max_col=7):
            for cell in row:
                cell.value = None
        for row in wb["職員"].iter_rows(min_row=2, min_col=6, max_col=6):
            row[0].value = 10
    _edit(path, edit)
    data = read_input(path)
    assert data.request_format == "independent"
    assert data.independent_requests == () and data.requests == () and data.paid_leaves == ()
    result = solve_schedule(data)
    assert result.schedule is not None
    wb = load_workbook(export_result(data, result.schedule, result, tmp_path))
    assert wb["希望休の採用結果"].max_column == 8
    assert wb["希望休の採用結果"]["H3"].value == "区分"
    assert wb["希望休の採用結果"].max_row == 3
    metadata = dict(wb["設定条件"].iter_rows(min_row=4, values_only=True))
    assert "独立" in metadata["希望休入力形式"]
    assert metadata["独立申請件数"] == 0
    wb.close()


def test_duplicate_day_and_duplicate_employee_rows_are_rejected(tmp_path):
    path = _sample(tmp_path)
    _edit(path, lambda wb: setattr(wb["希望休"]["D2"], "value", wb["希望休"]["B2"].value))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.column == "D2" and "重複" in issue.message for issue in caught.value.issues)
    other = _sample(tmp_path)
    _edit(other, lambda wb: wb["希望休"].append(["001", date(2026, 11, 28), "公休希望", None, None, None, None]))
    with pytest.raises(InputError) as caught:
        read_input(other)
    assert any(issue.source.column == "A5" and "1職員1行" in issue.message for issue in caught.value.issues)


@pytest.mark.parametrize("cell", ("A2", "B2", "C2", "F2", "G2"))
def test_formulas_in_used_wide_cells_block_reading(tmp_path, cell):
    path = _sample(tmp_path)
    _edit(path, lambda wb: setattr(wb["希望休"][cell], "value", "=1+1"))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.kind == "数式" and issue.source.row == 2 and issue.source.column == cell for issue in caught.value.issues)


def test_unused_formula_does_not_become_an_input(tmp_path):
    path = _sample(tmp_path)
    _edit(path, lambda wb: setattr(wb["希望休"]["M2"], "value", "=1+1"))
    assert len(read_input(path).independent_requests) == 9


def test_full_width_header_numbers_and_kind_aliases_are_normalized(tmp_path):
    path = _sample(tmp_path)
    def edit(wb):
        ws = wb["希望休"]
        for cell in ws[1][:7]:
            cell.value = str(cell.value).translate(str.maketrans("123", "１２３"))
        ws["C2"] = "公休"
        ws["G2"] = "有給"
    _edit(path, edit)
    data = read_input(path)
    assert len(data.independent_requests) == 9
    assert data.independent_requests[0].leave_type == PUBLIC_REQUEST
    assert data.independent_requests[2].leave_type == PAID_REQUEST


def test_explicit_format_with_renamed_columns_sheet_and_header_row(tmp_path):
    path = _sample(tmp_path)
    names = ("社員コード", "休暇日A", "休暇種別A", "休暇日B", "休暇種別B", "休暇日C", "休暇種別C")
    def edit(wb):
        ws = wb["希望休"]
        ws.title = "個人申請"
        ws.insert_rows(1)
        for cell, name in zip(ws[2][:7], names):
            cell.value = name
    _edit(path, edit)
    with pytest.raises(InputError):
        read_input(path)
    mapping = {"希望休": {"sheet": "個人申請", "header_row": 2, "format": "independent", "columns": dict(zip(WIDE_REQUEST_FIELDS, names))}}
    data = read_input(path, mapping)
    assert len(data.independent_requests) == 9
    assert data.independent_requests[0].source.row == 3
    assert data.independent_requests[0].source.column == "B3"


def test_old_saved_alias_mapping_without_format_key_keeps_alternative_semantics(tmp_path):
    path = export_template(tmp_path)
    names = ("受付番号", "社員コード", "休暇候補日", "候補順")
    def edit(wb):
        for cell, name in zip(wb["希望休"][1], names):
            cell.value = name
    _edit(path, edit)
    mapping = {"希望休": {"columns": dict(zip(("申請ID", "職員ID", "候補日", "優先順位"), names))}}
    data = read_input(path, mapping)
    assert data.request_format == "alternatives"
    assert len(data.requests) == 4
    assert sum(item.request_id == "R2" for item in data.requests) == 2
    assert not data.independent_requests


def test_mixed_headers_need_explicit_format_instead_of_guessing(tmp_path):
    path = _sample(tmp_path)
    def edit(wb):
        ws = wb["希望休"]
        ws.unmerge_cells("I1:L4")
        ws["H1"] = "申請ID"
        ws["I1"] = "候補日"
        ws["J1"] = "優先順位"
    _edit(path, edit)
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any("混在" in issue.message for issue in caught.value.issues)
    assert len(read_input(path, {"希望休": {"format": "independent"}}).independent_requests) == 9
