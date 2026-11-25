"""公休・有給・明けと勤務時刻の入出力。旧4表の互換性も保持する。"""
from dataclasses import replace
from decimal import Decimal

import pytest
from openpyxl import load_workbook

from sinta_shift.excel_io import (
    DERIVE_PUBLIC_DAYS, PUBLIC_HOLIDAYS, SETTINGS_SHEET, SHIFT_TIMES_SHEET,
    export_leave_sample, export_result, export_template, read_input,
    settings_from_payload, settings_to_payload,
)
from sinta_shift.models import (
    DAY_ONLY, TWO_SHIFT, THREE_SHIFT, InputError, ShiftSettings, ShiftTime,
    SolveResult, SolverStatus,
)
from sinta_shift.planning import configure_input
from sinta_shift.solver import solve_schedule


def _edit(path, action):
    workbook = load_workbook(path)
    try:
        action(workbook)
        workbook.save(path)
    finally:
        workbook.close()


def _confirm(path):
    def edit(wb):
        for row in wb[SETTINGS_SHEET]:
            if row[0].value == "勤務時刻確認済み":
                row[1].value = "確認済み"
    _edit(path, edit)


def test_leave_settings_and_global_history_times_round_trip():
    settings = ShiftSettings(
        leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS, default_public_holidays=8,
        weekday_ratios=(Decimal("50.25"),) * 7,
        shift_times=(ShiftTime("日勤", 480, 1020), ShiftTime("夜勤", 960, 1980)),
        shift_times_confirmed=True,
    )
    assert settings_from_payload(settings_to_payload(settings)) == settings


@pytest.mark.parametrize("field,value", (
    ("leave_enabled", "yes"), ("day_count_mode", "unknown"),
    ("default_public_holidays", -1), ("default_public_holidays", True),
    ("shift_times_confirmed", "confirmed"),
    ("shift_times", [{"label": "日勤", "start_minute": 480, "end_minute": 400}]),
    ("shift_times", [{"label": "日勤", "start_minute": True, "end_minute": 1020}]),
    ("shift_times", [{"label": "不明", "start_minute": 480, "end_minute": 1020}]),
    ("shift_times", [{"label": "日勤", "start_minute": 480, "end_minute": 2000}]),
))
def test_invalid_new_settings_rejected(field, value):
    with pytest.raises(InputError):
        settings_from_payload({field: value})


def test_default_leave_sample_thirty_days_eight_public_two_paid_twenty_work(tmp_path):
    path = export_leave_sample(tmp_path)
    raw = read_input(path)
    assert raw.settings.leave_enabled
    assert raw.settings.day_count_mode == DERIVE_PUBLIC_DAYS
    assert len(raw.days) == 30 and len(raw.employees) == 3
    assert all(employee.id.startswith("00") and employee.work_days is None and employee.public_holidays == 8 for employee in raw.employees)
    assert len(raw.paid_leaves) == 6
    assert raw.paid_leaves[0].source.file == path.name
    assert raw.paid_leaves[0].source.sheet == "有給"
    data = configure_input(raw)
    assert all(employee.work_days == 20 for employee in data.employees)
    result = solve_schedule(data)
    assert result.status in (SolverStatus.OPTIMAL, SolverStatus.FEASIBLE)
    output = export_result(data, result.schedule, result, tmp_path)
    workbook = load_workbook(output)
    assert workbook.sheetnames == ["勤務表", "日別人数", "希望休の採用結果", "設定条件"]
    headers = {cell.value: cell.column for cell in workbook["勤務表"][3]}
    for row in range(4, 7):
        for label, expected in (("勤務開始回数", 20), ("実勤務暦日（明け含む）", 20), ("公休目標", 8), ("公休実績", 8), ("有給日数", 2), ("明け日数", 0)):
            assert workbook["勤務表"].cell(row, headers[label]).value == expected
    metadata = dict(workbook["設定条件"].iter_rows(min_row=4, values_only=True))
    assert "残高は更新しない" in metadata["有給の扱い"]
    assert all(cell.data_type != "f" for ws in workbook for row in ws for cell in row)
    workbook.close()


def test_derived_work_days_column_can_be_missing_or_unused(tmp_path):
    path = export_leave_sample(tmp_path)
    _edit(path, lambda wb: wb["職員"].delete_cols(3))
    assert all(employee.work_days is None for employee in read_input(path).employees)
    legacy = export_template(tmp_path)
    data = read_input(legacy, {"職員": {"columns": {"対象期間の勤務日数": ""}}},
                      settings_override=ShiftSettings(leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS, default_public_holidays=1))
    assert all(employee.work_days is None for employee in data.employees)


def test_settings_override_allows_blank_work_days_only_in_derived_mode(tmp_path):
    path = export_template(tmp_path)
    _edit(path, lambda wb: setattr(wb["職員"]["C2"], "value", None))
    with pytest.raises(InputError):
        read_input(path)
    data = read_input(path, settings_override=ShiftSettings(leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS, default_public_holidays=1))
    assert data.employees[0].work_days is None
    with pytest.raises(InputError):
        read_input(path, settings_override=ShiftSettings(leave_enabled=True))


def test_paid_sheet_can_be_empty_renamed_or_explicitly_unused(tmp_path):
    path = export_leave_sample(tmp_path)
    _edit(path, lambda wb: setattr(wb["有給"], "title", "承認済有給"))
    data = read_input(path, {"有給": {"sheet": "承認済有給"}})
    assert len(data.paid_leaves) == 6
    assert read_input(path, {"有給": {"sheet": ""}}).paid_leaves == ()
    _edit(path, lambda wb: wb["承認済有給"].delete_rows(2, 100))
    assert read_input(path, {"有給": {"sheet": "承認済有給"}}).paid_leaves == ()


@pytest.mark.parametrize("problem", ("unknown", "duplicate", "outside", "fixed_conflict"))
def test_invalid_paid_leave_blocks_export_with_original_source(tmp_path, problem):
    path = export_leave_sample(tmp_path)
    def edit(wb):
        ws = wb["有給"]
        if problem == "unknown":
            ws["A2"] = "UNKNOWN"
        elif problem == "duplicate":
            ws.append([ws["A2"].value, ws["B2"].value])
        elif problem == "outside":
            ws["B2"] = "2026-12-01"
        else:
            wb["確定休"].append([ws["A2"].value, ws["B2"].value])
    _edit(path, edit)
    with pytest.raises(InputError) as caught:
        export_result(read_input(path), {}, SolveResult(SolverStatus.FEASIBLE), tmp_path)
    assert any(issue.source.file == path.name and issue.source.sheet == "有給" and issue.source.row >= 2 for issue in caught.value.issues)
    assert not list(tmp_path.glob("sinta_shift_result*"))


def test_paid_identifiers_preserve_literal_na_and_leading_zero(tmp_path):
    path = export_leave_sample(tmp_path)
    def edit(wb):
        wb["有給"]["A2"] = "NA"
        wb["有給"]["A3"] = "0007"
    _edit(path, edit)
    assert [item.employee_id for item in read_input(path).paid_leaves[:2]] == ["NA", "0007"]


@pytest.mark.parametrize("role,cell", (("職員", "F2"), ("有給", "B2"), (SHIFT_TIMES_SHEET, "B2")))
def test_new_used_columns_reject_formulas_with_source(tmp_path, role, cell):
    path = export_leave_sample(tmp_path, ShiftSettings(TWO_SHIFT, leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS))
    _edit(path, lambda wb: setattr(wb[role][cell], "value", "=1+1"))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.kind == "数式" and issue.source.file == path.name and issue.source.sheet == role and issue.source.row == 2 and issue.source.column == cell for issue in caught.value.issues)


def test_shift_time_formula_in_header_is_reported_as_formula(tmp_path):
    path = export_leave_sample(tmp_path, ShiftSettings(TWO_SHIFT, leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS))
    _edit(path, lambda wb: setattr(wb[SHIFT_TIMES_SHEET]["B1"], "value", '="開始時刻"'))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.kind == "数式" and issue.source.sheet == SHIFT_TIMES_SHEET and issue.source.row == 1 for issue in caught.value.issues)


def test_public_holiday_column_can_be_mapped_and_blank_is_global_override(tmp_path):
    path = export_leave_sample(tmp_path)
    def edit(wb):
        wb["職員"]["F1"] = "休日目標"
        wb["職員"]["F2"] = None
        wb["職員"]["F3"] = 7
    _edit(path, edit)
    data = read_input(path, {"職員": {"columns": {PUBLIC_HOLIDAYS: "休日目標"}}})
    assert data.employees[0].public_holidays is None
    assert data.employees[1].public_holidays == 7


def test_time_example_stays_unconfirmed_until_explicit_confirmation(tmp_path):
    path = export_leave_sample(tmp_path, ShiftSettings(TWO_SHIFT, leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS))
    data = read_input(path)
    assert data.settings.shift_times == (ShiftTime("日勤", 480, 1020), ShiftTime("夜勤", 960, 1980))
    assert not data.settings.shift_times_confirmed
    with pytest.raises(InputError) as caught:
        export_result(data, {}, SolveResult(SolverStatus.FEASIBLE), tmp_path)
    assert any("時刻" in issue.message and "確認" in issue.message for issue in caught.value.issues)


def test_three_shift_midnight_boundary_twenty_four_has_no_spurious_next_day(tmp_path):
    path = export_leave_sample(tmp_path, ShiftSettings(THREE_SHIFT, leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS))
    settings = read_input(path).settings
    assert next(item for item in settings.shift_times if item.label == "準夜勤").end_minute == 1440
    assert not next(item for item in settings.shift_times if item.label == "準夜勤").crosses_midnight


@pytest.mark.parametrize("cell,value", (("B2", "24:00"), ("C2", "08:00"), ("D2", 2), ("B2", True)))
def test_bad_shift_time_value_has_original_cell_and_is_not_guessed(tmp_path, cell, value):
    path = export_leave_sample(tmp_path, ShiftSettings(TWO_SHIFT, leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS))
    _edit(path, lambda wb: setattr(wb[SHIFT_TIMES_SHEET][cell], "value", value))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.sheet == SHIFT_TIMES_SHEET and issue.source.row == 2 for issue in caught.value.issues)


def test_confirmed_overnight_sample_export_separates_recovery_and_public_off(tmp_path):
    path = export_leave_sample(tmp_path, ShiftSettings(TWO_SHIFT, leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS))
    _confirm(path)
    data = configure_input(read_input(path))
    assert all(employee.work_days is None for employee in data.employees)
    result = solve_schedule(data)
    assert result.schedule is not None
    output = export_result(data, result.schedule, result, tmp_path)
    workbook = load_workbook(output)
    headers = {cell.value: cell.column for cell in workbook["勤務表"][3]}
    for row in range(4, 10):
        starts = workbook["勤務表"].cell(row, headers["勤務開始回数"]).value
        recovery = workbook["勤務表"].cell(row, headers["明け日数"]).value
        assert starts + recovery == 20
        assert workbook["勤務表"].cell(row, headers["公休実績"]).value == 8
        assert workbook["勤務表"].cell(row, headers["有給日数"]).value == 2
        assert recovery >= 0
        assert workbook["勤務表"].cell(row, headers["実勤務暦日（明け含む）"]).value == 20
    metadata = dict(workbook["設定条件"].iter_rows(min_row=4, values_only=True))
    assert "未確認" in metadata["期間外の翌日確認"]
    assert metadata["翌日明け確認対象の職員ID"] != "なし"
    assert any("明けとなる勤務" in str(cell.value) for row in workbook["勤務表"] for cell in row)
    workbook.close()
