"""週・日付の出勤率と、有給希望から公休希望への配分監査。"""
from dataclasses import replace
from datetime import date
from decimal import Decimal
import hashlib
import json

import pytest
from openpyxl import load_workbook

from sinta_shift.excel_io import (
    ATTENDANCE_FIELDS, ATTENDANCE_SHEET, export_independent_sample, export_result,
    export_template, read_input, settings_from_payload, settings_to_payload,
)
from sinta_shift.models import (
    DayRequirement, Employee, InputData, InputError, PaidLeave, PAID_REQUEST,
    PUBLIC_REQUEST, ShiftSettings, SolverStatus, Source,
)
from sinta_shift.planning import configure_input, set_leave_choices
from sinta_shift.solver import solve_schedule


MONDAY = date(2026, 11, 2)
TUESDAY = date(2026, 11, 3)


def _edit(path, action):
    wb = load_workbook(path)
    try:
        action(wb)
        wb.save(path)
    finally:
        wb.close()


def test_week_and_date_json_round_trip_keeps_decimal_precision_and_no_approval():
    settings = ShiftSettings(weekly_ratios=((MONDAY, Decimal("70.125000000000000001")),), date_ratios=((TUESDAY, Decimal("50")),))
    payload = settings_to_payload(settings)
    assert payload["weekly_ratios"] == [{"week_start": "2026-11-02", "percentage": "70.125000000000000001"}]
    assert payload["date_ratios"] == [{"day": "2026-11-03", "percentage": "50"}]
    assert settings_from_payload(payload) == settings
    assert "confirmed_paid_request_ids" not in payload
    assert "converted_public_request_ids" not in payload
    with pytest.raises(InputError):
        settings_from_payload({"converted_public_request_ids": ["arbitrary"]})


@pytest.mark.parametrize("payload", (
    {"weekly_ratios": [{"week_start": "2026-11-03", "percentage": "70"}]},
    {"weekly_ratios": [{"week_start": "2026-11-02", "percentage": "70"}] * 2},
    {"date_ratios": [{"day": "2026-11-03", "percentage": "70"}] * 2},
    {"date_ratios": [{"day": "2026-11-3", "percentage": "70"}]},
    {"date_ratios": [{"day": "2026-02-30", "percentage": "70"}]},
    {"date_ratios": [{"day": "2026-11-03", "percentage": "101"}]},
    {"date_ratios": [{"day": "2026-11-03", "percentage": True}]},
    {"date_ratios": [{"day": "2026-11-03", "percentage": 0.5}]},
    {"date_ratios": [{"day": "2026-11-03", "percentage": "NaN"}]},
    {"date_ratios": [{"day": "2026-11-03", "percentage": "50", "unknown": 1}]},
    {"weekly_ratios": None},
))
def test_invalid_override_json_is_rejected(payload):
    with pytest.raises(InputError):
        settings_from_payload(payload)


def test_optional_adjustment_sheet_round_trip_and_legacy_four_sheets(tmp_path):
    legacy = export_template(tmp_path)
    assert read_input(legacy).settings.weekly_ratios == ()
    assert read_input(legacy).settings.date_ratios == ()
    settings = ShiftSettings(weekly_ratios=((MONDAY, Decimal("66.66")),), date_ratios=((date(2026, 11, 8), Decimal("33.33")),))
    path = export_template(tmp_path, settings)
    data = read_input(path)
    assert data.settings == settings
    wb = load_workbook(path)
    assert tuple(cell.value for cell in wb[ATTENDANCE_SHEET][1]) == ATTENDANCE_FIELDS
    assert wb[ATTENDANCE_SHEET]["B2"].number_format == "yyyy-mm-dd"
    wb.close()


def _manual_adjustment(tmp_path, rows):
    path = export_template(tmp_path)
    def edit(wb):
        ws = wb.create_sheet(ATTENDANCE_SHEET)
        ws.append(ATTENDANCE_FIELDS)
        for row in rows:
            ws.append(row)
    _edit(path, edit)
    return path


def test_adjustment_sheet_is_loaded_without_optional_conditions_sheet(tmp_path):
    path = _manual_adjustment(tmp_path, [("週", MONDAY, 70), ("日付", date(2026, 11, 7), "50")])
    data = read_input(path)
    assert data.settings.weekly_ratios == ((MONDAY, Decimal("70")),)
    assert data.settings.date_ratios == ((date(2026, 11, 7), Decimal("50")),)


@pytest.mark.parametrize("row", (("日付", date(2026, 11, 10), 50), ("週", date(2026, 11, 16), 50)))
def test_outside_adjustments_keep_the_original_date_cell(tmp_path, row):
    path = _manual_adjustment(tmp_path, [row])
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.file == path.name and issue.source.sheet == ATTENDANCE_SHEET and issue.source.row == 2 and issue.source.column == "B2" for issue in caught.value.issues)


def test_week_start_before_period_is_valid_when_week_overlaps(tmp_path):
    path = _manual_adjustment(tmp_path, [("週", MONDAY, 70)])
    assert read_input(path).settings.weekly_ratios == ((MONDAY, Decimal("70")),)


def test_bad_base_period_is_not_replaced_by_an_adjustment_period_error(tmp_path):
    path = _manual_adjustment(tmp_path, [("週", date(2026, 11, 16), 50)])
    _edit(path, lambda wb: wb["必要人数"].delete_rows(3))
    data = read_input(path)
    assert len(data.requirements) == 2


@pytest.mark.parametrize("row,cell,reason", (
    (("月", MONDAY, 70), "A2", "指定単位"),
    (("週", TUESDAY, 70), "B2", "月曜日"),
    (("日付", "11/03", 70), "B2", "年を含む"),
    (("日付", TUESDAY, None), "C2", "出勤率"),
    (("日付", TUESDAY, -1), "C2", "出勤率"),
    (("日付", TUESDAY, 101), "C2", "出勤率"),
    (("日付", TUESDAY, "NaN"), "C2", "出勤率"),
    (("日付", TUESDAY, True), "C2", "出勤率"),
))
def test_bad_adjustment_values_show_original_source_cell(tmp_path, row, cell, reason):
    path = _manual_adjustment(tmp_path, [row])
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.file == path.name and issue.source.sheet == ATTENDANCE_SHEET and issue.source.row == 2 and issue.source.column == cell and reason in issue.message for issue in caught.value.issues)


@pytest.mark.parametrize("cell", ("A2", "B2", "C2"))
def test_formulas_in_adjustment_values_are_not_cached_or_executed(tmp_path, cell):
    path = _manual_adjustment(tmp_path, [("週", MONDAY, 70)])
    _edit(path, lambda wb: setattr(wb[ATTENDANCE_SHEET][cell], "value", "=1+1"))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.kind == "数式" and issue.source.column == cell for issue in caught.value.issues)


def test_duplicate_adjustment_day_is_rejected_but_distinct_units_are_allowed(tmp_path):
    path = _manual_adjustment(tmp_path, [("週", MONDAY, 70), ("週", MONDAY, 50)])
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.column == "B3" and "重複" in issue.message for issue in caught.value.issues)


@pytest.mark.parametrize("value,number_format", ((0.7, "0%"), (70, r"0\%"), (70, '0"%"')))
def test_excel_percentage_format_has_its_explicit_percentage_meaning(tmp_path, value, number_format):
    path = _manual_adjustment(tmp_path, [("週", MONDAY, value)])
    def edit(wb):
        wb[ATTENDANCE_SHEET]["C2"].number_format = number_format
        wb[ATTENDANCE_SHEET]["D2"] = "=1+1"  # 使用しない列の数式は無関係。
    _edit(path, edit)
    assert read_input(path).settings.weekly_ratios == ((MONDAY, Decimal("70.0")),)


def test_daily_export_keeps_old_positions_and_appends_effective_ratio_source_and_staff(tmp_path):
    days = (MONDAY, TUESDAY, date(2026, 11, 4))
    data = InputData(
        tuple(Employee(employee_id, employee_id, count) for employee_id, count in zip(("A", "B", "C", "D"), (2, 2, 2, 1))),
        tuple(DayRequirement(day, 2) for day in days),
        settings=ShiftSettings(weekday_ratios=(Decimal("50"),) * 7, weekly_ratios=((MONDAY, Decimal("75")),), date_ratios=((TUESDAY, Decimal("25")),)),
    )
    data = configure_input(data)
    result = solve_schedule(data)
    assert result.status in (SolverStatus.OPTIMAL, SolverStatus.FEASIBLE)
    wb = load_workbook(export_result(data, result.schedule, result, tmp_path))
    ws = wb["日別人数"]
    headers = [cell.value for cell in ws[3]]
    assert headers[:5] == ["日付", "必要配置人数", "勤務人数", "差", "判定"]
    assert headers[-3:] == ["有効出勤率％", "指定根拠", "職員数"]
    values = list(ws.iter_rows(min_row=4, values_only=True))
    assert [row[1] for row in values] == [3, 1, 3]
    assert [row[-3:] for row in values] == [(75, "週指定", 4), (25, "日付指定", 4), (75, "週指定", 4)]
    metadata = dict(wb["設定条件"].iter_rows(min_row=4, values_only=True))
    assert "日付指定" in metadata["出勤率の適用順序"]
    assert "2026-11-02" in metadata["週別出勤率"]
    assert "2026-11-03" in metadata["日付別出勤率"]
    wb.close()


def test_paid_to_public_conversion_has_ninth_column_and_keeps_fixed_paid_and_total_days(tmp_path):
    path = export_independent_sample(tmp_path)
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    raw = read_input(path)
    # 別に確定している有給1件を追加。公休基準を1日減らし、休暇合計を保つ。
    fixed = PaidLeave("002", date(2026, 11, 11), Source("fixed.xlsx", "有給", 2))
    raw = replace(raw, employees=tuple(replace(employee, public_holidays=8) if employee.id == "002" else employee for employee in raw.employees), paid_leaves=(fixed,))
    paid_ids = {item.request_id for item in raw.independent_requests if item.leave_type == PAID_REQUEST}
    converted_id = next(item.request_id for item in raw.independent_requests if item.employee_id == "001" and item.leave_type == PAID_REQUEST)
    data = configure_input(set_leave_choices(raw, paid_ids - {converted_id}, {converted_id}))
    assert fixed in data.paid_leaves
    assert data.converted_public_request_ids == frozenset({converted_id})
    result = solve_schedule(data)
    assert result.schedule is not None
    folder = tmp_path / "results"
    folder.mkdir()
    wb = load_workbook(export_result(data, result.schedule, result, folder, input_path=path))
    ws = wb["希望休の採用結果"]
    assert ws.max_column == 9 and ws["I3"].value == "元の希望"
    converted = next(row for row in ws.iter_rows(min_row=4, values_only=True) if row[0] == converted_id)
    assert converted[7] == PUBLIC_REQUEST and converted[8] == PAID_REQUEST and converted[5] == 1
    metadata = dict(wb["設定条件"].iter_rows(min_row=4, values_only=True))
    assert metadata["公休基準/001"] == 9
    assert metadata["公休への振替加算/001"] == 1
    assert metadata["適用公休目標/001"] == 10
    assert metadata["休暇合計（公休＋有給）/001"] == 10
    assert metadata["元からの固定有給件数"] == 1
    assert json.loads(metadata["公休希望へ変更した申請ID"]) == [converted_id]
    assert "30日" in metadata["暦日合計/001"]
    wb.close()
    assert hashlib.sha256(path.read_bytes()).hexdigest() == digest
    restored = read_input(path)
    assert not restored.confirmed_paid_request_ids and not restored.converted_public_request_ids
    assert all(employee.public_holiday_adjustment == 0 for employee in restored.employees)
