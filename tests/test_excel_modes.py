"""交代制・曜日比率・前期間連勤のExcel/JSON入出力を検証する。"""
from dataclasses import replace
from datetime import date
from decimal import Decimal

import pytest
from openpyxl import load_workbook

from sinta_shift.excel_io import (
    PREVIOUS_DAYS, PREVIOUS_SHIFT, SETTINGS_SHEET, WEEKDAY_KEYS, export_result, export_template,
    read_input, settings_from_payload, settings_to_payload,
)
from sinta_shift.models import (
    DAY_ONLY, TWO_SHIFT, THREE_SHIFT, SHIFT_MODES, DayRequirement, Employee,
    InputData, InputError, ShiftSettings, SolverStatus,
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


@pytest.mark.parametrize("mode", (DAY_ONLY, TWO_SHIFT, THREE_SHIFT))
def test_settings_json_round_trip_preserves_decimal(mode):
    settings = ShiftSettings(mode, 4, tuple(Decimal("66.6666666666666666667") for _ in range(7)), (1,) * len(SHIFT_MODES[mode]))
    payload = settings_to_payload(settings)
    assert all(isinstance(value, str) for value in payload["weekday_ratios"])
    assert settings_from_payload(payload) == settings


@pytest.mark.parametrize("payload", (
    {"mode": "未対応"}, {"mode": []}, {"unknown": 1},
    {"max_consecutive_days": 0}, {"max_consecutive_days": True},
    {"weekday_ratios": ["50"] * 6}, {"weekday_ratios": ["101"] * 7},
    {"weekday_ratios": ["NaN"] * 7}, {"weekday_ratios": [0.5] * 7},
    {"shift_weights": [0]}, {"shift_weights": [1.5]},
    {"mode": TWO_SHIFT, "shift_weights": [1]},
    {"night_rest": "yes"}, {"forbidden_transitions": ["日勤"]},
    {"mode": TWO_SHIFT, "forbidden_transitions": [["準夜勤", "日勤"]]},
))
def test_malformed_settings_json_is_rejected(payload):
    with pytest.raises(InputError):
        settings_from_payload(payload)


@pytest.mark.parametrize("mode", (TWO_SHIFT, THREE_SHIFT))
def test_advanced_template_round_trip_has_conditions_and_explicit_counts(tmp_path, mode):
    settings = ShiftSettings(mode, 2, shift_weights=(1,) * len(SHIFT_MODES[mode]))
    path = export_template(tmp_path, settings=settings)
    data = read_input(path)
    assert data.settings == settings
    assert all(employee.previous_work_days == 0 for employee in data.employees)
    assert all(set(dict(r.shift_counts)) == set(SHIFT_MODES[mode]) for r in data.requirements)
    assert all(sum(dict(r.shift_counts).values()) == r.required for r in data.requirements)
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["職員", "必要人数", "確定休", "希望休", SETTINGS_SHEET]
    assert workbook["職員"].cell(1, 4).value == PREVIOUS_DAYS
    assert [workbook["必要人数"].cell(1, column).value for column in range(3, 3 + len(SHIFT_MODES[mode]))] == [f"{label}人数" for label in SHIFT_MODES[mode]]
    workbook.close()


def test_ratio_template_round_trip_and_decimal_configuration(tmp_path):
    ratios = tuple(Decimal("66.66") for _ in range(7))
    settings = ShiftSettings(THREE_SHIFT, 3, ratios, (2, 1, 1))
    path = export_template(tmp_path, settings=settings)
    data = configure_input(read_input(path))
    assert data.settings == settings
    assert [r.required for r in data.requirements] == [2, 2, 2]
    assert all(r.shift_counts == (("日勤", 1), ("準夜勤", 1), ("深夜勤", 0)) for r in data.requirements)


def test_night_and_transition_settings_round_trip_in_json_and_excel(tmp_path):
    settings = ShiftSettings(THREE_SHIFT, night_rest=False, forbidden_transitions=(("準夜勤", "日勤"), ("深夜勤", "準夜勤")))
    assert settings_from_payload(settings_to_payload(settings)) == settings
    path = export_template(tmp_path, settings=settings)
    assert read_input(path).settings == settings


@pytest.mark.parametrize("sheet,coordinate", (("職員", "D2"), ("職員", "E2"), ("必要人数", "C2"), (SETTINGS_SHEET, "B3")))
def test_used_optional_formulas_rejected_with_original_location(tmp_path, sheet, coordinate):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT, 3))
    _edit(path, lambda wb: setattr(wb[sheet][coordinate], "value", "=1+1"))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.kind == "数式" and issue.source.file == path.name and issue.source.sheet == sheet and issue.source.row == 2 + (sheet == SETTINGS_SHEET) and issue.source.column == coordinate for issue in caught.value.issues)


def test_missing_prior_column_and_blank_prior_are_the_documented_zero_assumption(tmp_path):
    legacy = read_input(export_template(tmp_path))
    assert all(employee.previous_work_days == 0 for employee in legacy.employees)
    advanced_path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT, 3))
    _edit(advanced_path, lambda wb: setattr(wb["職員"]["D2"], "value", None))
    data = read_input(advanced_path)
    assert data.employees[0].previous_work_days == 0


@pytest.mark.parametrize("value", (-1, "abc", 1.5, True))
def test_invalid_prior_count_is_not_changed_to_zero(tmp_path, value):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT, 3))
    _edit(path, lambda wb: setattr(wb["職員"]["D2"], "value", value))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.sheet == "職員" and issue.source.row == 2 and issue.source.column == "D2" for issue in caught.value.issues)


def test_optional_prior_and_shift_columns_can_be_mapped(tmp_path):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT, 3))
    def edit(wb):
        wb["職員"]["D1"] = "前月末連勤"
        wb["職員"]["D2"] = 2
        wb["職員"]["E1"] = "昨日の区分"
        wb["職員"]["E2"] = "日勤"
        wb["必要人数"]["C1"] = "昼の人数"
        wb["必要人数"]["D1"] = "夜の人数"
    _edit(path, edit)
    mappings = {
        "職員": {"columns": {PREVIOUS_DAYS: "前月末連勤", PREVIOUS_SHIFT: "昨日の区分"}},
        "必要人数": {"columns": {"日勤人数": "昼の人数", "夜勤人数": "夜の人数"}},
    }
    data = read_input(path, mappings)
    assert data.employees[0].previous_work_days == 2
    assert data.employees[0].previous_shift == "日勤"
    assert data.requirements[0].shift_counts == (("日勤", 1), ("夜勤", 1))


def test_active_shift_count_blank_is_not_zero(tmp_path):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))
    _edit(path, lambda wb: setattr(wb["必要人数"]["D2"], "value", None))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.column == "D2" for issue in caught.value.issues)


def test_missing_one_explicit_shift_count_cannot_silently_be_zero(tmp_path):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))
    _edit(path, lambda wb: wb["必要人数"].delete_cols(4))
    with pytest.raises(InputError):
        configure_input(read_input(path))


def test_explicit_unused_optional_columns_override_canonical_detection(tmp_path):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))
    _edit(path, lambda wb: setattr(wb["職員"]["D2"], "value", 2))
    data = read_input(path, {
        "職員": {"columns": {PREVIOUS_DAYS: "", PREVIOUS_SHIFT: ""}},
        "必要人数": {"columns": {"日勤人数": "", "夜勤人数": "", "準夜勤人数": "", "深夜勤人数": ""}},
    })
    assert all(employee.previous_work_days == 0 and employee.previous_shift is None for employee in data.employees)
    assert all(requirement.shift_counts == () for requirement in data.requirements)


def test_explicit_unused_optional_formula_is_not_read(tmp_path):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))
    def edit(wb):
        wb["職員"]["D2"] = "=1+1"
        wb["職員"]["E2"] = '=IF(1,"夜勤","休み")'
        wb["必要人数"]["C2"] = "=1"
        wb["必要人数"]["D2"] = "=1"
    _edit(path, edit)
    with pytest.raises(InputError):
        read_input(path)
    data = read_input(path, {
        "職員": {"columns": {PREVIOUS_DAYS: "", PREVIOUS_SHIFT: ""}},
        "必要人数": {"columns": {"日勤人数": "", "夜勤人数": ""}},
    })
    assert data.employees[0].previous_work_days == 0
    assert data.employees[0].previous_shift is None
    assert data.requirements[0].shift_counts == ()


def test_empty_optional_mapping_is_valid_when_columns_are_absent(tmp_path):
    path = export_template(tmp_path)
    data = read_input(path, {
        "職員": {"columns": {PREVIOUS_DAYS: "", PREVIOUS_SHIFT: ""}},
        "必要人数": {"columns": {f"{label}人数": "" for label in ("日勤", "夜勤", "準夜勤", "深夜勤")}},
    })
    assert len(data.employees) == 3
    assert all(r.shift_counts == () for r in data.requirements)


def test_explicit_canonical_optional_mapping_reads_the_source_values(tmp_path):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))
    def edit(wb):
        wb["職員"]["D2"] = 2
        wb["職員"]["E2"] = "夜勤"
    _edit(path, edit)
    data = read_input(path, {
        "職員": {"columns": {PREVIOUS_DAYS: PREVIOUS_DAYS, PREVIOUS_SHIFT: PREVIOUS_SHIFT}},
        "必要人数": {"columns": {"日勤人数": "日勤人数", "夜勤人数": "夜勤人数"}},
    })
    assert data.employees[0].previous_work_days == 2
    assert data.employees[0].previous_shift == "夜勤"
    assert data.requirements[0].shift_counts == (("日勤", 1), ("夜勤", 1))


def test_condition_sheet_invalid_ratio_has_exact_source_row(tmp_path):
    settings = ShiftSettings(TWO_SHIFT, 3, tuple(Decimal("66") for _ in range(7)))
    path = export_template(tmp_path, settings=settings)
    def edit(wb):
        ws = wb[SETTINGS_SHEET]
        for row in ws:
            if row[0].value == WEEKDAY_KEYS[2]:
                row[1].value = -5
    _edit(path, edit)
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.sheet == SETTINGS_SHEET and issue.source.row == 7 and issue.source.column == "B7" for issue in caught.value.issues)


@pytest.mark.parametrize("mode", (TWO_SHIFT, THREE_SHIFT))
def test_multi_shift_export_has_actual_labels_counts_and_conditions(tmp_path, mode):
    settings = ShiftSettings(mode, 2, shift_weights=(1,) * len(SHIFT_MODES[mode]))
    data = configure_input(read_input(export_template(tmp_path, settings=settings)))
    result = solve_schedule(data)
    assert result.status in (SolverStatus.OPTIMAL, SolverStatus.FEASIBLE)
    output = export_result(data, result.schedule, result, tmp_path)
    workbook = load_workbook(output, data_only=False)
    assert workbook.sheetnames == ["勤務表", "日別人数", "希望休の採用結果", "設定条件"]
    assert all(cell.data_type != "f" for ws in workbook for row in ws for cell in row)
    labels = {workbook["勤務表"].cell(row, column).value for row in range(4, 7) for column in range(3, 6)}
    assert labels <= {*SHIFT_MODES[mode], "休み"}
    assert "休み" in labels
    headers = {cell.value: cell.column for cell in workbook["日別人数"][3]}
    for label in SHIFT_MODES[mode]:
        assert f"{label}人数" in headers
        assert f"{label}必要人数" in headers
        for row in range(4, 7):
            assert workbook["日別人数"].cell(row, headers[f"{label}人数"]).value == workbook["日別人数"].cell(row, headers[f"{label}必要人数"]).value
    table_headers = [cell.value for cell in workbook["勤務表"][3]]
    assert all(f"{label}日数" in table_headers for label in SHIFT_MODES[mode])
    metadata = dict(workbook["設定条件"].iter_rows(min_row=4, values_only=True))
    assert metadata["勤務区分"] == mode
    assert metadata["連続勤務上限"] == 2
    assert "前日は休み" in metadata["前期間の扱い"]
    assert "夜勤明け" in metadata["対象外"]
    workbook.close()


def test_export_rejects_prior_period_consecutive_violation(tmp_path):
    first, second = date(2026, 11, 7), date(2026, 11, 8)
    data = configure_input(InputData(
        (Employee("A", "職員A", 1, previous_work_days=1, previous_shift="日勤"), Employee("B", "職員B", 1)),
        (DayRequirement(first, 1), DayRequirement(second, 1)),
        settings=ShiftSettings(TWO_SHIFT, 1),
    ))
    result = solve_schedule(data)
    assert result.schedule is not None
    broken = {("A", first): "日勤", ("A", second): "休み", ("B", first): "休み", ("B", second): "日勤"}
    with pytest.raises(InputError) as caught:
        export_result(data, broken, result, tmp_path)
    assert any("連続勤務" in issue.message for issue in caught.value.issues)
    assert not list(tmp_path.glob("sinta_shift_result*"))


def test_export_rejects_wrong_shift_even_when_total_headcount_is_unchanged(tmp_path):
    data = configure_input(read_input(export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))))
    result = solve_schedule(data)
    broken = dict(result.schedule)
    key = next(key for key, value in broken.items() if value == "夜勤")
    broken[key] = "日勤"
    with pytest.raises(InputError):
        export_result(data, broken, result, tmp_path)
    assert not list(tmp_path.glob("sinta_shift_result*"))


@pytest.mark.parametrize("previous", ("遅番", True, 123))
def test_invalid_previous_shift_reports_original_cell(tmp_path, previous):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))
    _edit(path, lambda wb: setattr(wb["職員"]["E2"], "value", previous))
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.column == "E2" and issue.source.row == 2 for issue in caught.value.issues)


@pytest.mark.parametrize("key,value", (("夜勤後休み", "未定"), ("禁止勤務順序", "invalid json"), ("禁止勤務順序", '[["準夜勤","日勤"]]')))
def test_invalid_night_conditions_report_settings_source(tmp_path, key, value):
    path = export_template(tmp_path, settings=ShiftSettings(TWO_SHIFT))
    expected_row = []
    def edit(wb):
        for row in wb[SETTINGS_SHEET]:
            if row[0].value == key:
                expected_row.append(row[1].row)
                row[1].value = value
    _edit(path, edit)
    with pytest.raises(InputError) as caught:
        read_input(path)
    assert any(issue.source.sheet == SETTINGS_SHEET and issue.source.row == expected_row[0] for issue in caught.value.issues)


@pytest.mark.parametrize("mode,first_shift,forbidden,reason", (
    (TWO_SHIFT, "夜勤", (), "夜勤"),
    (THREE_SHIFT, "準夜勤", (("準夜勤", "日勤"),), "禁止"),
))
def test_export_blocks_night_rest_and_forbidden_transition_without_count_violations(tmp_path, mode, first_shift, forbidden, reason):
    days = tuple(date(2026, 11, day) for day in (7, 8, 9))
    active = SHIFT_MODES[mode]
    requirements = tuple(DayRequirement(day, 1, shift_counts=tuple((label, int(label == (first_shift if index == 0 else "日勤"))) for label in active)) for index, day in enumerate(days))
    data = configure_input(InputData(
        (Employee("A", "職員A", 2), Employee("B", "職員B", 1)), requirements,
        settings=ShiftSettings(mode, forbidden_transitions=forbidden),
    ))
    result = solve_schedule(data)
    assert result.schedule is not None
    broken = {
        ("A", days[0]): first_shift, ("A", days[1]): "日勤", ("A", days[2]): "休み",
        ("B", days[0]): "休み", ("B", days[1]): "休み", ("B", days[2]): "日勤",
    }
    with pytest.raises(InputError) as caught:
        export_result(data, broken, result, tmp_path)
    assert any(reason in issue.message for issue in caught.value.issues)
    assert not list(tmp_path.glob("sinta_shift_result*"))


def test_final_night_export_marks_next_day_outside_period_and_employee(tmp_path):
    day = date(2026, 11, 7)
    data = configure_input(InputData(
        (Employee("A", "職員A", 1),),
        (DayRequirement(day, 1, shift_counts=(("日勤", 0), ("夜勤", 1))),),
        settings=ShiftSettings(TWO_SHIFT),
    ))
    result = solve_schedule(data)
    assert result.schedule == {("A", day): "夜勤"}
    workbook = load_workbook(export_result(data, result.schedule, result, tmp_path))
    metadata = dict(workbook["設定条件"].iter_rows(min_row=4, values_only=True))
    assert "未確認" in metadata["期間外の翌日確認"]
    assert metadata["翌日休み確認対象の職員ID"] == "A"
    assert metadata["翌日確認対象日"].date() == date(2026, 11, 8)
    assert any("期間外未確認" in str(cell.value) and "2026-11-08" in str(cell.value) for row in workbook["勤務表"] for cell in row)
    workbook.close()
