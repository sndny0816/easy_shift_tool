"""勤務区分・曜日比率・連勤を、GUIとExcelを使わず検証する。"""
from dataclasses import replace
from datetime import date, timedelta
from decimal import Decimal

import pytest

from sinta_shift.evaluation import evaluate_schedule
from sinta_shift.models import (
    DAY_ONLY, TWO_SHIFT, THREE_SHIFT, SHIFT_MODES, OFF, DayRequirement,
    Employee, FixedOff, InputData, InputError, ShiftSettings, SolverStatus, Source,
    is_working,
)
from sinta_shift.planning import allocate_shifts, configure_input, required_from_ratio
from sinta_shift.sample import sample_data
from sinta_shift.solver import solve_schedule
from sinta_shift.validation import validate_input


START = date(2026, 11, 7)


def _days(n):
    return tuple(START + timedelta(days=i) for i in range(n))


def _multi_data(mode):
    return InputData(
        employees=tuple(Employee(s, f"職員{s}", 3) for s in "ABCD"),
        requirements=tuple(DayRequirement(day, 3) for day in _days(4)),
        settings=ShiftSettings(mode=mode),
    )


@pytest.mark.parametrize("count,pct,expected", [
    (30, "66", 20), (30, "0", 0), (30, "100", 30), (1, "0.1", 1),
    (3, "33.3333333333333333333333333333333333333333333333333333333", 1),
    (3, "33.3333333333333333333333333333333333333333333333333333334", 2),
])
def test_required_headcount_is_exact_ceiling(count, pct, expected):
    assert required_from_ratio(count, Decimal(pct)) == expected


@pytest.mark.parametrize("total,mode,weights,expected", [
    (7, THREE_SHIFT, (), (("日勤", 3), ("準夜勤", 2), ("深夜勤", 2))),
    (7, TWO_SHIFT, (), (("日勤", 4), ("夜勤", 3))),
    (5, TWO_SHIFT, (2, 1), (("日勤", 3), ("夜勤", 2))),
    (4, THREE_SHIFT, (0, 1, 0), (("日勤", 0), ("準夜勤", 4), ("深夜勤", 0))),
    (0, THREE_SHIFT, (), (("日勤", 0), ("準夜勤", 0), ("深夜勤", 0))),
])
def test_largest_remainder_allocation(total, mode, weights, expected):
    assert allocate_shifts(total, mode, weights) == expected
    assert sum(n for _, n in expected) == total


@pytest.mark.parametrize("weights", [(0, 0), (-1, 2), (True, 1), (1.5, 1), (1,), (Decimal("NaN"), 1), (float("inf"), 1)])
def test_invalid_shift_weights_rejected(weights):
    with pytest.raises(InputError):
        allocate_shifts(3, TWO_SHIFT, weights)


@pytest.mark.parametrize("pct", [Decimal("-0.1"), Decimal("100.1"), Decimal("NaN"), Decimal("Infinity"), 66.0])
def test_invalid_percentages_rejected(pct):
    with pytest.raises(InputError):
        required_from_ratio(30, pct)


def test_ratios_use_weekday_order_preserve_sources_and_leave_workdays_unchanged():
    ratios = (Decimal(100), Decimal(0), Decimal(0), Decimal(0), Decimal(0), Decimal(50), Decimal(0))
    requirements = tuple(DayRequirement(day, 1, Source("input.xlsx", "配置", i + 2)) for i, day in enumerate(_days(3)))
    data = InputData(tuple(Employee(f"E{i:02}", f"職員{i}", 0) for i in range(30)), requirements)
    settings = ShiftSettings(mode=TWO_SHIFT, weekday_ratios=ratios)
    configured = configure_input(data, settings)
    assert [r.required for r in configured.requirements] == [15, 0, 30]  # 土・日・月
    assert configured.employees == data.employees
    assert [r.source for r in configured.requirements] == [r.source for r in data.requirements]
    assert [r.required for r in data.requirements] == [1, 1, 1]
    assert configure_input(configured) == configured
    assert any("合計" in issue.message for issue in validate_input(configured))


def test_ratio_change_reports_total_mismatch_without_adjusting_individual_days():
    data = configure_input(sample_data(), ShiftSettings(weekday_ratios=(Decimal(100),) * 7))
    assert [e.work_days for e in data.employees] == [2, 2, 2]
    result = solve_schedule(data)
    assert result.status is SolverStatus.INFEASIBLE
    assert result.schedule is None
    assert "勤務日数合計6人日と配置人数合計9人日" in result.message


@pytest.mark.parametrize("settings", [
    ShiftSettings(mode="交代"), ShiftSettings(max_consecutive_days=True),
    ShiftSettings(max_consecutive_days=0), ShiftSettings(max_consecutive_days=-1),
    ShiftSettings(max_consecutive_days=1.5), ShiftSettings(weekday_ratios=(Decimal(66),) * 6),
    ShiftSettings(weekday_ratios=(Decimal("NaN"),) * 7),
])
def test_invalid_settings_produce_actionable_input_errors(settings):
    data = replace(sample_data(), settings=settings)
    assert validate_input(data)
    with pytest.raises(InputError):
        configure_input(data)
    assert solve_schedule(data).schedule is None


def test_mode_alias_normalizes_to_public_mode():
    assert configure_input(sample_data(), ShiftSettings(mode="2交代制")).settings.mode == TWO_SHIFT
    assert configure_input(sample_data(), ShiftSettings(mode="3交代制")).settings.mode == THREE_SHIFT


def test_explicit_counts_remain_exact_without_ratio_override():
    data = _multi_data(THREE_SHIFT)
    counts = (("深夜勤", 2), ("日勤", 0), ("準夜勤", 1))
    data = replace(data, requirements=tuple(replace(r, shift_counts=counts) for r in data.requirements))
    configured = configure_input(data)
    expected = (("日勤", 0), ("準夜勤", 1), ("深夜勤", 2))
    assert all(r.shift_counts == expected for r in configured.requirements)
    assert configure_input(configured) == configured
    # 比率を明示すると、以前の内訳も比率と配分から更新される。
    overridden = configure_input(data, ShiftSettings(mode=THREE_SHIFT, weekday_ratios=(Decimal(75),) * 7))
    assert all(r.shift_counts == (("日勤", 1), ("準夜勤", 1), ("深夜勤", 1)) for r in overridden.requirements)


@pytest.mark.parametrize("counts", [
    (("日勤", 3),), (("日勤", 1), ("日勤", 1), ("深夜勤", 1)),
    (("日勤", 1), ("夜勤", 1), ("深夜勤", 1)),
    (("日勤", -1), ("準夜勤", 2), ("深夜勤", 2)),
    (("日勤", 1), ("準夜勤", 1), ("深夜勤", 0)),
])
def test_explicit_shift_counts_require_complete_valid_and_balanced_numbers(counts):
    data = _multi_data(THREE_SHIFT)
    data = replace(data, requirements=(replace(data.requirements[0], shift_counts=counts), *data.requirements[1:]))
    assert validate_input(data)
    with pytest.raises(InputError):
        configure_input(data)


@pytest.mark.parametrize("mode", [DAY_ONLY, TWO_SHIFT, THREE_SHIFT])
def test_solver_satisfies_every_active_shift_and_workday_total(mode):
    data = configure_input(_multi_data(mode))
    result = solve_schedule(data)
    assert result.status is SolverStatus.OPTIMAL
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid, [i.message for i in evaluation.issues]
    for requirement in data.requirements:
        assert evaluation.daily_counts[requirement.day] == requirement.required
        assert evaluation.daily_shift_counts[requirement.day] == dict(requirement.shift_counts)
    for employee in data.employees:
        assert evaluation.employee_counts[employee.id] == employee.work_days
        assert sum(evaluation.employee_shift_counts[employee.id].values()) == employee.work_days
    if mode == DAY_ONLY:
        assert all(type(value) is bool for value in result.schedule.values())
    else:
        assert all(type(value) is str and value in (*SHIFT_MODES[mode], OFF) for value in result.schedule.values())


@pytest.mark.parametrize("mode", [TWO_SHIFT, THREE_SHIFT])
def test_leave_requests_keep_lexicographic_semantics_in_shift_modes(mode):
    data = configure_input(sample_data(), ShiftSettings(mode=mode))
    result = solve_schedule(data)
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid and evaluation.adopted_count == 3
    adopted = {r.request_id: r for r in evaluation.request_results}
    assert adopted["R2"].priority == 2
    assert result.schedule["A", START] == OFF


def test_changing_only_shift_label_detects_violation_even_when_workdays_still_match():
    data = configure_input(_multi_data(THREE_SHIFT))
    schedule = dict(solve_schedule(data).schedule)
    key = next(key for key, value in schedule.items() if value == "日勤")
    schedule[key] = "深夜勤"
    evaluation = evaluate_schedule(data, schedule)
    assert not evaluation.valid
    assert all(evaluation.daily_counts[r.day] == r.required for r in data.requirements)
    assert all(evaluation.employee_counts[e.id] == e.work_days for e in data.employees)
    assert any(i.kind == "勤務区分人数違反" and i.day == key[1] for i in evaluation.issues)
    assert any(i.employee_id == key[0] and i.day == key[1] and i.source.column for i in evaluation.issues)


def test_boolean_cells_are_not_accepted_in_ambiguous_shift_mode():
    data = configure_input(sample_data(), ShiftSettings(mode=TWO_SHIFT))
    schedule = dict(solve_schedule(data).schedule)
    key = next(iter(schedule))
    schedule[key] = False
    evaluation = evaluate_schedule(data, schedule)
    assert not evaluation.valid
    assert any(i.kind == "区分違反" and i.employee_id == key[0] and i.day == key[1] for i in evaluation.issues)


def test_consecutive_limit_counts_all_work_shifts_and_includes_previous_period():
    days = _days(4)
    data = InputData(
        (Employee("A", "職員A", 2, previous_work_days=1, previous_shift="日勤"), Employee("B", "職員B", 2)),
        tuple(DayRequirement(day, 1) for day in days),
        settings=ShiftSettings(mode=TWO_SHIFT, max_consecutive_days=1, shift_weights=(0, 1)),
    )
    result = solve_schedule(data)
    assert result.status is SolverStatus.OPTIMAL
    assert result.schedule["A", days[0]] == OFF
    assert all(value in (OFF, "夜勤") for value in result.schedule.values())
    assert evaluate_schedule(data, result.schedule).valid
    for employee in data.employees:
        run = employee.previous_work_days
        for day in days:
            run = run + 1 if is_working(result.schedule[employee.id, day], TWO_SHIFT) else 0
            assert run <= 1


def test_manual_continuous_shift_run_marks_exact_employee_date_and_cell():
    days = _days(4)
    data = InputData(
        (Employee("A", "職員A", 2), Employee("B", "職員B", 2)),
        tuple(DayRequirement(day, 1, shift_counts=(("日勤", 1 if i % 2 == 0 else 0), ("夜勤", 0 if i % 2 == 0 else 1))) for i, day in enumerate(days)),
        settings=ShiftSettings(mode=TWO_SHIFT, max_consecutive_days=1),
    )
    schedule = {("A", day): ("日勤" if i == 0 else "夜勤" if i == 1 else OFF) for i, day in enumerate(days)}
    schedule.update({("B", day): ("日勤" if i == 2 else "夜勤" if i == 3 else OFF) for i, day in enumerate(days)})
    evaluation = evaluate_schedule(data, schedule)
    assert not evaluation.valid
    assert all(evaluation.daily_counts[day] == 1 for day in days)
    violations = [i for i in evaluation.issues if i.kind == "連続勤務違反"]
    assert {(i.employee_id, i.day, i.source.column) for i in violations} == {("A", days[1], "D4"), ("B", days[3], "F5")}
    assert all("連続勤務2日" in i.message and "上限1日" in i.message for i in violations)


def test_previous_period_boundary_can_make_otherwise_balanced_plan_infeasible():
    day = START
    data = InputData(
        (Employee("A", "職員A", 1, previous_work_days=2),),
        (DayRequirement(day, 1),),
        settings=ShiftSettings(max_consecutive_days=2),
    )
    assert validate_input(data) == []
    assert solve_schedule(data).status is SolverStatus.INFEASIBLE
    evaluation = evaluate_schedule(data, {("A", day): True})
    assert not evaluation.valid
    assert any(i.kind == "連続勤務違反" and i.employee_id == "A" and i.day == day and i.source.column == "C4" for i in evaluation.issues)
    allowed = replace(data, employees=(replace(data.employees[0], previous_work_days=1),))
    assert evaluate_schedule(allowed, solve_schedule(allowed).schedule).valid


@pytest.mark.parametrize("previous,limit", [(-1, None), (True, None), (1.5, None), (3, 2)])
def test_invalid_previous_history_is_reported(previous, limit):
    data = sample_data()
    data = replace(data, employees=(replace(data.employees[0], previous_work_days=previous), *data.employees[1:]), settings=ShiftSettings(max_consecutive_days=limit))
    assert any(i.employee_id == "A" and "期間開始前" in i.message for i in validate_input(data))


def test_night_rest_and_transitions_can_be_explicitly_disabled():
    days = _days(2)
    data = InputData(
        (Employee("A", "職員A", 2), Employee("B", "職員B", 2)),
        tuple(DayRequirement(day, 2) for day in days),
        settings=ShiftSettings(mode=TWO_SHIFT, max_consecutive_days=2, night_rest=False),
    )
    schedule = {("A", days[0]): "夜勤", ("A", days[1]): "日勤", ("B", days[0]): "日勤", ("B", days[1]): "夜勤"}
    assert evaluate_schedule(data, schedule).valid


def test_zero_counts_for_inactive_known_shifts_are_ignored_but_not_positive_counts():
    data = _multi_data(DAY_ONLY)
    counts = (("日勤", 3), ("夜勤", 0), ("準夜勤", 0), ("深夜勤", 0))
    data = replace(data, requirements=tuple(replace(r, shift_counts=counts) for r in data.requirements))
    assert all(r.shift_counts == (("日勤", 3),) for r in configure_input(data).requirements)
    wrong = replace(data, requirements=(replace(data.requirements[0], shift_counts=(("日勤", 2), ("夜勤", 1))), *data.requirements[1:]))
    assert validate_input(wrong)


@pytest.mark.parametrize("mode,night,weights", [(TWO_SHIFT, "夜勤", (0, 1)), (THREE_SHIFT, "深夜勤", (0, 0, 1))])
def test_night_shift_requires_next_day_off_for_every_employee(mode, night, weights):
    days = _days(3)
    data = InputData(
        (Employee("A", "職員A", 2), Employee("B", "職員B", 1)),
        tuple(DayRequirement(day, 1) for day in days),
        settings=ShiftSettings(mode=mode, shift_weights=weights, night_rest=True),
    )
    result = solve_schedule(data)
    assert result.status is SolverStatus.OPTIMAL
    assert evaluate_schedule(data, result.schedule).valid
    assert result.schedule["A", days[0]] == result.schedule["A", days[2]] == night
    assert result.schedule["A", days[1]] == OFF
    assert result.schedule["B", days[1]] == night
    for employee in data.employees:
        for i, day in enumerate(days[:-1]):
            if result.schedule[employee.id, day] == night:
                assert result.schedule[employee.id, days[i + 1]] == OFF


@pytest.mark.parametrize("mode,night,following", [(TWO_SHIFT, "夜勤", "日勤"), (THREE_SHIFT, "深夜勤", "準夜勤")])
def test_manual_night_rest_violation_marks_both_cells_with_balanced_totals(mode, night, following):
    days = _days(4)
    labels = SHIFT_MODES[mode]
    requirements = tuple(DayRequirement(day, 1, shift_counts=tuple((label, int(label == (night if i % 2 == 0 else following))) for label in labels)) for i, day in enumerate(days))
    data = InputData((Employee("A", "職員A", 2), Employee("B", "職員B", 2)), requirements, settings=ShiftSettings(mode=mode))
    assert evaluate_schedule(data, solve_schedule(data).schedule).valid
    schedule = {("A", day): (night if i == 0 else following if i == 1 else OFF) for i, day in enumerate(days)}
    schedule.update({("B", day): (night if i == 2 else following if i == 3 else OFF) for i, day in enumerate(days)})
    evaluation = evaluate_schedule(data, schedule)
    assert not evaluation.valid
    assert all(evaluation.daily_counts[day] == 1 for day in days)
    assert all(evaluation.employee_counts[e.id] == 2 for e in data.employees)
    cells = {(i.employee_id, i.day, i.source.column) for i in evaluation.issues if i.kind == "夜勤翌日違反"}
    assert cells == {("A", days[0], "C4"), ("A", days[1], "D4"), ("B", days[2], "E5"), ("B", days[3], "F5")}


def test_forbidden_evening_to_day_transition_is_constrained_and_rechecked():
    days = _days(4)
    labels = SHIFT_MODES[THREE_SHIFT]
    requirements = tuple(DayRequirement(day, 1, shift_counts=tuple((label, int(label == ("準夜勤" if i % 2 == 0 else "日勤"))) for label in labels)) for i, day in enumerate(days))
    data = InputData(
        (Employee("A", "職員A", 2), Employee("B", "職員B", 2)), requirements,
        settings=ShiftSettings(mode=THREE_SHIFT, forbidden_transitions=(("準夜勤", "日勤"),)),
    )
    result = solve_schedule(data)
    assert result.status is SolverStatus.OPTIMAL
    assert evaluate_schedule(data, result.schedule).valid
    schedule = {("A", day): ("準夜勤" if i == 0 else "日勤" if i == 1 else OFF) for i, day in enumerate(days)}
    schedule.update({("B", day): ("準夜勤" if i == 2 else "日勤" if i == 3 else OFF) for i, day in enumerate(days)})
    evaluation = evaluate_schedule(data, schedule)
    assert not evaluation.valid
    assert {(i.employee_id, i.day) for i in evaluation.issues if i.kind == "勤務順序違反"} == {("A", days[0]), ("A", days[1]), ("B", days[2]), ("B", days[3])}
    disabled = replace(data, settings=replace(data.settings, forbidden_transitions=()))
    assert evaluate_schedule(disabled, schedule).valid


def test_previous_night_requires_first_day_off_even_after_switching_to_day_only():
    data = InputData(
        (Employee("A", "職員A", 1, previous_work_days=1, previous_shift="深夜勤"),),
        (DayRequirement(START, 1),), settings=ShiftSettings(mode=DAY_ONLY),
    )
    assert validate_input(data) == []
    assert solve_schedule(data).status is SolverStatus.INFEASIBLE
    evaluation = evaluate_schedule(data, {("A", START): True})
    assert any(i.kind == "夜勤翌日違反" and i.day == START and i.source.column == "C4" and "境界" in i.message for i in evaluation.issues)
    disabled = replace(data, settings=replace(data.settings, night_rest=False))
    assert evaluate_schedule(disabled, solve_schedule(disabled).schedule).valid


def test_previous_evening_forbids_first_day_day_shift():
    data = InputData(
        (Employee("A", "職員A", 1, previous_work_days=1, previous_shift="準夜勤"),),
        (DayRequirement(START, 1, shift_counts=(("日勤", 1), ("準夜勤", 0), ("深夜勤", 0))),),
        settings=ShiftSettings(mode=THREE_SHIFT, forbidden_transitions=(("準夜勤", "日勤"),)),
    )
    assert validate_input(data) == []
    assert solve_schedule(data).status is SolverStatus.INFEASIBLE
    evaluation = evaluate_schedule(data, {("A", START): "日勤"})
    assert any(i.kind == "勤務順序違反" and i.day == START and i.source.column == "C4" for i in evaluation.issues)


@pytest.mark.parametrize("previous,shift,mode", [(0, "夜勤", TWO_SHIFT), (1, OFF, TWO_SHIFT), (1, None, TWO_SHIFT), (1, "不明", DAY_ONLY)])
def test_previous_shift_must_be_known_and_consistent_with_working_run(previous, shift, mode):
    data = replace(sample_data(), settings=ShiftSettings(mode=mode))
    data = replace(data, employees=(replace(data.employees[0], previous_work_days=previous, previous_shift=shift), *data.employees[1:]))
    assert any(i.employee_id == "A" and "期間開始前" in i.message for i in validate_input(data))


@pytest.mark.parametrize("settings", [
    ShiftSettings(night_rest=1),
    ShiftSettings(mode=TWO_SHIFT, forbidden_transitions=(("準夜勤", "日勤"),)),
    ShiftSettings(mode=THREE_SHIFT, forbidden_transitions=(("準夜勤", "日勤"), ("準夜勤", "日勤"))),
    ShiftSettings(forbidden_transitions=(("日勤",),)),
])
def test_invalid_rest_or_transition_settings_are_rejected(settings):
    with pytest.raises(InputError):
        configure_input(sample_data(), settings)


@pytest.mark.parametrize("mode,night,weights", [(TWO_SHIFT, "夜勤", (0, 1)), (THREE_SHIFT, "深夜勤", (0, 0, 1))])
def test_final_day_night_is_valid_but_following_day_is_outside_this_plan(mode, night, weights):
    data = InputData((Employee("A", "職員A", 1),), (DayRequirement(START, 1),), settings=ShiftSettings(mode=mode, shift_weights=weights))
    result = solve_schedule(data)
    assert result.status is SolverStatus.OPTIMAL
    assert result.schedule["A", START] == night
    assert evaluate_schedule(data, result.schedule).valid


@pytest.mark.parametrize("mode", [DAY_ONLY, TWO_SHIFT, THREE_SHIFT])
def test_zero_required_staff_is_supported_for_every_mode(mode):
    data = InputData((Employee("A", "職員A", 0),), tuple(DayRequirement(day, 0) for day in _days(3)), settings=ShiftSettings(mode=mode))
    result = solve_schedule(data)
    assert result.status is SolverStatus.OPTIMAL
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid
    assert all(count == 0 for count in evaluation.daily_counts.values())
    assert all(count == 0 for per_day in evaluation.daily_shift_counts.values() for count in per_day.values())
