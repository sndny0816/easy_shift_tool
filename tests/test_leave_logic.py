"""公休・有給・明けの期待例。旧区分との混同や勤務回数の水増しを検証。"""
from dataclasses import replace
from datetime import date, timedelta

import pytest

from sinta_shift.models import (
    DAY_ONLY, TWO_SHIFT, THREE_SHIFT, DIRECT_DAYS, PUBLIC_DAYS,
    OFF, PUBLIC_OFF, PAID_LEAVE, RECOVERY, Employee, DayRequirement, FixedOff,
    PaidLeave, LeaveCandidate, InputData, ShiftSettings, ShiftTime, SolverStatus,
    InputError, Source, assignment_label, assignment_labels, is_working, is_actual_working,
)
from sinta_shift.planning import configure_input, crossing_shifts, public_holiday_target
from sinta_shift.validation import validate_input
from sinta_shift.solver import solve_schedule
from sinta_shift.evaluation import evaluate_schedule
from sinta_shift.sample import sample_data


START = date(2026, 11, 1)
DAY_TIME = ShiftTime("日勤", 9 * 60, 17 * 60)
NIGHT_TIME = ShiftTime("夜勤", 20 * 60, 32 * 60)


def _days(n):
    return tuple(START + timedelta(days=i) for i in range(n))


def _assert_valid(data, result):
    assert result.status in (SolverStatus.OPTIMAL, SolverStatus.FEASIBLE), result.message
    assert result.schedule is not None
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid, [issue.message for issue in evaluation.issues]
    return evaluation


def _night_example(*, direct=False):
    """期待: 夜勤→明け→公休→有給。4セルの役割がすべて異なる。"""
    days = _days(4)
    return InputData(
        employees=(Employee("A", "職員A", 1 if direct else None, public_holidays=1),),
        requirements=tuple(DayRequirement(day, int(i == 0), shift_counts=(("日勤", 0), ("夜勤", int(i == 0)))) for i, day in enumerate(days)),
        fixed_off=(FixedOff("A", days[2]),),
        paid_leaves=(PaidLeave("A", days[3], Source("input.xlsx", "有給", 2)),),
        settings=ShiftSettings(mode=TWO_SHIFT, leave_enabled=True,
            day_count_mode=DIRECT_DAYS if direct else PUBLIC_DAYS,
            default_public_holidays=3, shift_times=(DAY_TIME, NIGHT_TIME), shift_times_confirmed=True),
    )


def test_thirty_days_eight_public_two_paid_means_twenty_actual_starts():
    days = _days(30)
    data = InputData(
        employees=(Employee("A", "職員A", None),),
        requirements=tuple(DayRequirement(day, int(8 <= i < 28)) for i, day in enumerate(days)),
        paid_leaves=(PaidLeave("A", days[28]), PaidLeave("A", days[29])),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, default_public_holidays=8),
    )
    effective = configure_input(data)
    assert effective.employees[0].work_days == 20
    assert configure_input(effective) == effective
    result = solve_schedule(effective)
    ev = _assert_valid(effective, result)
    expected = [PUBLIC_OFF] * 8 + ["日勤"] * 20 + [PAID_LEAVE] * 2
    assert [result.schedule["A", day] for day in days] == expected
    assert ev.employee_counts == ev.employee_actual_work_days == {"A": 20}
    assert ev.employee_public_counts == {"A": 8}
    assert ev.employee_paid_counts == {"A": 2}
    assert ev.employee_recovery_counts == {"A": 0}


def test_individual_zero_public_target_is_not_replaced_with_common_value():
    days = _days(3)
    data = InputData(
        (Employee("A", "職員A", None, public_holidays=0),),
        tuple(DayRequirement(day, int(i < 2)) for i, day in enumerate(days)),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, default_public_holidays=8),
        paid_leaves=(PaidLeave("A", days[2]),),
    )
    effective = configure_input(data)
    assert public_holiday_target(effective.employees[0], effective.settings) == 0
    assert effective.employees[0].work_days == 2
    assert _assert_valid(effective, solve_schedule(effective)).employee_public_counts == {"A": 0}


def test_direct_count_is_independent_of_inactive_public_target():
    days = _days(4)
    data = InputData(
        (Employee("A", "職員A", 2, public_holidays=3),),
        tuple(DayRequirement(day, int(i < 2)) for i, day in enumerate(days)),
        settings=ShiftSettings(leave_enabled=True, default_public_holidays=3),
        paid_leaves=(PaidLeave("A", days[3]),),
    )
    assert public_holiday_target(data.employees[0], data.settings) is None
    ev = _assert_valid(data, solve_schedule(data))
    assert ev.employee_counts == {"A": 2}
    assert ev.employee_public_counts == {"A": 1}
    assert ev.employee_paid_counts == {"A": 1}


@pytest.mark.parametrize("direct", [False, True])
def test_night_recovery_public_and_paid_are_exclusive_distinct_counts(direct):
    data = configure_input(_night_example(direct=direct))
    assert data.employees[0].work_days == (1 if direct else None)
    assert configure_input(data) == data
    result = solve_schedule(data)
    ev = _assert_valid(data, result)
    assert [result.schedule["A", day] for day in data.days] == ["夜勤", RECOVERY, PUBLIC_OFF, PAID_LEAVE]
    assert ev.employee_counts == {"A": 1}
    assert ev.employee_actual_work_days == {"A": 2}
    assert ev.employee_public_counts == ev.employee_paid_counts == ev.employee_recovery_counts == {"A": 1}
    assert ev.daily_counts[data.days[1]] == 0
    assert is_working(RECOVERY, data.settings) is False
    assert is_actual_working(RECOVERY, data.settings) is True


def test_only_public_holidays_credit_leave_requests_without_double_counting():
    data = _night_example()
    days = data.days
    data = replace(data, requests=(
        LeaveCandidate("PUBLIC", "A", days[2], 1),
        LeaveCandidate("SAME_PUBLIC", "A", days[2], 2),
        LeaveCandidate("RECOVERY", "A", days[1], 1),
        LeaveCandidate("PAID", "A", days[3], 1),
    ))
    result = solve_schedule(data)
    ev = _assert_valid(data, result)
    assert result.status is SolverStatus.OPTIMAL
    assert ev.adopted_count == 1
    results = {r.request_id: r for r in ev.request_results}
    assert results["PUBLIC"].adopted_day == days[2]
    assert all(results[name].adopted_day is None for name in ("SAME_PUBLIC", "RECOVERY", "PAID"))


def test_derived_cross_calendar_equation_detects_mismatch_instead_of_changing_public_days():
    data = _night_example()
    data = replace(data, employees=(replace(data.employees[0], public_holidays=2),), fixed_off=())
    issues = validate_input(data)
    assert any("勤務開始1日＋明け1日＋公休2日＋有給1日" in issue.message for issue in issues)
    assert solve_schedule(data).schedule is None


@pytest.mark.parametrize("replacement", [PUBLIC_OFF, PAID_LEAVE, "日勤"])
def test_required_recovery_cannot_be_replaced_or_overlap_another_state(replacement):
    data = _night_example()
    result = solve_schedule(data)
    _assert_valid(data, result)
    broken = dict(result.schedule)
    broken["A", data.days[1]] = replacement
    ev = evaluate_schedule(data, broken)
    assert not ev.valid
    assert any(i.kind == "明け違反" and i.day == data.days[1] and i.source.column == "D4" for i in ev.issues)


def test_recovery_cannot_appear_without_previous_cross_shift():
    data = _night_example()
    result = solve_schedule(data)
    broken = dict(result.schedule)
    broken["A", data.days[2]] = RECOVERY
    ev = evaluate_schedule(data, broken)
    assert not ev.valid
    assert any(i.kind == "明け違反" and "前日に日跨ぎ勤務がない" in i.message for i in ev.issues)


def test_unapproved_paid_leave_cannot_be_created_and_approved_paid_cannot_be_dropped():
    data = _night_example()
    result = solve_schedule(data)
    broken = dict(result.schedule)
    broken["A", data.days[2]] = PAID_LEAVE
    broken["A", data.days[3]] = PUBLIC_OFF
    ev = evaluate_schedule(data, broken)
    assert not ev.valid
    assert any(i.kind == "未承認有給" and i.day == data.days[2] for i in ev.issues)
    assert any(i.kind == "有給違反" and i.day == data.days[3] for i in ev.issues)


@pytest.mark.parametrize("limit,expected", [(1, SolverStatus.INFEASIBLE), (2, SolverStatus.OPTIMAL)])
def test_recovery_is_actual_calendar_work_for_consecutive_limit(limit, expected):
    data = _night_example()
    data = replace(data, settings=replace(data.settings, max_consecutive_days=limit))
    result = solve_schedule(data)
    assert result.status is expected
    if result.schedule is not None:
        assert _assert_valid(data, result).employee_actual_work_days == {"A": 2}
    else:
        manual = {("A", day): label for day, label in zip(data.days, ("夜勤", RECOVERY, PUBLIC_OFF, PAID_LEAVE))}
        ev = evaluate_schedule(data, manual)
        assert any(i.kind == "連続勤務違反" and i.day == data.days[1] and "連続勤務2日" in i.message for i in ev.issues)


@pytest.mark.parametrize("collision", [PUBLIC_OFF, PAID_LEAVE])
def test_first_day_recovery_from_previous_cross_rejects_public_or_paid_collision(collision):
    days = _days(2)
    data = InputData(
        (Employee("A", "職員A", None, previous_work_days=1, previous_shift="夜勤", public_holidays=1 if collision == PUBLIC_OFF else 0),),
        tuple(DayRequirement(day, 0) for day in days),
        fixed_off=(FixedOff("A", days[0]),) if collision == PUBLIC_OFF else (),
        paid_leaves=(PaidLeave("A", days[0]),) if collision == PAID_LEAVE else (),
        settings=ShiftSettings(mode=DAY_ONLY, leave_enabled=True, day_count_mode=PUBLIC_DAYS, shift_times=(DAY_TIME, NIGHT_TIME), shift_times_confirmed=True),
    )
    assert any("初日は明け" in issue.message for issue in validate_input(data))
    assert solve_schedule(data).schedule is None


def test_previous_cross_still_creates_recovery_after_switch_to_day_only():
    days = _days(2)
    data = InputData(
        (Employee("A", "職員A", None, previous_work_days=1, previous_shift="夜勤", public_holidays=1),),
        tuple(DayRequirement(day, 0) for day in days),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, shift_times=(DAY_TIME, NIGHT_TIME), shift_times_confirmed=True),
    )
    effective = configure_input(data)
    assert effective.employees[0].work_days is None
    result = solve_schedule(effective)
    ev = _assert_valid(effective, result)
    assert result.schedule == {("A", days[0]): RECOVERY, ("A", days[1]): PUBLIC_OFF}
    assert ev.employee_counts == {"A": 0}
    assert ev.employee_actual_work_days == {"A": 1}


def test_previous_recovery_is_a_valid_actual_working_day_in_history():
    data = InputData(
        (Employee("A", "職員A", None, previous_work_days=2, previous_shift=RECOVERY, public_holidays=0),),
        (DayRequirement(START, 1),),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, max_consecutive_days=3),
    )
    assert _assert_valid(data, solve_schedule(data)).employee_actual_work_days == {"A": 1}
    too_short = replace(data, settings=replace(data.settings, max_consecutive_days=2))
    assert solve_schedule(too_short).status is SolverStatus.INFEASIBLE
    inconsistent = replace(data, employees=(replace(data.employees[0], previous_work_days=0),))
    assert any("連勤0日" in i.message for i in validate_input(inconsistent))


def test_shift_ending_at_midnight_does_not_create_recovery():
    days = _days(2)
    data = InputData(
        (Employee("A", "職員A", None, public_holidays=1),),
        (DayRequirement(days[0], 1, shift_counts=(("日勤", 0), ("夜勤", 1))), DayRequirement(days[1], 0, shift_counts=(("日勤", 0), ("夜勤", 0)))),
        settings=ShiftSettings(mode=TWO_SHIFT, leave_enabled=True, day_count_mode=PUBLIC_DAYS,
            shift_times=(DAY_TIME, ShiftTime("夜勤", 16 * 60, 1440)), shift_times_confirmed=True),
    )
    assert crossing_shifts(data.settings) == frozenset()
    effective = configure_input(data)
    assert effective.employees[0].work_days == 1
    result = solve_schedule(effective)
    ev = _assert_valid(effective, result)
    assert result.schedule == {("A", days[0]): "夜勤", ("A", days[1]): PUBLIC_OFF}
    assert ev.employee_recovery_counts == {"A": 0}


def test_day_label_can_itself_cross_midnight_when_confirmed_times_say_so():
    days = _days(2)
    data = InputData(
        (Employee("A", "職員A", None, public_holidays=0),),
        (DayRequirement(days[0], 1), DayRequirement(days[1], 0)),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, night_rest=False,
            shift_times=(ShiftTime("日勤", 22 * 60, 30 * 60),), shift_times_confirmed=True),
    )
    result = solve_schedule(data)
    assert result.schedule == {("A", days[0]): "日勤", ("A", days[1]): RECOVERY}
    assert _assert_valid(data, result).employee_actual_work_days == {"A": 2}


def test_final_cross_needs_no_in_period_recovery_but_is_not_a_future_day_guarantee():
    data = InputData(
        (Employee("A", "職員A", None, public_holidays=0),),
        (DayRequirement(START, 1, shift_counts=(("日勤", 0), ("夜勤", 1))),),
        settings=ShiftSettings(mode=TWO_SHIFT, leave_enabled=True, day_count_mode=PUBLIC_DAYS,
            shift_times=(DAY_TIME, NIGHT_TIME), shift_times_confirmed=True),
    )
    result = solve_schedule(data)
    ev = _assert_valid(data, result)
    assert result.schedule == {("A", START): "夜勤"}
    assert ev.employee_recovery_counts == {"A": 0}


@pytest.mark.parametrize("settings", [
    ShiftSettings(mode=TWO_SHIFT, leave_enabled=True),
    ShiftSettings(mode=TWO_SHIFT, leave_enabled=True, shift_times=(DAY_TIME, NIGHT_TIME)),
    ShiftSettings(mode=TWO_SHIFT, leave_enabled=True, shift_times=(DAY_TIME,), shift_times_confirmed=True),
    ShiftSettings(leave_enabled=True, shift_times=(DAY_TIME,), shift_times_confirmed=False),
    ShiftSettings(leave_enabled=True, shift_times=(ShiftTime("日勤", 1440, 1500),), shift_times_confirmed=True),
    ShiftSettings(leave_enabled=True, shift_times=(ShiftTime("日勤", 600, 600),), shift_times_confirmed=True),
    ShiftSettings(leave_enabled=True, shift_times=(ShiftTime("日勤", 600, 2041),), shift_times_confirmed=True),
    ShiftSettings(leave_enabled=True, shift_times=(ShiftTime("日勤", True, 600),), shift_times_confirmed=True),
    ShiftSettings(leave_enabled=True, shift_times=(DAY_TIME, DAY_TIME), shift_times_confirmed=True),
])
def test_unconfirmed_missing_or_invalid_shift_times_block_calculation(settings):
    data = replace(sample_data(), settings=settings)
    with pytest.raises(InputError):
        configure_input(data)
    assert solve_schedule(data).schedule is None


@pytest.mark.parametrize("paid", [
    (PaidLeave("UNKNOWN", START),),
    (PaidLeave("A", START - timedelta(days=1)),),
    (PaidLeave("A", START), PaidLeave("A", START)),
])
def test_invalid_paid_rows_keep_source_and_block_output(paid):
    paid = tuple(replace(row, source=Source("leave.xlsx", "有給", i + 2)) for i, row in enumerate(paid))
    data = InputData((Employee("A", "職員A", 0),), (DayRequirement(START, 0),), paid_leaves=paid, settings=ShiftSettings(leave_enabled=True))
    issues = validate_input(data)
    assert any(issue.source.file == "leave.xlsx" and issue.source.sheet == "有給" and issue.source.row >= 2 for issue in issues)
    assert solve_schedule(data).schedule is None


def test_paid_and_fixed_public_on_same_date_are_rejected():
    data = InputData(
        (Employee("A", "職員A", 0),), (DayRequirement(START, 0),),
        fixed_off=(FixedOff("A", START),), paid_leaves=(PaidLeave("A", START),),
        settings=ShiftSettings(leave_enabled=True),
    )
    assert any("確定公休と有給が重複" in issue.message for issue in validate_input(data))


def test_advanced_data_is_not_silently_ignored_when_feature_is_disabled():
    for data in (
        replace(sample_data(), paid_leaves=(PaidLeave("A", sample_data().days[0]),)),
        replace(sample_data(), employees=(replace(sample_data().employees[0], public_holidays=0), *sample_data().employees[1:])),
        replace(sample_data(), settings=ShiftSettings(day_count_mode=PUBLIC_DAYS, default_public_holidays=1)),
    ):
        assert validate_input(data)
        assert solve_schedule(data).schedule is None


def test_advanced_assignment_helpers_reject_ambiguous_legacy_values():
    settings = ShiftSettings(leave_enabled=True)
    assert assignment_labels(settings) == ("日勤", PUBLIC_OFF, PAID_LEAVE, RECOVERY)
    assert assignment_label(False, settings) is None
    assert assignment_label(OFF, settings) is None
    assert assignment_label(PUBLIC_OFF, settings) == PUBLIC_OFF
    assert assignment_label(PAID_LEAVE, settings) == PAID_LEAVE
    assert not is_working(PUBLIC_OFF, settings)
    assert not is_actual_working(PAID_LEAVE, settings)
    assert assignment_label(False) == OFF  # 旧設定は完全互換。


@pytest.mark.parametrize("previous", [None, OFF, PUBLIC_OFF])
def test_legacy_previous_off_matches_public_transition_assumption(previous):
    data = InputData(
        (Employee("A", "職員A", None, previous_work_days=0, previous_shift=previous, public_holidays=0),),
        (DayRequirement(START, 1),),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS,
                               forbidden_transitions=((PUBLIC_OFF, "日勤"),)),
    )
    assert validate_input(data) == []
    assert solve_schedule(data).status is SolverStatus.INFEASIBLE
    evaluation = evaluate_schedule(data, {("A", START): "日勤"})
    assert any(issue.kind == "勤務順序違反" and issue.day == START and issue.source.column == "C4"
               for issue in evaluation.issues)
    # 当日の旧「休み」を高度勤務表へ受け入れる変更ではない。
    assert assignment_label(OFF, data.settings) is None
