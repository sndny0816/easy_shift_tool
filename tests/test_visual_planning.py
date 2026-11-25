"""週・日付の率指定と、元有給申請の明示配分をGUIなしで検証する。"""
from dataclasses import replace
from datetime import date, timedelta
from decimal import Decimal

import pytest

from sinta_shift.models import (
    DAY_ONLY, TWO_SHIFT, PUBLIC_DAYS, DIRECT_DAYS, PUBLIC_OFF, PAID_LEAVE,
    PUBLIC_REQUEST, PAID_REQUEST, Source, ShiftSettings, Employee, DayRequirement,
    IndependentLeaveRequest, LeaveCandidate, PaidLeave, InputData, InputError, SolverStatus,
)
from sinta_shift.planning import (
    configure_input, week_start, attendance_ratio_for_day, attendance_source_for_day,
    public_holiday_target, confirm_paid_requests, set_leave_choices, plan_paid_mix,
)
from sinta_shift.validation import validate_input
from sinta_shift.solver import solve_schedule
from sinta_shift.evaluation import evaluate_schedule


MON = date(2026, 11, 2)


def _days(n=7):
    return tuple(MON + timedelta(days=i) for i in range(n))


def _ratio_data():
    return InputData(
        tuple(Employee(s, s, 0) for s in "ABCD"),
        tuple(DayRequirement(day, 3, Source("input.xlsx", "必要人数", i + 2), (("日勤", 1), ("夜勤", 2))) for i, day in enumerate(_days(9))),
        settings=ShiftSettings(mode=TWO_SHIFT),
    )


def _paid_data(*, days=4, public=1, base_paid=(), direct=False):
    requests = (
        IndependentLeaveRequest("P", "A", MON, PUBLIC_REQUEST, Source("input.xlsx", "希望休", 2, "B2")),
        IndependentLeaveRequest("A1", "A", MON + timedelta(days=1), PAID_REQUEST, Source("input.xlsx", "希望休", 2, "D2")),
        IndependentLeaveRequest("A2", "A", MON + timedelta(days=2), PAID_REQUEST, Source("input.xlsx", "希望休", 2, "F2")),
    )
    return InputData(
        (Employee("A", "職員A", 1 if direct else None),),
        tuple(DayRequirement(day, int(i == 3)) for i, day in enumerate(_days(days))),
        requests=(LeaveCandidate("P", "A", MON, 1, requests[0].source),),
        independent_requests=requests, paid_leaves=base_paid,
        settings=ShiftSettings(leave_enabled=True, day_count_mode=DIRECT_DAYS if direct else PUBLIC_DAYS,
                               default_public_holidays=public), request_format="independent",
    )


def test_week_start_and_override_priority_keep_zero_as_an_explicit_value():
    settings = ShiftSettings(weekday_ratios=(Decimal(50),) * 7,
        weekly_ratios=((MON, Decimal(75)),), date_ratios=((MON + timedelta(days=2), Decimal(0)),))
    assert week_start(MON + timedelta(days=6)) == MON
    assert attendance_ratio_for_day(MON + timedelta(days=2), settings) == Decimal(0)
    assert attendance_source_for_day(MON + timedelta(days=2), settings) == "日付指定"
    assert attendance_ratio_for_day(MON, settings) == Decimal(75)
    assert attendance_source_for_day(MON, settings) == "週指定"
    assert attendance_ratio_for_day(MON + timedelta(days=7), settings) == Decimal(50)
    assert attendance_source_for_day(MON + timedelta(days=7), settings) == "曜日指定"
    assert attendance_ratio_for_day(MON, ShiftSettings()) is None
    assert attendance_source_for_day(MON, ShiftSettings()) == "Excel人数"


def test_only_matching_override_dates_replace_excel_total_and_explicit_shift_counts():
    raw = _ratio_data()
    settings = replace(raw.settings, weekly_ratios=((MON, Decimal(50)),),
                       date_ratios=((MON + timedelta(days=2), Decimal(25)),))
    effective = configure_input(raw, settings)
    assert [r.required for r in effective.requirements] == [2, 2, 1, 2, 2, 2, 2, 3, 3]
    assert effective.requirements[0].shift_counts == (("日勤", 1), ("夜勤", 1))
    assert effective.requirements[2].shift_counts == (("日勤", 1), ("夜勤", 0))
    assert effective.requirements[7].shift_counts == raw.requirements[7].shift_counts
    assert effective.requirements[8].shift_counts == raw.requirements[8].shift_counts
    assert [r.source for r in effective.requirements] == [r.source for r in raw.requirements]
    assert all(r.required == 3 for r in raw.requirements)
    assert configure_input(effective) == effective


def test_override_then_weekday_then_excel_fallback_and_exact_ceiling():
    days = _days(9)
    raw = InputData(tuple(Employee(str(i), f"職員{i}", 0) for i in range(30)),
                    tuple(DayRequirement(day, 5) for day in days))
    settings = ShiftSettings(weekly_ratios=((MON, Decimal(66)),), date_ratios=((days[1], Decimal("0.1")),))
    effective = configure_input(raw, settings)
    assert effective.requirements[0].required == 20
    assert effective.requirements[1].required == 1
    assert effective.requirements[7].required == 5
    assert effective.employees == raw.employees  # 配置数変更を勤務日数の改変で隠さない。
    assert any("合計" in issue.message for issue in validate_input(effective))


def test_week_start_outside_month_is_allowed_when_that_week_intersects_period():
    sunday = date(2026, 11, 1)
    raw = InputData((Employee("A", "職員A", 1),), (DayRequirement(sunday, 0),))
    settings = ShiftSettings(weekly_ratios=((date(2026, 10, 26), Decimal(100)),))
    effective = configure_input(raw, settings)
    assert effective.requirements[0].required == 1
    assert solve_schedule(effective).status is SolverStatus.OPTIMAL


@pytest.mark.parametrize("field,values", [
    ("weekly_ratios", ((MON + timedelta(days=1), Decimal(50)),)),
    ("weekly_ratios", ((MON, Decimal(50)), (MON, Decimal(70)))),
    ("date_ratios", ((MON, Decimal(50)), (MON, Decimal(70)))),
    ("date_ratios", (("2026-11-02", Decimal(50)),)),
    ("date_ratios", ((MON, Decimal("NaN")),)),
    ("date_ratios", ((MON, Decimal("Infinity")),)),
    ("date_ratios", ((MON, Decimal(-1)),)),
    ("date_ratios", ((MON, Decimal(101)),)),
    ("date_ratios", ((MON, 50.0),)),
    ("weekly_ratios", ((MON - timedelta(days=7), Decimal(50)),)),
    ("date_ratios", ((MON - timedelta(days=1), Decimal(50)),)),
    ("weekly_ratios", ((MON,),)),
])
def test_bad_override_types_ranges_duplicates_and_periods_are_rejected(field, values):
    data = _ratio_data()
    settings = replace(data.settings, **{field: values})
    with pytest.raises(InputError):
        configure_input(data, settings)
    assert validate_input(replace(data, settings=settings))


def test_visual_ratio_conditions_produce_real_shift_plan():
    days = _days(7)
    targets = [3, 3, 1, 3, 3, 3, 3]  # 週75%、水曜だけ25%、4人。
    data = InputData(
        tuple(Employee(s, s, n) for s, n in zip("ABCD", (5, 5, 5, 4))),
        tuple(DayRequirement(day, 0) for day in days),
        settings=ShiftSettings(max_consecutive_days=4, weekday_ratios=(Decimal(50),) * 7,
                               weekly_ratios=((MON, Decimal(75)),), date_ratios=((days[2], Decimal(25)),)),
    )
    effective = configure_input(data)
    assert [r.required for r in effective.requirements] == targets
    result = solve_schedule(effective)
    assert result.status is SolverStatus.OPTIMAL
    assert evaluate_schedule(effective, result.schedule).valid


def test_paid_to_public_keeps_total_nonworking_days_and_makes_normal_public_candidate():
    data = _paid_data()
    selected = set_leave_choices(data, {"A1"}, {"A2"})
    employee = selected.employees[0]
    assert employee.public_holidays is None and employee.public_holiday_adjustment == 1
    assert public_holiday_target(employee, selected.settings) == 2
    assert selected.converted_public_request_ids == frozenset({"A2"})
    assert [(r.request_id, r.priority) for r in selected.requests] == [("P", 1), ("A2", 1)]
    converted = next(r for r in selected.independent_requests if r.request_id == "A2")
    assert converted.leave_type == PUBLIC_REQUEST and converted.original_leave_type == PAID_REQUEST
    assert len(selected.paid_leaves) == 1 and selected.paid_leaves[0].request_id == "A1"
    assert set_leave_choices(selected, {"A1"}, {"A2"}) == selected
    effective = configure_input(selected)
    assert effective.employees[0].work_days == 1
    assert configure_input(effective) == effective
    result = solve_schedule(effective)
    assert result.status is SolverStatus.OPTIMAL
    ev = evaluate_schedule(effective, result.schedule)
    assert ev.valid and ev.adopted_count == 2
    assert ev.employee_public_counts == {"A": 2}
    assert ev.employee_paid_counts == {"A": 1}
    assert [r.request_id for r in ev.paid_request_results] == ["A1"]
    assert {r.request_id for r in ev.request_results} == {"P", "A2"}


def test_reconfirming_converted_request_removes_public_adjustment_and_candidate_once():
    data = set_leave_choices(_paid_data(), {"A1"}, {"A2"})
    restored = confirm_paid_requests(data, {"A1", "A2"})
    assert restored.converted_public_request_ids == frozenset()
    assert restored.employees[0].public_holiday_adjustment == 0
    assert [r.request_id for r in restored.requests] == ["P"]
    assert len(restored.paid_leaves) == 2
    assert confirm_paid_requests(restored, {"A1", "A2"}) == restored
    assert evaluate_schedule(restored, solve_schedule(restored).schedule).valid


def test_removing_both_decisions_returns_original_paid_request_to_pending():
    data = set_leave_choices(_paid_data(), {"A1"}, {"A2"})
    pending = set_leave_choices(data, set(), set())
    assert pending.paid_leaves == ()
    assert pending.converted_public_request_ids == pending.confirmed_paid_request_ids == frozenset()
    assert pending.employees[0].public_holiday_adjustment == 0
    assert [r.request_id for r in pending.requests] == ["P"]
    assert all(r.leave_type == PAID_REQUEST for r in pending.independent_requests if r.request_id.startswith("A"))
    assert sum(issue.kind == "有給申請未確認" for issue in validate_input(pending)) == 2
    assert solve_schedule(pending).schedule is None


def test_direct_mode_keeps_work_start_count_without_public_target():
    selected = set_leave_choices(_paid_data(direct=True), set(), {"A1", "A2"})
    assert selected.employees[0].work_days == 1
    assert public_holiday_target(selected.employees[0], selected.settings) is None
    result = solve_schedule(selected)
    ev = evaluate_schedule(selected, result.schedule)
    assert ev.valid and ev.employee_counts == {"A": 1}
    assert ev.employee_public_counts == {"A": 3}


def test_base_paid_sheet_is_preserved_by_allocation_and_reversal():
    base = PaidLeave("A", MON + timedelta(days=4), Source("input.xlsx", "有給", 2))
    raw = _paid_data(days=5, base_paid=(base,))
    selected = set_leave_choices(raw, {"A1"}, {"A2"})
    assert selected.paid_leaves[0] == base
    assert set_leave_choices(selected, set(), set()).paid_leaves == (base,)
    ev = evaluate_schedule(selected, solve_schedule(selected).schedule)
    assert ev.valid and ev.employee_paid_counts == {"A": 2}


@pytest.mark.parametrize("paid,public", [({"P"}, set()), (set(), {"P"}), ({"UNKNOWN"}, set()),
                                         (set(), {"UNKNOWN"}), ({"A1"}, {"A1"})])
def test_choices_cannot_convert_original_public_unknown_ids_or_overlap(paid, public):
    with pytest.raises(InputError):
        set_leave_choices(_paid_data(), paid, public)


def test_forged_original_public_to_paid_metadata_is_rejected():
    raw = _paid_data()
    request = replace(raw.independent_requests[0], leave_type=PAID_REQUEST, original_leave_type=PUBLIC_REQUEST)
    forged = replace(raw, independent_requests=(request, *raw.independent_requests[1:]))
    assert any("元の公休希望を有給申請" in issue.message for issue in validate_input(forged))


def _mix_data():
    employees = (Employee("Z", "あ", 0), Employee("A", "い", 0), Employee("B", "い", 0))
    metadata = tuple(IndependentLeaveRequest(f"{e.id}{i}", e.id, MON + timedelta(days=i), PAID_REQUEST)
                     for e in employees for i in range(3))
    return InputData(employees, tuple(DayRequirement(day, 0) for day in _days(4)),
                     independent_requests=metadata, settings=ShiftSettings(leave_enabled=True), request_format="independent")


def test_mix_plan_is_stable_round_robin_and_does_not_apply_itself():
    raw = _mix_data()
    paid, public = plan_paid_mix(raw, 4)
    assert paid == frozenset({"Z0", "A0", "B0", "Z1"})
    assert public == frozenset({"A1", "A2", "B1", "B2", "Z2"})
    assert plan_paid_mix(raw, 4) == (paid, public)
    assert raw.confirmed_paid_request_ids == raw.converted_public_request_ids == frozenset()
    assert raw.paid_leaves == ()
    assert all(r.leave_type == PAID_REQUEST and r.original_leave_type == "" for r in raw.independent_requests)


def test_mix_plan_allocates_only_pending_and_preserves_existing_decisions():
    initial = set_leave_choices(_mix_data(), {"B2"}, {"Z2"})
    paid, public = plan_paid_mix(initial, 2)
    assert "B2" in paid and "Z2" in public
    assert len(paid) == 3 and len(public) == 6
    assert paid | public == {r.request_id for r in initial.independent_requests}
    assert plan_paid_mix(set_leave_choices(initial, paid, public), 0) == (paid, public)


@pytest.mark.parametrize("count", [-1, 10, True, 1.5])
def test_invalid_proposal_count_is_rejected(count):
    with pytest.raises(InputError):
        plan_paid_mix(_mix_data(), count)


def test_fixed_paid_overlap_is_rejected_without_rewriting_the_base_record():
    raw = _paid_data(base_paid=(PaidLeave("A", MON + timedelta(days=1)),))
    with pytest.raises(InputError):
        plan_paid_mix(raw, 1)
    with pytest.raises(InputError):
        set_leave_choices(raw, {"A1"}, {"A2"})
    legacy_confirmed = confirm_paid_requests(raw, {"A1"})
    assert len(legacy_confirmed.paid_leaves) == 2
    assert any("有給申請と元の有給一覧が重複" in issue.message for issue in validate_input(legacy_confirmed))


def test_converted_public_target_is_rechecked_after_manual_edit():
    data = set_leave_choices(_paid_data(), {"A1"}, {"A2"})
    result = solve_schedule(data)
    broken = dict(result.schedule)
    broken["A", MON + timedelta(days=2)] = PAID_LEAVE
    ev = evaluate_schedule(data, broken)
    assert not ev.valid
    assert any(issue.kind == "未承認有給" for issue in ev.issues)
    assert any(issue.kind == "公休日数違反" for issue in ev.issues)


@pytest.mark.parametrize("public_change", [False, True])
def test_manual_choices_reject_duplicate_employee_day_before_applying(public_change):
    raw = _paid_data()
    second = replace(raw.independent_requests[2], day=raw.independent_requests[1].day)
    raw = replace(raw, independent_requests=(*raw.independent_requests[:2], second))
    with pytest.raises(InputError, match="同じ職員・同じ日"):
        set_leave_choices(raw, {"A1"}, {"A2"} if public_change else set())


def test_manual_public_conversion_cannot_hide_base_paid_or_fixed_public_collision():
    raw = _paid_data(base_paid=(PaidLeave("A", MON + timedelta(days=2)),))
    with pytest.raises(InputError, match="固定有給"):
        set_leave_choices(raw, {"A1"}, {"A2"})
    from sinta_shift.models import FixedOff
    raw = replace(_paid_data(), fixed_off=(FixedOff("A", MON + timedelta(days=2)),))
    with pytest.raises(InputError, match="確定公休"):
        set_leave_choices(raw, {"A1"}, {"A2"})
