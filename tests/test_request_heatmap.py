"""希望人数・申請件数・代替候補・確定休暇を混同しない日別集計。"""
from dataclasses import FrozenInstanceError, replace
from datetime import date, timedelta
from decimal import Decimal

import pytest

from sinta_shift.models import (
    PUBLIC_OFF, PAID_LEAVE, RECOVERY, PUBLIC_DAYS, PUBLIC_REQUEST, PAID_REQUEST,
    InputData, Employee, DayRequirement, FixedOff, PaidLeave, LeaveCandidate,
    IndependentLeaveRequest, ShiftSettings, RequestResult, Source,
)
from sinta_shift.sample import sample_data
from sinta_shift.planning import configure_input, set_leave_choices
from sinta_shift.solver import solve_schedule
from sinta_shift.evaluation import evaluate_schedule
from sinta_shift.request_heatmap import (
    build_request_heatmap, ADOPTED, UNADOPTED, UNCALCULATED, OTHER_CANDIDATE,
    UNKNOWN_RESULT, PENDING,
)


D7, D8, D9 = (date(2026, 11, day) for day in (7, 8, 9))


def _index(rows):
    return {row.day: row for row in rows}


def _wide_data():
    days = (D7, D8, D9)
    metadata = (
        IndependentLeaveRequest("P1", "A", D7, PUBLIC_REQUEST, Source("wide.xlsx", "希望休", 2, "B2")),
        IndependentLeaveRequest("P2", "A", D8, PUBLIC_REQUEST, Source("wide.xlsx", "希望休", 2, "D2")),
        IndependentLeaveRequest("A1", "A", D9, PAID_REQUEST, Source("wide.xlsx", "希望休", 2, "F2")),
    )
    return InputData(
        (Employee("A", "職員A", None),), tuple(DayRequirement(day, 0) for day in days),
        requests=tuple(LeaveCandidate(r.request_id, r.employee_id, r.day, 1, r.source) for r in metadata if r.leave_type == PUBLIC_REQUEST),
        independent_requests=metadata,
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, default_public_holidays=2),
        request_format="independent",
    )


def test_original_acceptance_example_keeps_concentration_and_distinguishes_other_candidate():
    data = sample_data()
    before = _index(build_request_heatmap(data))
    assert [before[day].request_employee_count for day in (D7, D8, D9)] == [2, 1, 1]
    assert all(row.adopted_employee_count is None and row.unadopted_request_count is None for row in before.values())
    assert all(entry.status == UNCALCULATED for row in before.values() for entry in row.entries)
    result = solve_schedule(data)
    evaluation = evaluate_schedule(data, result.schedule)
    after = _index(build_request_heatmap(data, result.schedule, evaluation))
    assert [after[day].request_employee_count for day in (D7, D8, D9)] == [2, 1, 1]
    assert [after[day].adopted_employee_count for day in (D7, D8, D9)] == [1, 1, 1]
    assert after[D7].other_candidate_adopted_count == 1
    assert after[D7].unadopted_request_count == 0
    assert next(entry for entry in after[D7].entries if entry.request_id == "R2").status == OTHER_CANDIDATE
    adopted = next(entry for entry in after[D8].entries if entry.request_id == "R2")
    assert adopted.status == ADOPTED and adopted.adopted_day == D8 and adopted.priority == 2
    assert sum(row.candidate_count for row in after.values()) == 4
    assert len({entry.request_id for row in after.values() for entry in row.entries}) == 3
    assert all(not row.is_reference and row.current_staff_count == 2 for row in after.values())


def test_request_and_metadata_duplicate_is_one_detail_and_one_employee():
    data = _wide_data()
    rows = _index(build_request_heatmap(data))
    assert [row.request_employee_count for row in rows.values()] == [1, 1, 1]
    assert all(row.candidate_count == 1 for row in rows.values())
    assert rows[D7].public_request_employee_count == 1 and rows[D7].paid_request_employee_count == 0
    assert rows[D9].paid_request_employee_count == 1 and rows[D9].entries[0].approval_status == PENDING
    assert rows[D7].entries[0].source.column == "B2"
    assert rows[D7].entries[0].is_alternative is False


def test_same_employee_same_day_different_legacy_requests_keep_details_but_count_one():
    raw = sample_data()
    data = replace(raw, requests=(LeaveCandidate("R1", "A", D7, 1), LeaveCandidate("R2", "A", D7, 2)))
    result = solve_schedule(data)
    row = _index(build_request_heatmap(data, result.schedule))[D7]
    assert row.request_employee_count == 1
    assert row.public_request_employee_count == 1
    assert row.candidate_count == row.request_count == 2
    assert row.adopted_employee_count == 1 and row.unadopted_request_count == 1
    assert {entry.status for entry in row.entries} == {ADOPTED, UNADOPTED}


def test_another_candidate_adoption_is_not_overwritten_by_physical_rest_on_this_day():
    days = (D7, D8, D9)
    data = InputData(
        (Employee("A", "職員A", 1), Employee("B", "職員B", 2)),
        tuple(DayRequirement(day, 1) for day in days),
        requests=(LeaveCandidate("R", "A", D7, 2), LeaveCandidate("R", "A", D8, 1)),
    )
    schedule = {("A", day): day == D9 for day in days}
    schedule.update({("B", day): day != D9 for day in days})
    rows = _index(build_request_heatmap(data, schedule))
    assert rows[D7].entries[0].status == OTHER_CANDIDATE
    assert rows[D7].entries[0].adopted_day == D8
    assert rows[D7].unadopted_request_count == 0


def test_fixed_days_are_separate_from_requests_and_derived_paid_count_is_not_doubled():
    data = configure_input(set_leave_choices(_wide_data(), {"A1"}, set()))
    data = replace(data, fixed_off=(FixedOff("A", D7),))
    result = solve_schedule(data)
    rows = _index(build_request_heatmap(data, result.schedule))
    assert rows[D7].request_employee_count == 1 and rows[D7].fixed_off_employee_count == 1
    assert rows[D9].request_employee_count == 1 and rows[D9].fixed_paid_employee_count == 1
    assert rows[D9].adopted_employee_count == 1
    assert rows[D9].entries[0].status == ADOPTED
    assert rows[D9].entries[0].approval_status == "有給確定"
    fixed_only = InputData((Employee("A", "職員A", 0),), (DayRequirement(D7, 0),),
        fixed_off=(FixedOff("A", D7), FixedOff("A", D7)),
        paid_leaves=(PaidLeave("A", D7), PaidLeave("A", D7)), settings=ShiftSettings(leave_enabled=True))
    row = build_request_heatmap(fixed_only)[0]
    assert row.request_employee_count == row.candidate_count == 0
    assert row.fixed_off_employee_count == row.fixed_paid_employee_count == 1


def test_paid_approval_alone_never_means_adopted_without_a_schedule():
    data = set_leave_choices(_wide_data(), {"A1"}, set())
    row = _index(build_request_heatmap(data))[D9]
    assert row.entries[0].approval_status == "有給確定"
    assert row.entries[0].status == UNCALCULATED
    assert row.adopted_employee_count is None
    assert row.current_staff_count is None


def test_manual_paid_change_uses_actual_assignment_and_marks_invalid_outcomes_reference():
    data = configure_input(set_leave_choices(_wide_data(), {"A1"}, set()))
    result = solve_schedule(data)
    broken = dict(result.schedule)
    broken["A", D9] = PUBLIC_OFF
    evaluation = evaluate_schedule(data, broken)
    # paid_request_resultsは確認状態しか示さないため、実配置との照合が必要。
    assert evaluation.paid_request_results[0].adopted_day == D9
    row = _index(build_request_heatmap(data, broken, evaluation))[D9]
    assert row.is_reference and row.warnings
    assert row.entries[0].status == UNADOPTED
    assert row.entries[0].approval_status == "有給確定"
    assert row.adopted_employee_count == 0
    assert row.unadopted_request_count == 1


def test_manual_public_change_recalculates_status_and_preserves_concentration():
    data = sample_data()
    result = solve_schedule(data)
    before = _index(build_request_heatmap(data, result.schedule))
    broken = dict(result.schedule)
    broken["A", D7] = True
    after = _index(build_request_heatmap(data, broken))
    assert after[D7].request_employee_count == before[D7].request_employee_count == 2
    assert next(entry for entry in after[D7].entries if entry.employee_id == "A").status == UNADOPTED
    assert after[D7].is_reference
    assert after[D7].current_staff_count == 3


def test_converted_paid_is_public_in_counts_and_keeps_original_kind():
    data = configure_input(set_leave_choices(_wide_data(), set(), {"A1"}))
    result = solve_schedule(data)
    row = _index(build_request_heatmap(data, result.schedule))[D9]
    assert row.public_request_employee_count == 1 and row.paid_request_employee_count == 0
    assert row.fixed_paid_employee_count == 0
    assert row.entries[0].kind == PUBLIC_REQUEST
    assert row.entries[0].original_kind == PAID_REQUEST
    assert row.entries[0].approval_status == "公休へ変更"
    assert row.entries[0].status == ADOPTED


def test_pending_paid_with_a_physical_paid_cell_is_not_automatically_adopted():
    data = _wide_data()
    schedule = {("A", day): PUBLIC_OFF for day in data.days}
    schedule["A", D9] = PAID_LEAVE
    row = _index(build_request_heatmap(data, schedule))[D9]
    assert row.entries[0].status == PENDING
    assert row.adopted_employee_count is None
    assert row.is_reference


def test_missing_unknown_duplicate_or_stale_results_are_not_silently_unadopted():
    data = sample_data()
    result = solve_schedule(data)
    evaluation = evaluate_schedule(data, result.schedule)
    for bad in (
        replace(evaluation, request_results=[]),
        replace(evaluation, request_results=[RequestResult("UNKNOWN", "A", D7, 1, "第1希望を採用")]),
        replace(evaluation, request_results=[*evaluation.request_results, evaluation.request_results[0]]),
        replace(evaluation, request_results=[replace(r, status="不明") if r.request_id == "R1" else r for r in evaluation.request_results]),
        replace(evaluation, request_results=[replace(r, adopted_day=D8) if r.request_id == "R1" else r for r in evaluation.request_results]),
    ):
        row = _index(build_request_heatmap(data, result.schedule, bad))[D7]
        assert row.is_reference and row.warnings
        assert row.entries[0].status == UNKNOWN_RESULT
        assert row.adopted_employee_count is None and row.unadopted_request_count is None


def test_zero_request_days_and_period_crossing_month_are_retained():
    start = date(2026, 11, 30)
    days = tuple(start + timedelta(days=i) for i in range(3))
    data = InputData((Employee("A", "職員A", 0),), tuple(DayRequirement(day, 0) for day in days),
                     requests=(LeaveCandidate("R", "A", days[1], 1),))
    rows = build_request_heatmap(data)
    assert [row.day for row in rows] == list(days)
    assert [row.request_employee_count for row in rows] == [0, 1, 0]
    assert all(row.adopted_employee_count is None for row in rows)
    result = solve_schedule(data)
    after = build_request_heatmap(data, result.schedule)
    assert [row.adopted_employee_count for row in after] == [0, 1, 0]


def test_thirty_people_and_duplicate_representations_count_distinct_staff_only():
    employees = tuple(Employee(f"{i:03d}", f"職員{i}", 0) for i in range(30))
    metadata = tuple(IndependentLeaveRequest(f"R{i}", e.id, D7, PUBLIC_REQUEST) for i, e in enumerate(employees))
    candidates = tuple(LeaveCandidate(r.request_id, r.employee_id, r.day, 1) for r in metadata)
    data = InputData(employees, (DayRequirement(D7, 0),), requests=candidates,
                     independent_requests=metadata, request_format="independent")
    row = build_request_heatmap(data)[0]
    assert row.request_employee_count == row.public_request_employee_count == row.candidate_count == 30
    assert row.entries[0].employee_id == "000"
    assert all(entry.priority == 1 for entry in row.entries)


def test_infeasible_data_without_a_schedule_keeps_requests_uncalculated():
    data = replace(sample_data(), fixed_off=(FixedOff("A", D7), FixedOff("B", D7)))
    result = solve_schedule(data)
    assert result.schedule is None
    rows = build_request_heatmap(data, result.schedule)
    assert [row.request_employee_count for row in rows] == [2, 1, 1]
    assert all(entry.status == UNCALCULATED for row in rows for entry in row.entries)
    assert all(row.adopted_employee_count is None for row in rows)


def test_raw_ratio_mismatch_and_explicit_fallback_are_marked_without_changing_counts():
    data = replace(sample_data(), settings=ShiftSettings(weekday_ratios=(Decimal(100),) * 7))
    rows = build_request_heatmap(data)
    assert all(row.required_count == 2 and row.is_reference and row.warnings for row in rows)
    assert all(row.required == 2 for row in data.requirements)
    explicit = build_request_heatmap(sample_data(), raw_fallback=True)
    assert all(row.is_reference and any("設定適用前" in w for w in row.warnings) for row in explicit)


def test_outside_period_request_is_visible_as_reference_and_unknown_staff_is_named_as_unknown():
    day = D9 + timedelta(days=1)
    data = replace(sample_data(), requests=(LeaveCandidate("OUTSIDE", "UNKNOWN", day, 1),))
    rows = _index(build_request_heatmap(data))
    assert rows[day].in_period is False and rows[day].required_count is None
    assert rows[day].is_reference
    assert rows[day].entries[0].employee_name == "氏名不明"
    assert any("対象期間外" in warning for warning in rows[day].warnings)


def test_incomplete_schedule_is_not_reported_as_zero_staff():
    data = sample_data()
    row = build_request_heatmap(data, {})[0]
    assert row.current_staff_count is None and row.is_reference


def test_evaluation_without_a_current_schedule_does_not_reuse_adoption_status():
    data = sample_data()
    result = solve_schedule(data)
    evaluation = evaluate_schedule(data, result.schedule)
    rows = build_request_heatmap(data, evaluation=evaluation)
    assert all(entry.status == UNCALCULATED for row in rows for entry in row.entries)
    assert all(row.current_staff_count is None for row in rows)


def test_results_are_immutable_and_do_not_mutate_input_or_schedule():
    data = sample_data()
    result = solve_schedule(data)
    schedule = dict(result.schedule)
    before = dict(schedule)
    row = build_request_heatmap(data, schedule)[0]
    assert schedule == before
    with pytest.raises(FrozenInstanceError):
        row.request_employee_count = 999
    with pytest.raises(FrozenInstanceError):
        row.entries[0].status = "変更"
    assert row.entries[0].id == row.entries[0].request_id


def test_old_evaluation_with_same_daily_staff_count_cannot_hide_employee_violation():
    data = sample_data()
    result = solve_schedule(data)
    old = evaluate_schedule(data, result.schedule)
    broken = dict(result.schedule)
    broken["A", D7] = True
    broken["B", D7] = False
    assert sum(broken[e.id, D7] for e in data.employees) == old.daily_counts[D7]
    assert not evaluate_schedule(data, broken).valid
    rows = build_request_heatmap(data, broken, old)
    assert all(row.is_reference for row in rows)
    assert all(any("現在の案" in warning for warning in row.warnings) for row in rows)


def test_shift_kind_change_with_unchanged_daily_count_is_a_reference_when_evaluation_is_old():
    from sinta_shift.models import TWO_SHIFT
    data = configure_input(sample_data(), ShiftSettings(mode=TWO_SHIFT))
    result = solve_schedule(data)
    old = evaluate_schedule(data, result.schedule)
    broken = dict(result.schedule)
    key = next(key for key, value in broken.items() if value == "夜勤")
    broken[key] = "日勤"
    rows = build_request_heatmap(data, broken, old)
    assert all(row.is_reference for row in rows)
    assert _index(rows)[key[1]].current_staff_count == old.daily_counts[key[1]]
