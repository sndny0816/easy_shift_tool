"""横形式の独立休暇申請と、有給を明示確認する業務境界。"""
from dataclasses import replace
from datetime import date, timedelta

import pytest

from sinta_shift.models import (
    Employee, DayRequirement, PaidLeave, LeaveCandidate, IndependentLeaveRequest,
    InputData, ShiftSettings, InputError, Source, SolverStatus, PUBLIC_DAYS,
    PUBLIC_OFF, PAID_LEAVE, PUBLIC_REQUEST, PAID_REQUEST,
)
from sinta_shift.planning import configure_input, confirm_paid_requests
from sinta_shift.validation import validate_input
from sinta_shift.solver import solve_schedule
from sinta_shift.evaluation import evaluate_schedule


START = date(2026, 11, 7)


def _request(request_id, offset, leave_type=PUBLIC_REQUEST, employee_id="A"):
    return IndependentLeaveRequest(request_id, employee_id, START + timedelta(days=offset), leave_type,
                                   Source("wide.xlsx", "希望休", 2, f"{chr(66 + offset * 2)}2"))


def _data(requests, *, days=3, public=2, base_paid=()):
    return InputData(
        employees=(Employee("A", "職員A", None),),
        requirements=tuple(DayRequirement(START + timedelta(days=i), 0) for i in range(days)),
        requests=tuple(LeaveCandidate(r.request_id, r.employee_id, r.day, 1, r.source)
                       for r in requests if r.leave_type == PUBLIC_REQUEST),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS,
                               default_public_holidays=public),
        paid_leaves=base_paid,
        independent_requests=tuple(requests),
        request_format="independent",
    )


def _mixed():
    return _data((_request("P1", 0), _request("P2", 1), _request("A1", 2, PAID_REQUEST)))


def test_three_dates_are_three_independent_requests_of_equal_rank():
    data = _data(tuple(_request(f"R{i + 1}", i) for i in range(3)), public=3)
    result = solve_schedule(data)
    assert result.status is SolverStatus.OPTIMAL
    assert all(value == PUBLIC_OFF for value in result.schedule.values())
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid
    assert len(evaluation.request_results) == evaluation.adopted_count == 3
    assert evaluation.priority_sum == 3
    assert {r.adopted_day for r in evaluation.request_results} == set(data.days)
    assert evaluation.paid_request_results == []


def test_paid_request_stops_calculation_until_explicit_confirmation_with_source_cell():
    data = _mixed()
    effective = configure_input(data)
    assert effective.paid_leaves == ()
    issues = validate_input(effective)
    pending = next(issue for issue in issues if issue.kind == "有給申請未確認")
    assert (pending.source.file, pending.source.sheet, pending.source.row, pending.source.column) == ("wide.xlsx", "希望休", 2, "F2")
    result = solve_schedule(effective)
    assert result.status is SolverStatus.INFEASIBLE
    assert result.schedule is None
    evaluation = evaluate_schedule(effective, {})
    assert not evaluation.valid
    assert evaluation.adopted_count == 0
    assert len(evaluation.paid_request_results) == 1
    assert evaluation.paid_request_results[0].status == "確認待ち"
    assert evaluation.paid_request_results[0].adopted_day is None
    assert evaluation.paid_request_results[0].priority is None


def test_confirmed_paid_is_fixed_and_never_included_in_public_request_score():
    data = confirm_paid_requests(_mixed(), {"A1"})
    assert data.confirmed_paid_request_ids == frozenset({"A1"})
    assert len(data.paid_leaves) == 1
    paid = data.paid_leaves[0]
    assert (paid.employee_id, paid.day, paid.request_id) == ("A", START + timedelta(days=2), "A1")
    assert paid.source == data.independent_requests[2].source
    effective = configure_input(data)
    assert effective.employees[0].work_days == 0
    assert configure_input(effective) == effective
    result = solve_schedule(effective)
    assert result.status is SolverStatus.OPTIMAL
    assert result.schedule["A", paid.day] == PAID_LEAVE
    evaluation = evaluate_schedule(effective, result.schedule)
    assert evaluation.valid and evaluation.adopted_count == 2
    assert len(evaluation.request_results) == 2
    assert len(evaluation.paid_request_results) == 1
    receipt = evaluation.paid_request_results[0]
    assert (receipt.request_id, receipt.status, receipt.adopted_day, receipt.priority) == ("A1", "有給確定", paid.day, None)
    broken = dict(result.schedule)
    broken["A", paid.day] = PUBLIC_OFF
    edited = evaluate_schedule(effective, broken)
    assert not edited.valid
    assert any(issue.kind == "有給違反" and issue.day == paid.day for issue in edited.issues)


def test_confirmation_is_idempotent_and_reversal_removes_only_derived_paid():
    base = PaidLeave("A", START + timedelta(days=3), Source("wide.xlsx", "有給", 2))
    data = _data(_mixed().independent_requests, days=4, base_paid=(base,))
    confirmed = confirm_paid_requests(data, {"A1"})
    assert confirmed.paid_leaves == (base, PaidLeave("A", START + timedelta(days=2), data.independent_requests[2].source, "A1"))
    assert confirm_paid_requests(confirmed, {"A1"}) == confirmed
    effective = configure_input(confirmed)
    result = solve_schedule(effective)
    evaluation = evaluate_schedule(effective, result.schedule)
    assert evaluation.valid and evaluation.employee_paid_counts == {"A": 2}
    assert len(evaluation.paid_request_results) == 1  # 元の固定有給は新申請ではない。
    unconfirmed = confirm_paid_requests(effective, set())
    assert unconfirmed.paid_leaves == (base,)
    assert unconfirmed.confirmed_paid_request_ids == frozenset()
    assert configure_input(unconfirmed).employees[0].work_days == 1
    assert any(issue.kind == "有給申請未確認" for issue in validate_input(unconfirmed))
    assert solve_schedule(unconfirmed).schedule is None


def test_partial_confirmation_keeps_remaining_paid_requests_pending():
    data = _data((_request("P1", 0), _request("A1", 1, PAID_REQUEST), _request("A2", 2, PAID_REQUEST)), public=1)
    partial = confirm_paid_requests(data, {"A1"})
    assert len(partial.paid_leaves) == 1
    pending = [issue for issue in validate_input(partial) if issue.kind == "有給申請未確認"]
    assert len(pending) == 1 and pending[0].day == START + timedelta(days=2)
    assert solve_schedule(partial).schedule is None
    complete = confirm_paid_requests(partial, {"A1", "A2"})
    result = solve_schedule(complete)
    evaluation = evaluate_schedule(complete, result.schedule)
    assert evaluation.valid and evaluation.adopted_count == 1
    assert [row.status for row in evaluation.paid_request_results] == ["有給確定", "有給確定"]


@pytest.mark.parametrize("ids", [{"UNKNOWN"}, {"P1"}, "A1", {1}, {""}])
def test_unknown_public_or_badly_typed_confirmation_ids_are_rejected(ids):
    with pytest.raises(InputError):
        confirm_paid_requests(_mixed(), ids)


def test_configure_applies_confirmation_ids_before_counting_paid_days():
    data = replace(_mixed(), confirmed_paid_request_ids=frozenset({"A1"}))
    assert data.paid_leaves == ()
    effective = configure_input(data)
    assert len(effective.paid_leaves) == 1
    assert effective.employees[0].work_days == 0
    assert validate_input(effective) == []


@pytest.mark.parametrize("change,reason", [
    (lambda r: replace(r, employee_id="UNKNOWN"), "未知の職員ID"),
    (lambda r: replace(r, day=START - timedelta(days=1)), "対象期間外"),
    (lambda r: replace(r, leave_type="有給"), "区分"),
    (lambda r: replace(r, request_id=""), "申請ID"),
])
def test_invalid_independent_request_keeps_original_cell(change, reason):
    request = change(_request("R1", 0))
    data = _data((request,), public=3)
    issues = validate_input(data)
    assert any(reason in issue.message and issue.source == request.source for issue in issues)
    assert solve_schedule(data).schedule is None


@pytest.mark.parametrize("types", [(PUBLIC_REQUEST, PUBLIC_REQUEST), (PUBLIC_REQUEST, PAID_REQUEST)])
def test_duplicate_employee_day_is_rejected_even_between_leave_types(types):
    data = _data((_request("R1", 0, types[0]), _request("R2", 0, types[1])), public=3)
    assert any("同じ職員・同じ日" in issue.message for issue in validate_input(data))


def test_duplicate_independent_request_id_is_rejected():
    data = _data((_request("R1", 0), _request("R1", 1)), public=3)
    assert any("独立申請ID" in issue.message and "重複" in issue.message for issue in validate_input(data))


def test_wide_public_request_cannot_overlap_fixed_paid_but_legacy_alternative_can():
    paid = PaidLeave("A", START, Source("wide.xlsx", "有給", 2))
    data = _data((_request("P1", 0),), public=2, base_paid=(paid,))
    assert any("独立した公休希望と固定有給が重複" in issue.message for issue in validate_input(data))
    legacy = replace(data, independent_requests=(), request_format="alternatives")
    result = solve_schedule(legacy)
    assert result.status is SolverStatus.OPTIMAL
    evaluation = evaluate_schedule(legacy, result.schedule)
    assert evaluation.valid and evaluation.adopted_count == 0


def test_wide_paid_and_base_paid_same_day_are_not_merged_or_double_counted():
    data = _mixed()
    base = PaidLeave("A", data.independent_requests[2].day, Source("wide.xlsx", "有給", 2))
    data = replace(data, paid_leaves=(base,))
    assert any("有給申請と元の有給一覧が重複" in issue.message for issue in validate_input(data))
    confirmed = confirm_paid_requests(data, {"A1"})
    assert len(confirmed.paid_leaves) == 2  # 原入力を黙って修復・重複排除しない。
    assert any("同じ職員・日付の有給が重複" in issue.message for issue in validate_input(confirmed))


def test_paid_request_cannot_be_disguised_as_soft_public_candidate():
    data = _mixed()
    paid = data.independent_requests[2]
    data = replace(data, requests=(*data.requests, LeaveCandidate(paid.request_id, paid.employee_id, paid.day, 1, paid.source)))
    assert any("有給申請を未採用にできる希望休候補" in issue.message for issue in validate_input(data))


@pytest.mark.parametrize("violation", ["missing", "other_day", "priority", "alternative"])
def test_independent_public_metadata_must_match_exactly_one_rank_one_candidate(violation):
    data = _data((_request("R1", 0),), public=3)
    first = data.requests[0]
    if violation == "missing":
        candidates = ()
    elif violation == "other_day":
        candidates = (replace(first, day=START + timedelta(days=1)),)
    elif violation == "priority":
        candidates = (replace(first, priority=2),)
    else:
        candidates = (first, replace(first, day=START + timedelta(days=1), priority=2))
    data = replace(data, requests=candidates)
    assert any("入力日の1候補・同順位1" in issue.message for issue in validate_input(data))


def test_empty_independent_format_remains_explicit_and_unknown_format_is_rejected():
    data = _data((), public=3)
    assert configure_input(data).request_format == "independent"
    assert configure_input(replace(data, request_format="alternatives")).request_format == "alternatives"
    public = _data((_request("P1", 0),), public=3)
    assert configure_input(replace(public, request_format="alternatives")).request_format == "independent"
    assert validate_input(replace(data, request_format="dindependent"))
