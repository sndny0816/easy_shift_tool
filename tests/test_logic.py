"""勤務表仕様に対する独立した受入テスト。GUIやExcelは使わない。"""
from dataclasses import replace
from datetime import date, timedelta
from itertools import product

import pytest

from sinta_shift.evaluation import evaluate_schedule
from sinta_shift.models import (
    DayRequirement,
    Employee,
    FixedOff,
    InputData,
    LeaveCandidate,
    SolverStatus,
    Source,
)
from sinta_shift.sample import sample_data
from sinta_shift.solver import solve_schedule
from sinta_shift.validation import validate_input


D7 = date(2026, 11, 7)
D8 = date(2026, 11, 8)
D9 = date(2026, 11, 9)


def _assert_valid_solution(data, result):
    assert result.status in (SolverStatus.OPTIMAL, SolverStatus.FEASIBLE)
    assert result.schedule is not None
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid, [i.message for i in evaluation.issues]
    assert len(result.schedule) == len(data.employees) * len(data.days)
    for employee in data.employees:
        assert sum(result.schedule[employee.id, day] for day in data.days) == employee.work_days
    for requirement in data.requirements:
        assert sum(result.schedule[e.id, requirement.day] for e in data.employees) == requirement.required
    for fixed in data.fixed_off:
        assert not result.schedule[fixed.employee_id, fixed.day]
    return evaluation


def test_sample_acceptance_exact_schedule_and_second_choice():
    data = sample_data()
    assert not validate_input(data)
    result = solve_schedule(data)
    evaluation = _assert_valid_solution(data, result)
    assert result.status is SolverStatus.OPTIMAL
    assert result.count_optimal and result.priority_optimal
    expected_off = {("A", D7), ("B", D8), ("C", D9)}
    assert {cell for cell, work in result.schedule.items() if not work} == expected_off
    assert evaluation.adopted_count == 3
    assert evaluation.priority_sum == 4
    requests = {r.request_id: r for r in evaluation.request_results}
    assert requests["R2"].adopted_day == D8
    assert requests["R2"].priority == 2
    assert "2" in requests["R2"].status or "２" in requests["R2"].status


def test_fixed_off_capacity_conflict_is_infeasible_with_clear_reason():
    data = replace(sample_data(), fixed_off=(FixedOff("A", D7), FixedOff("B", D7)))
    result = solve_schedule(data)
    assert result.status is SolverStatus.INFEASIBLE
    assert result.schedule is None
    detail = result.message + "\n" + "\n".join(i.message for i in result.issues)
    assert "必要2人" in detail
    assert "勤務可能1人" in detail


def test_competing_saturday_requests_keep_hard_constraints_and_show_two_unadopted():
    data = replace(sample_data(), requests=tuple(
        LeaveCandidate(f"R{n}", employee_id, D7, 1)
        for n, employee_id in enumerate(("A", "B", "C"), 1)
    ))
    result = solve_schedule(data)
    evaluation = _assert_valid_solution(data, result)
    assert evaluation.adopted_count == 1
    assert len(evaluation.request_results) == 3
    assert sum(r.adopted_day is None for r in evaluation.request_results) == 2
    assert all("未採用" in r.status for r in evaluation.request_results if r.adopted_day is None)


def test_manual_change_reports_employee_and_day():
    data = sample_data()
    result = solve_schedule(data)
    _assert_valid_solution(data, result)
    schedule = dict(result.schedule)
    schedule["A", D7] = True
    evaluation = evaluate_schedule(data, schedule)
    assert not evaluation.valid
    assert any(i.day == D7 and "人数" in i.message for i in evaluation.issues)
    assert any(i.employee_id == "A" and "勤務" in i.message for i in evaluation.issues)
    assert any(i.day == D7 and i.employee_id == "A" for i in evaluation.issues)
    assert evaluation.adopted_count < 3


def test_manual_change_violates_fixed_off():
    data = replace(sample_data(), fixed_off=(FixedOff("A", D7),))
    result = solve_schedule(data)
    _assert_valid_solution(data, result)
    schedule = dict(result.schedule)
    schedule["A", D7] = True
    evaluation = evaluate_schedule(data, schedule)
    assert not evaluation.valid
    assert any(i.employee_id == "A" and i.day == D7 and "確定休" in i.message for i in evaluation.issues)


def _invalid_cases():
    data = sample_data()
    return [
        ("duplicate_staff", replace(data, employees=(data.employees[0], data.employees[0], data.employees[2]))),
        ("unknown_fixed_id", replace(data, fixed_off=(FixedOff("UNKNOWN", D7),))),
        ("unknown_request_id", replace(data, requests=(LeaveCandidate("R", "UNKNOWN", D7, 1),))),
        ("fixed_outside", replace(data, fixed_off=(FixedOff("A", D7 - timedelta(days=1)),))),
        ("request_outside", replace(data, requests=(LeaveCandidate("R", "A", D9 + timedelta(days=1), 1),))),
        ("duplicate_candidate", replace(data, requests=(data.requests[0], data.requests[0]))),
        ("negative_requirement", replace(data, requirements=(replace(data.requirements[0], required=-1), *data.requirements[1:]))),
        ("negative_workdays", replace(data, employees=(replace(data.employees[0], work_days=-1), *data.employees[1:]))),
        ("too_many_workdays", replace(data, employees=(replace(data.employees[0], work_days=4), *data.employees[1:]))),
        ("too_many_required", replace(data, requirements=(replace(data.requirements[0], required=4), *data.requirements[1:]))),
        ("workdays_sum_mismatch", replace(data, employees=(replace(data.employees[0], work_days=1), *data.employees[1:]))),
        ("date_gap", replace(data, requirements=(data.requirements[0], data.requirements[2]))),
        ("duplicate_day", replace(data, requirements=(data.requirements[0], data.requirements[0], data.requirements[2]))),
        ("nonpositive_priority", replace(data, requests=(replace(data.requests[0], priority=0),))),
        ("request_shared_by_two_people", replace(data, requests=(LeaveCandidate("R", "A", D7, 1), LeaveCandidate("R", "B", D8, 2)))),
    ]


@pytest.mark.parametrize("case,data", _invalid_cases(), ids=lambda value: value if isinstance(value, str) else None)
def test_invalid_input_detected_before_solve(case, data):
    issues = validate_input(data)
    assert issues, case
    assert all(i.message for i in issues)


def test_input_issues_retain_source_location():
    data = sample_data()
    source = Source("broken.xlsx", "希望休", 8, "職員ID")
    data = replace(data, requests=(LeaveCandidate("R", "UNKNOWN", D7, 1, source),))
    issues = validate_input(data)
    assert any(i.source.file == "broken.xlsx" and i.source.sheet == "希望休" and i.source.row == 8 for i in issues)


def test_multiple_people_can_be_off_on_same_day():
    data = InputData(
        employees=tuple(Employee(str(n), f"職員{n}", 1) for n in range(4)),
        requirements=(DayRequirement(D7, 2), DayRequirement(D8, 2)),
    )
    assert not validate_input(data)
    result = solve_schedule(data)
    evaluation = _assert_valid_solution(data, result)
    assert evaluation.adopted_count == 0
    assert all(sum(not result.schedule[e.id, day] for e in data.employees) == 2 for day in data.days)


def test_same_off_day_cannot_credit_two_requests():
    data = replace(sample_data(), requests=(
        LeaveCandidate("R1", "A", D7, 1),
        LeaveCandidate("R2", "A", D7, 1),
    ))
    result = solve_schedule(data)
    evaluation = _assert_valid_solution(data, result)
    assert evaluation.adopted_count == 1
    assert sum(r.adopted_day is None for r in evaluation.request_results) == 1


def test_re_evaluation_matches_alternatives_globally_instead_of_greedily():
    data = InputData(
        employees=(Employee("A", "職員A", 1), Employee("B", "職員B", 2)),
        requirements=tuple(DayRequirement(day, 1) for day in (D7, D8, D9)),
        requests=(
            LeaveCandidate("R1", "A", D7, 1),
            LeaveCandidate("R1", "A", D8, 2),
            LeaveCandidate("R2", "A", D7, 1),
        ),
    )
    # Aの7日と8日が休み。R1を先に7日へ割り当てるとR2を不当に落としてしまう。
    schedule = {
        (employee.id, day): (day == D9 if employee.id == "A" else day != D9)
        for employee in data.employees
        for day in data.days
    }
    evaluation = evaluate_schedule(data, schedule)
    assert evaluation.valid
    assert evaluation.adopted_count == 2
    assert evaluation.priority_sum == 3
    result_by_id = {r.request_id: r for r in evaluation.request_results}
    assert result_by_id["R1"].adopted_day == D8
    assert result_by_id["R2"].adopted_day == D7


def test_number_of_adopted_requests_has_priority_over_any_rank_sum():
    data = sample_data()
    data = replace(data, requests=tuple(
        replace(r, priority=1_000_000) if r.request_id == "R2" and r.day == D8 else r
        for r in data.requests
    ))
    result = solve_schedule(data)
    evaluation = _assert_valid_solution(data, result)
    assert result.status is SolverStatus.OPTIMAL
    assert evaluation.adopted_count == 3
    assert evaluation.priority_sum == 1_000_002


def test_no_requests_and_no_fixed_off_are_supported():
    data = replace(sample_data(), fixed_off=(), requests=())
    result = solve_schedule(data)
    evaluation = _assert_valid_solution(data, result)
    assert evaluation.request_results == []
    assert evaluation.adopted_count == evaluation.priority_sum == 0


def test_tiny_budget_means_unknown_not_proven_infeasible():
    result = solve_schedule(sample_data(), time_limit=1e-9)
    assert result.status is SolverStatus.UNKNOWN
    assert result.schedule is None


def _independent_request_objective(data, schedule):
    """申請ごとに未採用か1候補を全探索する。CP-SAT・本体再評価を使わない。"""
    grouped = {}
    for candidate in data.requests:
        grouped.setdefault(candidate.request_id, []).append(candidate)
    best = (0, 0)
    for selected in product(*([None, *rows] for rows in grouped.values())):
        candidates = [r for r in selected if r is not None]
        rest_cells = [(r.employee_id, r.day) for r in candidates]
        if len(rest_cells) != len(set(rest_cells)):
            continue
        if any(schedule[cell] for cell in rest_cells):
            continue
        score = (len(candidates), -sum(r.priority for r in candidates))
        best = max(best, score)
    return best


def test_sample_optimum_matches_independent_exhaustive_search():
    data = sample_data()
    cells = [(employee.id, day) for employee in data.employees for day in data.days]
    best = (-1, 0)
    feasible_count = 0
    for bits in product((False, True), repeat=len(cells)):
        schedule = dict(zip(cells, bits))
        if any(sum(schedule[employee.id, day] for day in data.days) != employee.work_days for employee in data.employees):
            continue
        if any(sum(schedule[employee.id, r.day] for employee in data.employees) != r.required for r in data.requirements):
            continue
        if any(schedule[fixed.employee_id, fixed.day] for fixed in data.fixed_off):
            continue
        feasible_count += 1
        best = max(best, _independent_request_objective(data, schedule))
    assert feasible_count == 6
    assert best == (3, -4)
    result = solve_schedule(data)
    evaluation = _assert_valid_solution(data, result)
    assert (evaluation.adopted_count, -evaluation.priority_sum) == best
