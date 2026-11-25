"""小規模な交代制を独立した全探索と照合する受入検証。"""
from datetime import date
from itertools import permutations, product

import pytest

from sinta_shift.models import (
    TWO_SHIFT, THREE_SHIFT, OFF, DayRequirement, Employee, InputData,
    LeaveCandidate, ShiftSettings, SolverStatus,
)
from sinta_shift.evaluation import evaluate_schedule
from sinta_shift.solver import solve_schedule


def objective_from_requests(data, schedule):
    grouped = {}
    for request in data.requests:
        grouped.setdefault(request.request_id, []).append(request)
    best = (0, 0)
    for selected in product(*([None, *items] for items in grouped.values())):
        adopted = [item for item in selected if item is not None]
        cells = [(item.employee_id, item.day) for item in adopted]
        if len(cells) != len(set(cells)) or any(schedule[cell] != OFF for cell in cells):
            continue
        best = max(best, (len(adopted), -sum(item.priority for item in adopted)))
    return best


@pytest.mark.parametrize("mode,coverage", [
    (TWO_SHIFT, (("日勤", 2), ("夜勤", 1))),
    (THREE_SHIFT, (("日勤", 1), ("準夜勤", 1), ("深夜勤", 1))),
])
def test_multishift_optimum_matches_independent_complete_enumeration(mode, coverage):
    days = (date(2026, 11, 7), date(2026, 11, 8))
    employees = (
        Employee("A", "職員A", 2), Employee("B", "職員B", 2),
        Employee("C", "職員C", 1, previous_work_days=2, previous_shift="日勤"), Employee("D", "職員D", 1),
    )
    data = InputData(employees,
        tuple(DayRequirement(day, 3, shift_counts=coverage) for day in days),
        requests=(LeaveCandidate("R1", "C", days[0], 1),
                  LeaveCandidate("R2", "D", days[0], 1), LeaveCandidate("R2", "D", days[1], 2),
                  LeaveCandidate("R3", "C", days[0], 1)),
        settings=ShiftSettings(mode=mode, max_consecutive_days=2,
                               forbidden_transitions=(("準夜勤", "日勤"),) if mode == THREE_SHIFT else ()))
    daily_slots = [label for label, count in coverage for _ in range(count)] + [OFF]
    daily_assignments = sorted(set(permutations(daily_slots)))
    best = (-1, 0)
    feasible = 0
    for candidates in product(daily_assignments, repeat=2):
        schedule = {(employee.id, day): candidates[day_index][employee_index]
                    for employee_index, employee in enumerate(employees)
                    for day_index, day in enumerate(days)}
        if any(sum(schedule[e.id, day] != OFF for day in days) != e.work_days for e in employees):
            continue
        valid = True
        for employee in employees:
            run = employee.previous_work_days
            previous = employee.previous_shift or OFF
            for day in days:
                current = schedule[employee.id, day]
                if previous in ("夜勤", "深夜勤") and current != OFF:
                    valid = False
                if (previous, current) in data.settings.forbidden_transitions:
                    valid = False
                run = run + 1 if current != OFF else 0
                if run > 2:
                    valid = False
                previous = current
        if valid:
            feasible += 1
            best = max(best, objective_from_requests(data, schedule))
    assert feasible > 0 and best == (2, -3)
    result = solve_schedule(data, time_limit=30)
    assert result.status == SolverStatus.OPTIMAL, result.message
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid, [i.message for i in evaluation.issues]
    assert (evaluation.adopted_count, -evaluation.priority_sum) == best
    assert result.schedule["C", days[0]] == OFF
    assert all(evaluation.daily_shift_counts[day] == dict(coverage) for day in days)
