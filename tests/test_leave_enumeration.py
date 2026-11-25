"""小さい勤務表を独立に全列挙し、公休・有給・明けの制約と目的を照合する。"""
from collections import Counter, defaultdict
from datetime import date, timedelta
from itertools import product

import pytest

from sinta_shift.models import (
    DAY_ONLY, TWO_SHIFT, THREE_SHIFT, SHIFT_MODES, PUBLIC_OFF, PAID_LEAVE, RECOVERY,
    ShiftSettings, ShiftTime, Employee, DayRequirement, InputData, PaidLeave,
    LeaveCandidate, SolverStatus,
)
from sinta_shift.solver import solve_schedule
from sinta_shift.evaluation import evaluate_schedule


def _matching_score(rows, candidates, days):
    groups = defaultdict(list)
    for candidate in candidates:
        groups[candidate.request_id].append(candidate)
    best = (0, 0)
    for choices in product(*[(None, *values) for values in groups.values()]):
        selected = [value for value in choices if value is not None]
        keys = [(value.employee_id, value.day) for value in selected]
        if len(keys) != len(set(keys)):
            continue
        if any(rows[c.employee_id][days.index(c.day)] != PUBLIC_OFF for c in selected):
            continue
        best = max(best, (len(selected), -sum(c.priority for c in selected)))
    return best


def _independent_best(data, crossing):
    """本体のconfigure/validation/ラベル判定を使わない暦日全探索。"""
    days = data.days
    active = SHIFT_MODES[data.settings.mode]
    paid = {(item.employee_id, item.day) for item in data.paid_leaves}
    possibilities = []
    for employee in data.employees:
        valid_rows = []
        for row in product((*active, PUBLIC_OFF, PAID_LEAVE, RECOVERY), repeat=len(days)):
            if any((label == PAID_LEAVE) != ((employee.id, day) in paid)
                   for label, day in zip(row, days)):
                continue
            previous = employee.previous_shift or PUBLIC_OFF
            run = employee.previous_work_days
            valid = True
            for day, label in zip(days, row):
                if (label == RECOVERY) != (previous in crossing):
                    valid = False
                if data.settings.night_rest and previous in ("夜勤", "深夜勤") and label in active:
                    valid = False
                if (previous, label) in data.settings.forbidden_transitions:
                    valid = False
                if any(item.employee_id == employee.id and item.day == day
                       and label != PUBLIC_OFF for item in data.fixed_off):
                    valid = False
                run = run + 1 if label in active or label == RECOVERY else 0
                if data.settings.max_consecutive_days and run > data.settings.max_consecutive_days:
                    valid = False
                previous = label
            if not valid:
                continue
            if data.settings.day_count_mode == "公休日数から算出":
                target = (employee.public_holidays if employee.public_holidays is not None
                          else data.settings.default_public_holidays)
                if row.count(PUBLIC_OFF) != target:
                    continue
            elif sum(label in active for label in row) != employee.work_days:
                continue
            valid_rows.append(row)
        possibilities.append(valid_rows)
    best = None
    for selected in product(*possibilities):
        rows = dict(zip((employee.id for employee in data.employees), selected))
        if any(any(Counter(row[index] for row in selected)[label] != required
                       for label, required in requirement.shift_counts)
               for index, requirement in enumerate(data.requirements)):
            continue
        score = _matching_score(rows, data.requests, days)
        best = score if best is None else max(best, score)
    return best


@pytest.mark.parametrize("mode", [DAY_ONLY, TWO_SHIFT, THREE_SHIFT])
@pytest.mark.parametrize("count_mode", ["勤務日数を指定", "公休日数から算出"])
def test_leave_optimization_matches_independent_enumeration(mode, count_mode):
    if mode == DAY_ONLY:
        examples = [(PAID_LEAVE, PUBLIC_OFF, "日勤"), ("日勤", "日勤", PUBLIC_OFF)]
        times, crossing = (), set()
    elif mode == TWO_SHIFT:
        examples = [("夜勤", RECOVERY, PUBLIC_OFF, "日勤"),
                    ("日勤", PAID_LEAVE, "夜勤", RECOVERY)]
        times = (ShiftTime("日勤", 480, 1020), ShiftTime("夜勤", 960, 1980))
        crossing = {"夜勤"}
    else:
        examples = [("日勤", PUBLIC_OFF, "深夜勤", PUBLIC_OFF),
                    ("準夜勤", PUBLIC_OFF, "日勤", PAID_LEAVE)]
        times = (ShiftTime("日勤", 480, 960), ShiftTime("準夜勤", 960, 1440),
                 ShiftTime("深夜勤", 0, 480))
        crossing = set()
    days = tuple(date(2026, 11, 7) + timedelta(days=index) for index in range(len(examples[0])))
    active = SHIFT_MODES[mode]
    employees = tuple(Employee(str(index), f"職員{index}",
                               sum(label in active for label in row) if count_mode == "勤務日数を指定" else None,
                               public_holidays=row.count(PUBLIC_OFF))
                      for index, row in enumerate(examples))
    requirements = tuple(DayRequirement(day, sum(row[index] in active for row in examples),
                                       shift_counts=tuple((label, sum(row[index] == label for row in examples))
                                                          for label in active))
                         for index, day in enumerate(days))
    leaves = tuple(PaidLeave(str(index), day) for index, row in enumerate(examples)
                   for day, label in zip(days, row) if label == PAID_LEAVE)
    requests = tuple(LeaveCandidate(f"R{index}-{offset % 2}", str(index), day, offset + 1)
                     for index in range(2) for offset, day in enumerate(days))
    settings = ShiftSettings(mode=mode, max_consecutive_days=3, night_rest=True,
                             forbidden_transitions=(("準夜勤", "日勤"),) if mode == THREE_SHIFT else (),
                             leave_enabled=True, day_count_mode=count_mode,
                             shift_times=times, shift_times_confirmed=True)
    data = InputData(employees, requirements, requests=requests, settings=settings, paid_leaves=leaves)
    expected = _independent_best(data, crossing)
    assert expected is not None
    result = solve_schedule(data)
    assert result.status == SolverStatus.OPTIMAL, result.message
    assert result.schedule is not None
    evaluation = evaluate_schedule(data, result.schedule)
    assert evaluation.valid, [item.message for item in evaluation.issues]
    assert (evaluation.adopted_count, -evaluation.priority_sum) == expected
