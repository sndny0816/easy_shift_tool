"""架空の受入サンプル。結果は格納せず、条件だけを定義する。"""
from datetime import date
from .models import DayRequirement, Employee, InputData, LeaveCandidate


def sample_data() -> InputData:
    return InputData(
        employees=tuple(Employee(i, f"職員{i}", 2) for i in "ABC"),
        requirements=tuple(DayRequirement(date(2026, 11, d), 2) for d in (7, 8, 9)),
        requests=(
            LeaveCandidate("R1", "A", date(2026, 11, 7), 1),
            LeaveCandidate("R2", "B", date(2026, 11, 7), 1),
            LeaveCandidate("R2", "B", date(2026, 11, 8), 2),
            LeaveCandidate("R3", "C", date(2026, 11, 9), 1),
        ),
    )
