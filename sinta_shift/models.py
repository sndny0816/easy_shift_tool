"""GUIやExcelに依存しない業務データ。日数・人数は整数単位。"""
from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from decimal import Decimal

DAY_ONLY = "日勤のみ"
TWO_SHIFT = "２交代制"
THREE_SHIFT = "３交代制"
OFF = "休み"
PUBLIC_OFF = "公休"
PAID_LEAVE = "有給"
RECOVERY = "明け"
DIRECT_DAYS = "勤務日数を指定"
PUBLIC_DAYS = "公休日数から算出"
PUBLIC_REQUEST = "公休希望"
PAID_REQUEST = "有給申請"
SHIFT_MODES = {DAY_ONLY: ("日勤",), TWO_SHIFT: ("日勤", "夜勤"), THREE_SHIFT: ("日勤", "準夜勤", "深夜勤")}


@dataclass(frozen=True)
class ShiftTime:
    label: str
    start_minute: int
    end_minute: int  # 開始日0時基準。翌9:00は1980。
    @property
    def crosses_midnight(self) -> bool:
        return self.end_minute > 1440


@dataclass(frozen=True)
class ShiftSettings:
    mode: str = DAY_ONLY
    max_consecutive_days: int | None = None
    weekday_ratios: tuple[Decimal, ...] | None = None  # 月〜日の0〜100％。NoneならExcel人数。
    shift_weights: tuple[int, ...] = ()  # 有効勤務区分の配分。空なら均等。
    night_rest: bool = True  # 夜勤・深夜勤の翌日は休み。
    forbidden_transitions: tuple[tuple[str, str], ...] = ()
    leave_enabled: bool = False
    day_count_mode: str = DIRECT_DAYS
    default_public_holidays: int | None = None
    shift_times: tuple[ShiftTime, ...] = ()
    shift_times_confirmed: bool = False
    weekly_ratios: tuple[tuple[date, Decimal], ...] = ()
    date_ratios: tuple[tuple[date, Decimal], ...] = ()


def _assignment_context(mode, leave_enabled=False):
    if isinstance(mode, ShiftSettings):
        return mode.mode, mode.leave_enabled
    return mode, leave_enabled


def assignment_labels(mode=DAY_ONLY, leave_enabled=False) -> tuple[str, ...]:
    mode, leave_enabled = _assignment_context(mode, leave_enabled)
    return (*SHIFT_MODES.get(mode, ()), *((PUBLIC_OFF, PAID_LEAVE, RECOVERY) if leave_enabled else (OFF,)))


def assignment_label(value, mode=DAY_ONLY, leave_enabled=False) -> str | None:
    """高度モードは公休・有給・明けを区別し、曖昧なboolを受け付けない。"""
    mode, leave_enabled = _assignment_context(mode, leave_enabled)
    if type(value) is bool:
        return ("日勤" if value else OFF) if mode == DAY_ONLY and not leave_enabled else None
    return value if isinstance(value, str) and value in assignment_labels(mode, leave_enabled) else None


def is_working(value, mode=DAY_ONLY, leave_enabled=False) -> bool:
    """配置・勤務開始回数に数える勤務区分。明けを含めない。"""
    mode, leave_enabled = _assignment_context(mode, leave_enabled)
    return assignment_label(value, mode, leave_enabled) in SHIFT_MODES.get(mode, ())


def is_actual_working(value, mode=DAY_ONLY, leave_enabled=False) -> bool:
    """実労働のある暦日。高度モードの明けも連勤に数える。"""
    mode, leave_enabled = _assignment_context(mode, leave_enabled)
    label = assignment_label(value, mode, leave_enabled)
    return label in SHIFT_MODES.get(mode, ()) or (leave_enabled and label == RECOVERY)


@dataclass(frozen=True)
class Source:
    file: str = ""
    sheet: str = ""
    row: int = 0
    column: str = ""


@dataclass(frozen=True)
class Employee:
    id: str
    name: str
    work_days: int | None
    source: Source = Source()
    previous_work_days: int = 0
    previous_shift: str | None = None
    public_holidays: int | None = None
    public_holiday_adjustment: int = 0  # 元有給申請を公休へ変更した内部派生件数。


@dataclass(frozen=True)
class DayRequirement:
    day: date
    required: int
    source: Source = Source()
    shift_counts: tuple[tuple[str, int], ...] = ()


@dataclass(frozen=True)
class FixedOff:
    employee_id: str
    day: date
    source: Source = Source()


@dataclass(frozen=True)
class PaidLeave:
    employee_id: str
    day: date
    source: Source = Source()
    request_id: str = ""  # 空は元の有給一覧、非空は担当者確認した横形式申請。


@dataclass(frozen=True)
class LeaveCandidate:
    request_id: str
    employee_id: str
    day: date
    priority: int
    source: Source = Source()


@dataclass(frozen=True)
class IndependentLeaveRequest:
    request_id: str
    employee_id: str
    day: date
    leave_type: str  # 公休希望 / 有給申請。1レコードは独立した1日分。
    source: Source = Source()
    original_leave_type: str = ""


@dataclass(frozen=True)
class InputData:
    employees: tuple[Employee, ...]
    requirements: tuple[DayRequirement, ...]
    fixed_off: tuple[FixedOff, ...] = ()
    requests: tuple[LeaveCandidate, ...] = ()
    settings: ShiftSettings = ShiftSettings()
    paid_leaves: tuple[PaidLeave, ...] = ()
    independent_requests: tuple[IndependentLeaveRequest, ...] = ()
    confirmed_paid_request_ids: frozenset[str] = frozenset()
    request_format: str = "alternatives"
    converted_public_request_ids: frozenset[str] = frozenset()
    @property
    def days(self) -> tuple[date, ...]:
        return tuple(sorted(r.day for r in self.requirements))


@dataclass(frozen=True)
class Issue:
    kind: str
    message: str
    source: Source = Source()
    employee_id: str = ""
    day: date | None = None


class InputError(ValueError):
    def __init__(self, issues: list[Issue]):
        self.issues = issues
        super().__init__("\n".join(i.message for i in issues))


class SolverStatus(str, Enum):
    OPTIMAL = "最適性確認済み"
    FEASIBLE = "実行可能"
    INFEASIBLE = "条件不成立"
    UNKNOWN = "時間内に案未発見"
    MODEL_ERROR = "モデルエラー"


Schedule = dict[tuple[str, date], bool | str]


@dataclass(frozen=True)
class RequestResult:
    request_id: str
    employee_id: str
    adopted_day: date | None
    priority: int | None
    status: str


@dataclass
class Evaluation:
    issues: list[Issue]
    request_results: list[RequestResult]
    daily_counts: dict[date, int]
    employee_counts: dict[str, int]
    daily_shift_counts: dict[date, dict[str, int]] = field(default_factory=dict)
    employee_shift_counts: dict[str, dict[str, int]] = field(default_factory=dict)
    employee_public_counts: dict[str, int] = field(default_factory=dict)
    employee_paid_counts: dict[str, int] = field(default_factory=dict)
    employee_recovery_counts: dict[str, int] = field(default_factory=dict)
    employee_actual_work_days: dict[str, int] = field(default_factory=dict)
    paid_request_results: list[RequestResult] = field(default_factory=list)
    @property
    def valid(self) -> bool:
        return not self.issues
    @property
    def adopted_count(self) -> int:
        return sum(r.adopted_day is not None for r in self.request_results)
    @property
    def priority_sum(self) -> int:
        return sum(r.priority for r in self.request_results if r.priority is not None)


@dataclass
class SolveResult:
    status: SolverStatus
    schedule: Schedule | None = None
    count_optimal: bool = False
    priority_optimal: bool = False
    elapsed_seconds: float = 0.0
    message: str = ""
    issues: list[Issue] = field(default_factory=list)
