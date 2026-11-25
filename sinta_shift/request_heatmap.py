"""希望集中の読み取り専用集計。候補行数・申請数・希望職員人数を区別する。

日別の希望人数は公休/有給を希望する異なる職員数。同一職員同日でも申請IDの
異なる詳細行は保持する。確定休と確定有給（元＋確認由来）は別の人数で表示する。
確認由来の有給希望は独立メタから一度だけ数え、固定有給一覧から希望人数へ足さない。
必要人数は渡されたInputDataの値を表示するため、GUIは設定適用後データを渡す。
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date
from collections import defaultdict

from .models import (
    InputData, Evaluation, Schedule, Source, PUBLIC_REQUEST, PAID_REQUEST,
    PUBLIC_OFF, PAID_LEAVE, OFF, assignment_label, is_working,
)
from .evaluation import evaluate_schedule
from .planning import attendance_ratio_for_day, required_from_ratio, allocate_shifts

UNCALCULATED = "未計算"
ADOPTED = "採用"
UNADOPTED = "未採用"
OTHER_CANDIDATE = "他候補で採用"
UNKNOWN_RESULT = "結果未確認"
PENDING = "確認待ち"


@dataclass(frozen=True)
class HeatmapEntry:
    request_id: str
    employee_id: str
    employee_name: str
    day: date
    kind: str
    original_kind: str
    priority: int | None
    source: Source
    adopted_day: date | None = None
    status: str = UNCALCULATED
    approval_status: str = ""
    is_alternative: bool = False

    @property
    def id(self):
        return self.request_id


@dataclass(frozen=True)
class DayRequestSummary:
    day: date
    request_employee_count: int
    public_request_employee_count: int
    paid_request_employee_count: int
    adopted_employee_count: int | None
    unadopted_request_count: int | None
    other_candidate_adopted_count: int | None
    fixed_off_employee_count: int
    fixed_paid_employee_count: int
    required_count: int | None
    current_staff_count: int | None
    entries: tuple[HeatmapEntry, ...]
    is_reference: bool = False
    warnings: tuple[str, ...] = ()
    in_period: bool = True

    @property
    def candidate_count(self):
        return len(self.entries)

    @property
    def request_count(self):
        return len({entry.request_id for entry in self.entries})


def build_request_heatmap(data: InputData, schedule: Schedule | None = None,
                          evaluation: Evaluation | None = None, *, raw_fallback: bool = False
                          ) -> tuple[DayRequestSummary, ...]:
    """0希望日を含め日付順で返す。案なしでは採否を推測せずNoneとする。

    raw入力の構文・業務検証はここでは行わない。期間外希望もin_period=Falseの
    行として残す。提示された評価が欠損/矛盾する場合は未採用と断定しない。
    明示確認した有給も、実案が有給でなければ採用扱いにしない。
    """
    names = {employee.id: employee.name for employee in data.employees}
    employees = tuple(names)
    requirements = {row.day: row for row in data.requirements if type(row.day) is date}
    period = set(requirements)
    rows = {}
    day_warnings = defaultdict(list)
    global_warnings = []

    def add(request_id, employee_id, day, kind, original_kind, priority, source, approval="", alternative=False):
        if type(day) is not date:
            global_warnings.append(f"申請{request_id}の日付が不正なため日別集計できません。入力エラーを確認してください。")
            return
        key = (request_id, employee_id, day)
        entry = HeatmapEntry(request_id, employee_id, names.get(employee_id, "氏名不明"), day,
                             kind, original_kind, priority, source, approval_status=approval,
                             is_alternative=alternative)
        if key in rows:
            old = rows[key]
            if (old.kind, old.priority) != (entry.kind, entry.priority):
                day_warnings[day].append(f"申請{request_id}の同日データが重複または矛盾しています。")
            return
        rows[key] = entry
        if employee_id not in names:
            day_warnings[day].append(f"未知の職員IDです: {employee_id}")

    # 独立メタを優先し、同じ公休候補がrequestsにもあることを二重計上しない。
    for request in data.independent_requests:
        original = request.original_leave_type or request.leave_type
        approval = ("公休へ変更" if original == PAID_REQUEST and request.leave_type == PUBLIC_REQUEST else
                    "有給確定" if request.leave_type == PAID_REQUEST and request.request_id in data.confirmed_paid_request_ids else
                    PENDING if request.leave_type == PAID_REQUEST else "")
        add(request.request_id, request.employee_id, request.day, request.leave_type, original,
            1 if request.leave_type == PUBLIC_REQUEST else None, request.source, approval)
    for request in data.requests:
        key = (request.request_id, request.employee_id, request.day)
        if key in rows:
            if rows[key].kind != PUBLIC_REQUEST:
                day_warnings[request.day].append(f"有給申請{request.request_id}が公休候補にも含まれています。")
            continue
        add(request.request_id, request.employee_id, request.day, PUBLIC_REQUEST, PUBLIC_REQUEST,
            request.priority, request.source, alternative=True)

    by_day = defaultdict(list)
    by_id = defaultdict(list)
    for entry in rows.values():
        by_day[entry.day].append(entry)
        by_id[entry.request_id].append(entry)
    fixed_off = defaultdict(set)
    fixed_paid = defaultdict(set)
    for item in data.fixed_off:
        if type(item.day) is date:
            fixed_off[item.day].add(item.employee_id)
    for item in data.paid_leaves:
        if type(item.day) is date:
            fixed_paid[item.day].add(item.employee_id)
    # 明示確認済みIDは固定有給へ展開する前のデータでも同じ人数になる。
    for request in data.independent_requests:
        if (request.leave_type == PAID_REQUEST and request.request_id in data.confirmed_paid_request_ids
                and type(request.day) is date):
            fixed_paid[request.day].add(request.employee_id)

    # GUIは通常freshな評価を渡すが、人数が同じでも職員間の入替・夜勤順序等で
    # 古い評価は有効でなくなる。完成案との誤認を防ぐため必須条件を再確認する。
    current_evaluation = evaluate_schedule(data, schedule) if schedule is not None else None
    evaluated = current_evaluation if evaluation is None else evaluation
    results = defaultdict(list)
    if schedule is not None and evaluated is not None:
        if evaluation is not None and current_evaluation is not None:
            if (evaluation.daily_counts != current_evaluation.daily_counts
                    or evaluation.employee_counts != current_evaluation.employee_counts
                    or evaluation.daily_shift_counts != current_evaluation.daily_shift_counts
                    or evaluation.employee_actual_work_days != current_evaluation.employee_actual_work_days
                    or evaluation.request_results != current_evaluation.request_results
                    or evaluation.paid_request_results != current_evaluation.paid_request_results
                    or bool(evaluation.issues) != bool(current_evaluation.issues)):
                global_warnings.append("渡された評価が現在の案と一致しません。採否は参考として再評価が必要です。")
        if current_evaluation is not None and current_evaluation.issues:
            global_warnings.append(f"現在の案に必須条件違反または入力不備{len(current_evaluation.issues)}件があります。採否は参考で、完成案ではありません。")
        for result in evaluated.request_results:
            results[result.request_id].append(result)
        unknown = {result.request_id for result in (*evaluated.request_results, *evaluated.paid_request_results)} - set(by_id)
        if unknown:
            global_warnings.append("評価に未知の申請IDがあります: " + " / ".join(sorted(unknown)))
    if raw_fallback:
        global_warnings.append("設定適用前の入力を表示しています。必要人数は元データの参考値です。")

    def outcome(entry):
        if schedule is None:
            return entry
        if entry.kind == PAID_REQUEST:
            if entry.request_id not in data.confirmed_paid_request_ids:
                day_warnings[entry.day].append(f"有給申請{entry.request_id}は確認待ちです。実案の採否は参考です。")
                return replace(entry, status=PENDING)
            actual = assignment_label(schedule.get((entry.employee_id, entry.day)), data.settings)
            if actual == PAID_LEAVE:
                return replace(entry, status=ADOPTED, adopted_day=entry.day)
            day_warnings[entry.day].append(f"確定有給{entry.request_id}が実案に配置されていません。採否は参考です。")
            return replace(entry, status=UNADOPTED)
        if entry.kind != PUBLIC_REQUEST:
            day_warnings[entry.day].append(f"申請{entry.request_id}の区分が不明です。")
            return replace(entry, status=UNKNOWN_RESULT)
        matched = results.get(entry.request_id, ())
        if len(matched) != 1 or matched[0].employee_id != entry.employee_id:
            day_warnings[entry.day].append(f"申請{entry.request_id}の採否結果が欠落・重複または職員ID不一致です。")
            return replace(entry, status=UNKNOWN_RESULT)
        result = matched[0]
        if result.adopted_day is None:
            if result.status != UNADOPTED:
                day_warnings[entry.day].append(f"申請{entry.request_id}の採否状態が不明です。")
                return replace(entry, status=UNKNOWN_RESULT)
            return replace(entry, status=UNADOPTED)
        if result.status not in (ADOPTED, f"第{result.priority}希望を採用"):
            day_warnings[entry.day].append(f"申請{entry.request_id}の採否状態が不明です。")
            return replace(entry, status=UNKNOWN_RESULT)
        candidates = [item for item in by_id[entry.request_id] if item.employee_id == entry.employee_id and item.day == result.adopted_day]
        expected_rest = PUBLIC_OFF if data.settings.leave_enabled else OFF
        if (len(candidates) != 1 or candidates[0].priority != result.priority
                or assignment_label(schedule.get((entry.employee_id, result.adopted_day)), data.settings) != expected_rest):
            day_warnings[entry.day].append(f"申請{entry.request_id}の採用日・順位と実案が一致しません。")
            return replace(entry, status=UNKNOWN_RESULT)
        return replace(entry, status=ADOPTED if entry.day == result.adopted_day else OTHER_CANDIDATE,
                             adopted_day=result.adopted_day)

    transformed = {day: tuple(outcome(entry) for entry in entries) for day, entries in by_day.items()}
    summaries = []
    # 期間外の希望・固定日も隠さず、必要人数なしの参考行として返す。
    days = sorted(period | set(by_day) | set(fixed_off) | set(fixed_paid))
    for day in days:
        entries = tuple(sorted(transformed.get(day, ()), key=lambda e: (e.employee_name, e.employee_id, e.request_id, e.priority or 0)))
        warnings = [*global_warnings, *day_warnings[day]]
        requirement = requirements.get(day)
        required = requirement.required if requirement is not None else None
        if day not in period:
            warnings.append("対象期間外の日付です。入力エラーを確認してください。")
        if requirement is not None:
            try:
                ratio = attendance_ratio_for_day(day, data.settings)
                if ratio is not None:
                    expected = required_from_ratio(len(employees), ratio)
                    expected_shifts = dict(allocate_shifts(expected, data.settings.mode, data.settings.shift_weights))
                    if required != expected or (requirement.shift_counts and dict(requirement.shift_counts) != expected_shifts):
                        warnings.append("必要人数に出勤率の設定が未反映です。元入力の参考人数を表示しています。")
            except ValueError as exc:
                warnings.append(f"出勤率を適用できない入力の参考値です: {exc}")
        current = None
        if schedule is not None and day in period:
            labels = [assignment_label(schedule.get((employee_id, day)), data.settings) for employee_id in employees]
            if all(label is not None for label in labels):
                current = sum(is_working(schedule[(employee_id, day)], data.settings) for employee_id in employees)
            else:
                warnings.append("実案に未入力または不正な区分があります。現在配置人数は未確認です。")
            if evaluated is None or day not in evaluated.daily_counts:
                warnings.append("日別人数の再評価結果がありません。")
            elif current is not None and current != evaluated.daily_counts[day]:
                warnings.append("実案と再評価済み人数が一致しません。再評価してください。")
        pending = any(entry.status in (UNCALCULATED, UNKNOWN_RESULT, PENDING) for entry in entries)
        known_outcomes = schedule is not None and not pending
        summaries.append(DayRequestSummary(
            day=day,
            request_employee_count=len({entry.employee_id for entry in entries}),
            public_request_employee_count=len({entry.employee_id for entry in entries if entry.kind == PUBLIC_REQUEST}),
            paid_request_employee_count=len({entry.employee_id for entry in entries if entry.kind == PAID_REQUEST}),
            adopted_employee_count=len({entry.employee_id for entry in entries if entry.status == ADOPTED}) if known_outcomes else None,
            unadopted_request_count=len({entry.request_id for entry in entries if entry.status == UNADOPTED}) if known_outcomes else None,
            other_candidate_adopted_count=len({entry.request_id for entry in entries if entry.status == OTHER_CANDIDATE}) if known_outcomes else None,
            fixed_off_employee_count=len(fixed_off[day]), fixed_paid_employee_count=len(fixed_paid[day]),
            required_count=required, current_staff_count=current, entries=entries,
            is_reference=bool(warnings), warnings=tuple(dict.fromkeys(warnings)), in_period=day in period,
        ))
    return tuple(summaries)
