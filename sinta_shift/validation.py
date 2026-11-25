"""入力検証。日数・人数は整数であり、欠損値は補完しない。"""
from collections import Counter, defaultdict
from datetime import date
from .models import DAY_ONLY, OFF, PUBLIC_OFF, PAID_LEAVE, RECOVERY, PUBLIC_DAYS, PUBLIC_REQUEST, PAID_REQUEST, SHIFT_MODES, InputData, InputError, Issue, Source
from .planning import configure_input, crossing_shifts, previous_shift_label, public_holiday_target


def _integer(value) -> bool:
    return isinstance(value, int) and not isinstance(value, bool)


def validate_input(data: InputData) -> list[Issue]:
    try:
        data = configure_input(data)
    except InputError as exc:
        return exc.issues
    issues: list[Issue] = []
    advanced = data.settings.leave_enabled
    crosses = crossing_shifts(data.settings) if advanced else frozenset()
    records = (data.employees, data.requirements, data.fixed_off, data.requests, data.paid_leaves, data.independent_requests)
    input_file = next((record.source.file for rows in records for record in rows if record.source.file), "")
    def table_source(label, rows=()):
        first = rows[0].source if rows else Source()
        return Source(input_file, first.sheet or label)
    def add(message, source=Source(), employee_id="", day=None, kind="入力不備"):
        issues.append(Issue(kind, message, source, employee_id, day))

    if not data.employees:
        add("職員が0件です。1人以上記入してください。", table_source("職員"))
    if not data.requirements:
        add("必要人数が0件です。対象期間の全日を連続して記入してください。", table_source("必要人数", data.requirements))
    ids = set()
    for employee in data.employees:
        if not isinstance(employee.id, str) or not employee.id.strip():
            add("職員IDを文字列で記入してください。", employee.source)
        elif employee.id in ids:
            add(f"職員ID「{employee.id}」が重複しています。", employee.source, employee.id)
        else:
            ids.add(employee.id)
        if not isinstance(employee.name, str) or not employee.name.strip():
            add("氏名が空欄です。", employee.source, employee.id)
        if not _integer(employee.previous_work_days) or employee.previous_work_days < 0:
            add("期間開始前の連続勤務日数は0以上の整数にしてください。", employee.source, employee.id)
        elif data.settings.max_consecutive_days is not None and employee.previous_work_days > data.settings.max_consecutive_days:
            add(f"期間開始前の連続勤務{employee.previous_work_days}日は上限{data.settings.max_consecutive_days}日を超えています。前期間の記録または上限を確認してください。", employee.source, employee.id)
        working_labels = {label for group in SHIFT_MODES.values() for label in group}
        nonworking_history = {OFF, PUBLIC_OFF, PAID_LEAVE} if advanced else {OFF}
        actual_history = working_labels | ({RECOVERY} if advanced else set())
        known_previous_labels = actual_history | nonworking_history
        if employee.previous_shift is not None and (not isinstance(employee.previous_shift, str) or employee.previous_shift not in known_previous_labels):
            add("期間開始前の勤務区分が不正です。有効な勤務区分・休み区分から指定してください。", employee.source, employee.id)
        elif _integer(employee.previous_work_days) and employee.previous_work_days >= 0:
            previous = previous_shift_label(employee, data.settings)
            if employee.previous_work_days == 0 and previous in actual_history:
                add("期間開始前の連勤0日と勤務区分が矛盾しています。0日は前日に実労働なしとして扱います。", employee.source, employee.id)
            elif employee.previous_work_days > 0 and previous in nonworking_history:
                add("期間開始前の連勤が1日以上なのに前日区分が休みです。履歴を確認してください。", employee.source, employee.id)
            elif employee.previous_work_days > 0 and previous is None:
                add("交代制で期間開始前の連勤が1日以上ある場合、前日の勤務区分を指定してください。", employee.source, employee.id)
            if advanced and previous in working_labels:
                time_labels = {shift.label for shift in data.settings.shift_times}
                if previous not in time_labels and not (data.settings.mode == DAY_ONLY and previous == "日勤"):
                    add(f"期間開始前の{previous}の勤務時刻が未設定です。初日の明けを判定するため時刻を指定してください。", employee.source, employee.id)
    days = []
    for requirement in data.requirements:
        if type(requirement.day) is not date:
            add("必要人数の日付は年を含む日付にしてください。", requirement.source)
        else:
            if requirement.day in days:
                add(f"日付{requirement.day}が重複しています。", requirement.source, day=requirement.day)
            days.append(requirement.day)
        if not _integer(requirement.required) or not 0 <= requirement.required <= len(data.employees):
            add(f"配置人数は0〜職員数（{len(data.employees)}）の整数にしてください。", requirement.source, day=requirement.day)
    unique_days = sorted(set(days))
    if unique_days and (unique_days[-1] - unique_days[0]).days + 1 != len(unique_days):
        add("必要人数の日付が連続していません。対象期間の全日を記入してください。", table_source("必要人数", data.requirements))
    day_set = set(unique_days)
    for employee in data.employees:
        if not (advanced and data.settings.day_count_mode == PUBLIC_DAYS and employee.work_days is None):
            if not _integer(employee.work_days) or not 0 <= employee.work_days <= len(unique_days):
                add(f"対象期間の勤務日数（勤務開始回数）は0〜{len(unique_days)}の整数にしてください。", employee.source, employee.id)
        target = public_holiday_target(employee, data.settings)
        if target is not None and not 0 <= target <= len(unique_days):
            add(f"対象期間の公休日数は0〜{len(unique_days)}の整数にしてください。", employee.source, employee.id)
    fixed_seen = set()
    for fixed in data.fixed_off:
        if fixed.employee_id not in ids:
            add(f"確定休に未知の職員ID「{fixed.employee_id}」があります。", fixed.source, fixed.employee_id, fixed.day)
        if type(fixed.day) is not date or fixed.day not in day_set:
            add("確定休の日付が対象期間外です。", fixed.source, fixed.employee_id, fixed.day)
        key = (fixed.employee_id, fixed.day)
        if key in fixed_seen:
            add("同じ職員・日付の確定休が重複しています。", fixed.source, fixed.employee_id, fixed.day)
        fixed_seen.add(key)
    paid_seen = set()
    for paid in data.paid_leaves:
        if paid.employee_id not in ids:
            add(f"有給に未知の職員ID「{paid.employee_id}」があります。", paid.source, paid.employee_id, paid.day)
        if type(paid.day) is not date or paid.day not in day_set:
            add("有給の日付が対象期間外です。", paid.source, paid.employee_id, paid.day)
        key = (paid.employee_id, paid.day)
        if key in paid_seen:
            add("同じ職員・日付の有給が重複しています。", paid.source, paid.employee_id, paid.day)
        if key in fixed_seen:
            add("同じ職員・日付に確定公休と有給が重複しています。", paid.source, paid.employee_id, paid.day)
        paid_seen.add(key)
    request_ids: dict[str, str] = {}
    candidates_seen = set()
    for candidate in data.requests:
        if not isinstance(candidate.request_id, str) or not candidate.request_id.strip():
            add("申請IDが空欄です。", candidate.source, candidate.employee_id, candidate.day)
        if candidate.request_id in request_ids and request_ids[candidate.request_id] != candidate.employee_id:
            add(f"申請ID「{candidate.request_id}」に複数の職員が混在しています。", candidate.source, candidate.employee_id, candidate.day)
        request_ids[candidate.request_id] = candidate.employee_id
        if candidate.employee_id not in ids:
            add(f"希望休に未知の職員ID「{candidate.employee_id}」があります。", candidate.source, candidate.employee_id, candidate.day)
        if type(candidate.day) is not date or candidate.day not in day_set:
            add("希望休の候補日が対象期間外です。", candidate.source, candidate.employee_id, candidate.day)
        key = (candidate.request_id, candidate.day)
        if key in candidates_seen:
            add(f"申請「{candidate.request_id}」の候補日が重複しています。", candidate.source, candidate.employee_id, candidate.day)
        candidates_seen.add(key)
        if not _integer(candidate.priority) or not 1 <= candidate.priority <= 1_000_000_000:
            add("優先順位は1〜1000000000の整数にしてください。", candidate.source, candidate.employee_id, candidate.day)
    independent_ids = set()
    independent_cells = set()
    base_paid_cells = {(paid.employee_id, paid.day) for paid in data.paid_leaves if not paid.request_id}
    candidates_by_id = defaultdict(list)
    for candidate in data.requests:
        candidates_by_id[candidate.request_id].append(candidate)
    for request in data.independent_requests:
        valid_request_id = isinstance(request.request_id, str) and bool(request.request_id.strip())
        if not valid_request_id:
            add("独立した希望の申請IDは空欄でない文字列にしてください。", request.source, request.employee_id, request.day)
        elif request.request_id in independent_ids:
            add(f"独立申請ID「{request.request_id}」が重複しています。", request.source, request.employee_id, request.day)
        else:
            independent_ids.add(request.request_id)
        if request.employee_id not in ids:
            add(f"独立した希望に未知の職員ID「{request.employee_id}」があります。", request.source, request.employee_id, request.day)
        if type(request.day) is not date or request.day not in day_set:
            add("独立した希望の日付が対象期間外です。", request.source, request.employee_id, request.day)
        key = (request.employee_id, request.day)
        if key in independent_cells:
            add("同じ職員・同じ日の独立した希望が重複しています。公休希望と有給申請も同日に重ねられません。", request.source, request.employee_id, request.day)
        independent_cells.add(key)
        if request.leave_type not in (PUBLIC_REQUEST, PAID_REQUEST):
            add("独立した希望の区分は公休希望・有給申請から選んでください。", request.source, request.employee_id, request.day)
            continue
        candidates = candidates_by_id.get(request.request_id, ()) if valid_request_id else ()
        if request.leave_type == PUBLIC_REQUEST:
            if key in paid_seen:
                add("同じ職員・同じ日に独立した公休希望と固定有給が重複しています。", request.source, request.employee_id, request.day)
            if valid_request_id and (len(candidates) != 1 or candidates[0].employee_id != request.employee_id or candidates[0].day != request.day or candidates[0].priority != 1):
                add("独立した公休希望は1申請につき入力日の1候補・同順位1で対応させてください。代替候補として結合できません。", request.source, request.employee_id, request.day)
        else:
            if key in base_paid_cells:
                add("同じ職員・同じ日に有給申請と元の有給一覧が重複しています。どちらか一方に記入してください。", request.source, request.employee_id, request.day)
            if key in fixed_seen:
                add("同じ職員・同じ日に有給申請と確定公休が重複しています。", request.source, request.employee_id, request.day)
            if candidates:
                add("有給申請を未採用にできる希望休候補へ混ぜられません。担当者確認後の固定有給として扱います。", request.source, request.employee_id, request.day)
            if not valid_request_id or request.request_id not in data.confirmed_paid_request_ids:
                add("有給申請は担当者の明示確認が必要です。確認するまで計算・完成版保存を停止します。", request.source, request.employee_id, request.day, "有給申請未確認")
    # 型・参照不備のある表から算術的な結論を出さない。
    if issues:
        return issues
    required_total = sum(r.required for r in data.requirements)
    first_recovery = {e.id for e in data.employees if advanced and previous_shift_label(e, data.settings) in crosses}
    recovery_by_day = Counter()
    if advanced and unique_days:
        recovery_by_day[unique_days[0]] = len(first_recovery)
        by_day = {requirement.day: requirement for requirement in data.requirements}
        for index in range(1, len(unique_days)):
            recovery_by_day[unique_days[index]] = sum(count for label, count in by_day[unique_days[index - 1]].shift_counts if label in crosses)
    if advanced and data.settings.day_count_mode == PUBLIC_DAYS:
        public_total = sum(public_holiday_target(e, data.settings) for e in data.employees)
        recovery_total = sum(recovery_by_day.values())
        paid_total = len(data.paid_leaves)
        capacity = len(data.employees) * len(unique_days)
        if required_total + recovery_total + public_total + paid_total != capacity:
            add(f"日数が一致しません：勤務開始{required_total}日＋明け{recovery_total}日＋公休{public_total}日＋有給{paid_total}日＝{required_total + recovery_total + public_total + paid_total}日、全員の対象暦日数は{capacity}日です。個人勤務日数を自動調整して隠しません。", table_source("職員／必要人数／有給"))
    else:
        employee_total = sum(e.work_days for e in data.employees)
        if employee_total != required_total:
            add(f"勤務日数合計{employee_total}人日と配置人数合計{required_total}人日が一致しません。", table_source("職員／必要人数"))
    fixed_by_day = Counter(f.day for f in data.fixed_off)
    fixed_by_employee = Counter(f.employee_id for f in data.fixed_off)
    paid_by_day = Counter(p.day for p in data.paid_leaves)
    paid_by_employee = Counter(p.employee_id for p in data.paid_leaves)
    for requirement in data.requirements:
        available_raw = len(data.employees) - fixed_by_day[requirement.day] - paid_by_day[requirement.day] - recovery_by_day[requirement.day]
        if available_raw < 0:
            add(f"{requirement.day}：確定公休・有給・必要な明けの合計が職員数を超えています。同じ日の区分を重ねられません。", requirement.source, day=requirement.day, kind="条件不成立")
        available = max(0, available_raw)
        if available < requirement.required:
            add(f"{requirement.day}：必要{requirement.required}人・勤務可能{available}人。確定休・有給・明けまたは配置人数を見直してください。",
                requirement.source, day=requirement.day, kind="条件不成立")
    for employee in data.employees:
        available = len(unique_days) - fixed_by_employee[employee.id] - paid_by_employee[employee.id] - int(employee.id in first_recovery)
        if employee.work_days is not None and employee.work_days > available:
            add(f"職員{employee.id}：勤務日数{employee.work_days}日・勤務可能{available}日。確定休・有給・明けと勤務日数が両立しません。",
                employee.source, employee.id, kind="条件不成立")
        target = public_holiday_target(employee, data.settings)
        if target is not None and fixed_by_employee[employee.id] > target:
            add(f"職員{employee.id}：確定公休{fixed_by_employee[employee.id]}日が公休目標{target}日を超えています。", employee.source, employee.id)
        if unique_days and employee.id in first_recovery:
            first = unique_days[0]
            if (employee.id, first) in fixed_seen or (employee.id, first) in paid_seen:
                add("前日の勤務が日跨ぎのため初日は明けが必要です。公休・有給と同じ日に重ねられません。", employee.source, employee.id, first, "条件不成立")
    if sum(c.priority for c in data.requests) >= 2 ** 60:
        add("優先順位の合計が計算可能な範囲を超えています。", table_source("希望休", data.requests))
    return issues
