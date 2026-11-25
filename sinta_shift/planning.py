"""曜日比率と勤務区分配分から、整数の必要人数を組み立てる。"""
from dataclasses import replace
from collections import Counter, defaultdict, deque
from datetime import date, timedelta
from decimal import Decimal

from .models import InputData, InputError, Issue, ShiftSettings, SHIFT_MODES, DAY_ONLY, TWO_SHIFT, THREE_SHIFT, OFF, PUBLIC_OFF, PAID_LEAVE, RECOVERY, DIRECT_DAYS, PUBLIC_DAYS, PUBLIC_REQUEST, PAID_REQUEST, PaidLeave, LeaveCandidate, ShiftTime, Source, assignment_labels


def _integer(value):
    return isinstance(value, int) and not isinstance(value, bool)


def _error(message, source=Source(sheet="勤務条件")):
    raise InputError([Issue("勤務条件", message, source)])


def _percentage(value):
    if _integer(value):
        value = Decimal(value)
    if not isinstance(value, Decimal) or not value.is_finite() or not Decimal(0) <= value <= Decimal(100):
        _error("曜日別の配置比率は0〜100％の有限の数値で指定してください。")
    return value


def required_from_ratio(n: int, percentage: Decimal) -> int:
    """全職員数×比率を切り上げる。勤務日数を自動変更しない。"""
    if not _integer(n) or n < 0:
        _error("職員数は0以上の整数にしてください。")
    percentage = _percentage(percentage)
    numerator, denominator = percentage.as_integer_ratio()
    denominator *= 100
    return (n * numerator + denominator - 1) // denominator


def allocate_shifts(total: int, mode: str, weights=()) -> tuple[tuple[str, int], ...]:
    """最大剰余法で整数配分。同率の余りは区分の表示順へ割り当てる。"""
    mode = {"2交代制": TWO_SHIFT, "3交代制": THREE_SHIFT}.get(mode, mode) if isinstance(mode, str) else mode
    if not isinstance(mode, str) or mode not in SHIFT_MODES:
        _error("勤務区分は日勤のみ・２交代制・３交代制から選んでください。")
    if not _integer(total) or total < 0:
        _error("配置人数は0以上の整数にしてください。")
    labels = SHIFT_MODES[mode]
    if not isinstance(weights, (tuple, list)):
        _error("勤務区分の配分は整数の組で指定してください。")
    actual = tuple(weights) if weights else (1,) * len(labels)
    if len(actual) != len(labels) or any(not _integer(w) or w < 0 for w in actual) or sum(actual) <= 0:
        _error("配分は各勤務区分に0以上の整数を1つずつ指定し、合計を1以上にしてください。")
    weight_sum = sum(actual)
    quotients = [total * weight // weight_sum for weight in actual]
    remainder_order = sorted(range(len(labels)), key=lambda i: (-(total * actual[i] % weight_sum), i))
    for index in remainder_order[:total - sum(quotients)]:
        quotients[index] += 1
    return tuple(zip(labels, quotients))


def _choice_ids(values, label, source):
    if not isinstance(values, (set, frozenset, tuple, list)) or any(not isinstance(value, str) or not value.strip() for value in values):
        raise InputError([Issue("有給確認設定", f"{label}の申請IDは空欄でない文字列の集合にしてください。", source)])
    return frozenset(values)


def original_request_type(request):
    return request.original_leave_type or request.leave_type


def _apply_leave_choices(data: InputData, paid_ids, public_ids, *, strict: bool) -> InputData:
    """決定値から派生データを再生成。旧確認APIのエラー保持互換を分離する。"""
    default_source = next((request.source for request in data.independent_requests), Source(sheet="希望休"))
    paid_ids = _choice_ids(paid_ids, "有給確定", default_source)
    public_ids = _choice_ids(public_ids, "公休へ変更", default_source)
    if paid_ids & public_ids:
        raise InputError([Issue("有給確認設定", "同じ申請を有給確定と公休変更の両方に指定できません。", default_source)])
    requests = {}
    issues = []
    for request in data.independent_requests:
        if not isinstance(request.request_id, str) or not request.request_id.strip():
            issues.append(Issue("入力不備", "独立した希望の申請IDは空欄でない文字列にしてください。", request.source, request.employee_id, request.day))
            continue
        if request.request_id in requests:
            issues.append(Issue("入力不備", f"独立申請ID「{request.request_id}」が重複しています。", request.source, request.employee_id, request.day))
        requests[request.request_id] = request
        if request.leave_type not in (PUBLIC_REQUEST, PAID_REQUEST) or request.original_leave_type not in ("", PUBLIC_REQUEST, PAID_REQUEST):
            issues.append(Issue("入力不備", "独立した希望の元区分・現在区分は公休希望・有給申請にしてください。", request.source, request.employee_id, request.day))
        elif original_request_type(request) == PUBLIC_REQUEST and request.leave_type != PUBLIC_REQUEST:
            issues.append(Issue("有給確認設定", "元の公休希望を有給申請へ勝手に変更できません。", request.source, request.employee_id, request.day))
    for request_id in sorted(paid_ids | public_ids):
        request = requests.get(request_id)
        if request is None:
            issues.append(Issue("有給確認設定", f"未知の有給申請ID「{request_id}」が配分設定にあります。", default_source))
        elif original_request_type(request) != PAID_REQUEST:
            issues.append(Issue("有給確認設定", f"申請「{request_id}」は元の有給申請ではないため配分できません。", request.source, request.employee_id, request.day))
    original_paid_ids = {request_id for request_id, request in requests.items() if original_request_type(request) == PAID_REQUEST}
    old_derived_ids = set(data.converted_public_request_ids) | {
        request.request_id for request in data.independent_requests
        if request.original_leave_type == PAID_REQUEST and request.leave_type == PUBLIC_REQUEST
    }
    for candidate in data.requests:
        if candidate.request_id in original_paid_ids and candidate.request_id not in old_derived_ids:
            issues.append(Issue("入力不備", "有給申請を未採用にできる希望休候補へ混ぜられません。担当者の公休変更操作から派生させてください。", candidate.source, candidate.employee_id, candidate.day))
    if strict:
        seen_cells = set()
        base_paid_cells = {(paid.employee_id, paid.day) for paid in data.paid_leaves if not paid.request_id}
        fixed_cells = {(fixed.employee_id, fixed.day) for fixed in data.fixed_off}
        staff_ids = {employee.id for employee in data.employees}
        period_days = {requirement.day for requirement in data.requirements if type(requirement.day) is date}
        for request in data.independent_requests:
            if request.employee_id not in staff_ids or type(request.day) is not date or request.day not in period_days:
                issues.append(Issue("入力不備", "配分対象の職員ID・日付が不正または対象期間外です。", request.source, request.employee_id, request.day))
            key = (request.employee_id, request.day)
            if key in seen_cells:
                issues.append(Issue("入力不備", "同じ職員・同じ日の独立した希望が重複しています。修正後に配分してください。", request.source, request.employee_id, request.day))
            seen_cells.add(key)
            if key in base_paid_cells:
                issues.append(Issue("入力不備", "独立した希望と元の固定有給が同じ職員・同じ日に重複しています。", request.source, request.employee_id, request.day))
            if original_request_type(request) == PAID_REQUEST and key in fixed_cells:
                issues.append(Issue("入力不備", "元の有給申請と確定公休が同じ職員・同じ日に重複しています。", request.source, request.employee_id, request.day))
    if issues:
        raise InputError(issues)
    metadata = []
    for request in data.independent_requests:
        if original_request_type(request) == PAID_REQUEST:
            converted = request.request_id in public_ids
            original = PAID_REQUEST if converted or request.original_leave_type else ""
            request = replace(request, leave_type=PUBLIC_REQUEST if converted else PAID_REQUEST, original_leave_type=original)
        metadata.append(request)
    base_candidates = tuple(candidate for candidate in data.requests if candidate.request_id not in original_paid_ids)
    derived_candidates = tuple(LeaveCandidate(request.request_id, request.employee_id, request.day, 1, request.source)
                               for request in metadata if request.request_id in public_ids)
    base_paid = tuple(paid for paid in data.paid_leaves if not paid.request_id)
    derived_paid = tuple(PaidLeave(request.employee_id, request.day, request.source, request.request_id)
                         for request in metadata if request.request_id in paid_ids)
    adjustments = Counter(request.employee_id for request in metadata if request.request_id in public_ids)
    employees = tuple(replace(employee, public_holiday_adjustment=adjustments[employee.id]) for employee in data.employees)
    return replace(data, employees=employees, requests=(*base_candidates, *derived_candidates),
                   independent_requests=tuple(metadata), paid_leaves=(*base_paid, *derived_paid),
                   confirmed_paid_request_ids=paid_ids, converted_public_request_ids=public_ids)


def set_leave_choices(data: InputData, paid_ids, public_ids) -> InputData:
    """元有給申請への明示決定。重複や元の固定休暇との衝突は適用前に拒否。"""
    return _apply_leave_choices(data, paid_ids, public_ids, strict=True)


def confirm_paid_requests(data: InputData, confirmed_ids) -> InputData:
    """旧API。指定した有給確定IDを公休変更から戻し、他の公休変更は保持する。"""
    source = next((request.source for request in data.independent_requests), Source(sheet="希望休"))
    confirmed_ids = _choice_ids(confirmed_ids, "有給確定", source)
    public_ids = _choice_ids(data.converted_public_request_ids, "公休へ変更", source)
    # 旧APIは不正な元入力を黙って修復せず保持し、validate_inputで一覧化する。
    return _apply_leave_choices(data, confirmed_ids, public_ids - confirmed_ids, strict=False)


def plan_paid_mix(data: InputData, paid_count: int):
    """未決定分だけの配分案。既存決定・元有給一覧を変更せずID集合を返す。"""
    prepared = set_leave_choices(data, data.confirmed_paid_request_ids, data.converted_public_request_ids)
    source = next((request.source for request in data.independent_requests), Source(sheet="希望休"))
    decided = prepared.confirmed_paid_request_ids | prepared.converted_public_request_ids
    pending = [request for request in prepared.independent_requests
               if original_request_type(request) == PAID_REQUEST and request.request_id not in decided]
    if not _integer(paid_count) or not 0 <= paid_count <= len(pending):
        raise InputError([Issue("有給配分", f"今回の有給件数は未決定件数0〜{len(pending)}の整数にしてください。", source)])
    groups = defaultdict(list)
    staff = {employee.id: employee.name for employee in prepared.employees}
    period_days = {requirement.day for requirement in prepared.requirements if type(requirement.day) is date}
    for request in pending:
        if request.employee_id not in staff or type(request.day) is not date or request.day not in period_days:
            raise InputError([Issue("有給配分", "配分対象の職員ID・日付が不正です。", request.source, request.employee_id, request.day)])
        groups[request.employee_id].append(request)
    queues = {employee_id: deque(sorted(rows, key=lambda r: (r.day, r.request_id))) for employee_id, rows in groups.items()}
    order = sorted(queues, key=lambda employee_id: (staff[employee_id], employee_id))
    round_robin = []
    while any(queues.values()):
        for employee_id in order:
            if queues[employee_id]:
                round_robin.append(queues[employee_id].popleft().request_id)
    selected_paid = frozenset(round_robin[:paid_count])
    selected_public = frozenset(round_robin[paid_count:])
    return (prepared.confirmed_paid_request_ids | selected_paid,
            prepared.converted_public_request_ids | selected_public)


def week_start(day: date) -> date:
    if type(day) is not date:
        _error("週を求める日付は年を含む日付にしてください。")
    return day - timedelta(days=day.weekday())


def _ratio_overrides(values, *, weekly, source=Source(sheet="勤務条件")):
    label = "週指定" if weekly else "日付指定"
    if not isinstance(values, (tuple, list)):
        _error(f"{label}の出勤率は日付と比率の組で指定してください。", source)
    parsed = []
    seen = set()
    for pair in values:
        if not isinstance(pair, (tuple, list)) or len(pair) != 2 or type(pair[0]) is not date:
            _error(f"{label}の出勤率には年を含む日付と比率を指定してください。", source)
        day, value = pair
        if weekly and day.weekday() != 0:
            _error("週指定の開始日は月曜日にしてください。", source)
        if day in seen:
            _error(f"{label}の出勤率の日付が重複しています: {day}", source)
        seen.add(day)
        try:
            value = _percentage(value)
        except InputError as exc:
            raise InputError([replace(issue, source=source) for issue in exc.issues]) from exc
        parsed.append((day, value))
    return tuple(sorted(parsed))


def attendance_ratio_for_day(day: date, settings: ShiftSettings) -> Decimal | None:
    if type(day) is not date:
        _error("出勤率を求める日付は年を含む日付にしてください。")
    dates = dict(_ratio_overrides(settings.date_ratios, weekly=False))
    weeks = dict(_ratio_overrides(settings.weekly_ratios, weekly=True))
    if day in dates:
        return dates[day]
    if week_start(day) in weeks:
        return weeks[week_start(day)]
    return _percentage(settings.weekday_ratios[day.weekday()]) if settings.weekday_ratios is not None else None


def attendance_source_for_day(day: date, settings: ShiftSettings) -> str:
    if type(day) is not date:
        _error("出勤率を求める日付は年を含む日付にしてください。")
    if day in dict(_ratio_overrides(settings.date_ratios, weekly=False)):
        return "日付指定"
    if week_start(day) in dict(_ratio_overrides(settings.weekly_ratios, weekly=True)):
        return "週指定"
    return "曜日指定" if settings.weekday_ratios is not None else "Excel人数"


def configure_input(data: InputData, settings: ShiftSettings | None = None) -> InputData:
    """元の個人勤務日数を保ち、指定条件で必要人数だけを確定する。

    同一条件で繰り返し呼び出しても結果は変わらない。条件を変更する場合は、
    自動生成した配分と明示内訳を混同しないよう読込時の原データを渡す。
    """
    data = confirm_paid_requests(data, data.confirmed_paid_request_ids)
    if data.request_format not in ("alternatives", "independent"):
        source = next((request.source for request in data.independent_requests), Source(sheet="希望休"))
        raise InputError([Issue("希望休形式", "希望休形式はalternatives（代替候補）またはindependent（独立日）にしてください。", source)])
    if data.independent_requests and data.request_format == "alternatives":
        data = replace(data, request_format="independent")
    settings = data.settings if settings is None else settings
    input_file = next((r.source.file for rows in (data.employees, data.requirements, data.fixed_off, data.requests, data.paid_leaves, data.independent_requests) for r in rows if r.source.file), "")
    setting_source = Source(input_file, "勤務条件")
    if not isinstance(settings, ShiftSettings):
        _error("勤務条件の形式が不正です。", setting_source)
    canonical_mode = {"2交代制": TWO_SHIFT, "3交代制": THREE_SHIFT}.get(settings.mode, settings.mode) if isinstance(settings.mode, str) else settings.mode
    if not isinstance(canonical_mode, str) or canonical_mode not in SHIFT_MODES:
        _error("勤務区分は日勤のみ・２交代制・３交代制から選んでください。", setting_source)
    settings = replace(settings, mode=canonical_mode)
    if type(settings.leave_enabled) is not bool:
        _error("公休・有給・明けの区分は有効／無効で指定してください。", setting_source)
    if settings.day_count_mode not in (DIRECT_DAYS, PUBLIC_DAYS):
        _error("日数の指定方法は勤務日数を指定・公休日数から算出から選んでください。", setting_source)
    if settings.default_public_holidays is not None and (not _integer(settings.default_public_holidays) or settings.default_public_holidays < 0):
        _error("共通の公休日数は0以上の整数、または未設定にしてください。", setting_source)
    if not settings.leave_enabled and (data.paid_leaves or any(e.public_holidays is not None for e in data.employees) or settings.day_count_mode != DIRECT_DAYS):
        _error("公休数・有給の入力があります。公休・有給・明けの区分を有効にしてください。", setting_source)
    if type(settings.shift_times_confirmed) is not bool:
        _error("勤務時刻の確認は有効／無効で指定してください。", setting_source)
    if not isinstance(settings.shift_times, (tuple, list)):
        _error("勤務時刻は勤務区分ごとの一覧で指定してください。", setting_source)
    global_labels = {label for group in SHIFT_MODES.values() for label in group}
    time_labels = []
    for shift_time in settings.shift_times:
        if not isinstance(shift_time, ShiftTime) or shift_time.label not in global_labels:
            _error("勤務時刻の区分名が不正です。", setting_source)
        if not _integer(shift_time.start_minute) or not _integer(shift_time.end_minute) or not 0 <= shift_time.start_minute < 1440 or not shift_time.start_minute < shift_time.end_minute <= shift_time.start_minute + 1440:
            _error(f"{shift_time.label}の時刻は開始0〜1439分、終了は開始より後で24時間以内にしてください。翌日分は1440を加えます。", setting_source)
        time_labels.append(shift_time.label)
    if len(time_labels) != len(set(time_labels)):
        _error("勤務時刻の勤務区分が重複しています。", setting_source)
    settings = replace(settings, shift_times=tuple(settings.shift_times))
    if settings.leave_enabled:
        if settings.mode != DAY_ONLY and (not settings.shift_times_confirmed or not set(SHIFT_MODES[settings.mode]).issubset(time_labels)):
            _error("交代制で公休・有給・明けを扱うには、有効な全勤務区分の時刻を入力し、確認済みにしてください。", setting_source)
        if settings.shift_times and not settings.shift_times_confirmed:
            _error("勤務時刻を確認済みにしてください。日跨ぎと明けの分類に使用します。", setting_source)
    if type(settings.night_rest) is not bool:
        _error("夜勤翌日休みの設定は有効／無効で指定してください。", setting_source)
    transitions = settings.forbidden_transitions
    if not isinstance(transitions, (tuple, list)) or any(not isinstance(pair, (tuple, list)) or len(pair) != 2 for pair in transitions):
        _error("禁止する勤務順序は前日区分・翌日区分の組で指定してください。", setting_source)
    allowed = assignment_labels(settings)
    if any(not isinstance(label, str) or label not in allowed for pair in transitions for label in pair):
        _error("禁止する勤務順序には選択した勤務区分と有効な休み区分だけを指定してください。", setting_source)
    transitions = tuple(tuple(pair) for pair in transitions)
    if len(transitions) != len(set(transitions)):
        _error("禁止する勤務順序が重複しています。", setting_source)
    settings = replace(settings, forbidden_transitions=transitions)
    limit = settings.max_consecutive_days
    if limit is not None and (not _integer(limit) or limit < 1):
        _error("連続勤務上限は1以上の整数にしてください。", setting_source)
    try:
        # 人数0でも配分設定そのものを検証する。
        allocate_shifts(0, settings.mode, settings.shift_weights)
        ratios = settings.weekday_ratios
        if ratios is not None:
            if not isinstance(ratios, (tuple, list)) or len(ratios) != 7:
                _error("曜日別の配置比率は月曜日から日曜日まで7個指定してください。")
            ratios = tuple(_percentage(value) for value in ratios)
    except InputError as exc:
        raise InputError([replace(issue, source=setting_source) for issue in exc.issues]) from exc
    weekly_ratios = _ratio_overrides(settings.weekly_ratios, weekly=True, source=setting_source)
    date_ratios = _ratio_overrides(settings.date_ratios, weekly=False, source=setting_source)
    period_days = {requirement.day for requirement in data.requirements if type(requirement.day) is date}
    for start, _ in weekly_ratios:
        if not any(0 <= (day - start).days < 7 for day in period_days):
            _error(f"週指定{start}の週が対象期間と交わりません。", setting_source)
    for day, _ in date_ratios:
        if day not in period_days:
            _error(f"日付指定{day}は対象期間外です。", setting_source)
    settings = replace(settings, weekday_ratios=ratios, shift_weights=tuple(settings.shift_weights),
                       weekly_ratios=weekly_ratios, date_ratios=date_ratios)
    labels = SHIFT_MODES[settings.mode]
    requirements = []
    for requirement in data.requirements:
        if not _integer(requirement.required) or requirement.required < 0:
            _error("配置人数は0以上の整数にしてください。", requirement.source)
        use_ratio = bool(ratios is not None or weekly_ratios or date_ratios)
        if use_ratio and type(requirement.day) is not date:
            _error("必要人数の日付は年を含む日付にしてください。", requirement.source)
        ratio = attendance_ratio_for_day(requirement.day, settings) if use_ratio else None
        if ratio is not None:
            total = required_from_ratio(len(data.employees), ratio)
            counts = allocate_shifts(total, settings.mode, settings.shift_weights)
        else:
            total = requirement.required
            counts = requirement.shift_counts
            if counts:
                if not isinstance(counts, (tuple, list)) or any(not isinstance(pair, (tuple, list)) or len(pair) != 2 for pair in counts):
                    _error("勤務区分別人数の形式が不正です。", requirement.source)
                names = [pair[0] for pair in counts]
                known_labels = {label for group in SHIFT_MODES.values() for label in group}
                if any(not isinstance(name, str) for name in names) or len(names) != len(set(names)) or any(name not in known_labels for name in names):
                    _error("勤務区分別人数の区分名が不正または重複しています。", requirement.source)
                if any(not _integer(count) or count < 0 for _, count in counts):
                    _error("勤務区分別人数は0以上の整数にしてください。", requirement.source)
                if any(count != 0 for name, count in counts if name not in labels):
                    _error("選択していない勤務区分の人数は0または空欄にしてください。", requirement.source)
                counts = tuple((name, count) for name, count in counts if name in labels)
                names = [name for name, _ in counts]
                if set(names) != set(labels):
                    _error(f"勤務区分別人数には{'・'.join(labels)}を各1回指定してください（0人も明示）。", requirement.source)
                if any(not _integer(count) or count < 0 for _, count in counts) or sum(count for _, count in counts) != total:
                    _error("勤務区分別人数は0以上の整数とし、合計を配置人数に合わせてください。", requirement.source)
                by_label = dict(counts)
                counts = tuple((label, by_label[label]) for label in labels)
            else:
                counts = allocate_shifts(total, settings.mode, settings.shift_weights)
        requirements.append(replace(requirement, required=total, shift_counts=counts))
    employees = data.employees
    if settings.leave_enabled:
        for employee in employees:
            if employee.public_holidays is not None and (not _integer(employee.public_holidays) or employee.public_holidays < 0):
                _error("個別の公休日数は0以上の整数にしてください。", employee.source)
        if settings.day_count_mode == PUBLIC_DAYS:
            paid_counts = Counter(p.employee_id for p in data.paid_leaves)
            crosses = crossing_shifts(settings)
            has_recovery = bool(crosses & set(labels)) or any(previous_shift_label(e, settings) in crosses for e in employees)
            updated = []
            for employee in employees:
                public = public_holiday_target(employee, settings)
                if public is None:
                    _error("公休日数から算出する場合、共通または個別の公休日数を指定してください。", employee.source)
                work_days = None if has_recovery else len(requirements) - public - paid_counts[employee.id]
                updated.append(replace(employee, work_days=work_days))
            employees = tuple(updated)
    return replace(data, employees=employees, requirements=tuple(requirements), settings=settings)


def previous_shift_label(employee, mode):
    """直前日の区分。高度モードでは連勤0日の未記入を公休相当とする。"""
    settings = mode if isinstance(mode, ShiftSettings) else None
    actual_mode = settings.mode if settings is not None else mode
    if employee.previous_shift is not None:
        # 旧形式の「休み」と未入力0日を同じ互換仮定で遷移判定する。
        if settings is not None and settings.leave_enabled and employee.previous_shift == OFF:
            return PUBLIC_OFF
        return employee.previous_shift
    if employee.previous_work_days == 0:
        return PUBLIC_OFF if settings is not None and settings.leave_enabled else OFF
    return "日勤" if actual_mode == DAY_ONLY else None


def crossing_shifts(settings: ShiftSettings) -> frozenset[str]:
    """明けを伴う勤務。時刻の妥当性・確認状態はconfigureで検証する。"""
    return frozenset(shift.label for shift in settings.shift_times if shift.end_minute > 1440)


def public_holiday_target(employee, settings: ShiftSettings) -> int | None:
    if not settings.leave_enabled or settings.day_count_mode != PUBLIC_DAYS:
        return None
    base = employee.public_holidays if employee.public_holidays is not None else settings.default_public_holidays
    return None if base is None else base + employee.public_holiday_adjustment
