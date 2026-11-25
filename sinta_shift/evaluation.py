"""手修正した勤務表も同じ規則で検証し、希望申請の採否を再計算する。"""
from collections import deque
from .models import OFF, PUBLIC_OFF, PAID_LEAVE, RECOVERY, PUBLIC_DAYS, PAID_REQUEST, SHIFT_MODES, Evaluation, InputData, InputError, Issue, RequestResult, Schedule, Source, assignment_label, assignment_labels, is_working, is_actual_working
from .planning import configure_input, previous_shift_label, crossing_shifts, public_holiday_target
from .validation import validate_input


def _requests(data: InputData, schedule: Schedule) -> list[RequestResult]:
    """最大流・最小費用の二部マッチング。1休みを1申請だけに対応付ける。

    増加路がなくなるまで1件ずつ流すため、費用の大きさによらず採用数が
    最優先となる。最短増加路の残余逆辺により既存の採用を付け替えられる。
    """
    request_ids = list(dict.fromkeys(c.request_id for c in data.requests))
    request_employee = {c.request_id: c.employee_id for c in data.requests}
    rests = list(dict.fromkeys((c.employee_id, c.day) for c in data.requests
                              if assignment_label(schedule.get((c.employee_id, c.day)), data.settings) == (PUBLIC_OFF if data.settings.leave_enabled else OFF)))
    req_node = {r: n + 1 for n, r in enumerate(request_ids)}
    rest_node = {r: n + 1 + len(req_node) for n, r in enumerate(rests)}
    sink = len(req_node) + len(rest_node) + 1
    graph: list[list[list[int]]] = [[] for _ in range(sink + 1)]
    def edge(a, b, cost):
        forward = [b, len(graph[b]), 1, cost]
        reverse = [a, len(graph[a]), 0, -cost]
        graph[a].append(forward)
        graph[b].append(reverse)
        return forward
    for node in req_node.values():
        edge(0, node, 0)
    for node in rest_node.values():
        edge(node, sink, 0)
    references = []
    for candidate in data.requests:
        rest = (candidate.employee_id, candidate.day)
        if rest in rest_node:
            references.append((candidate, edge(req_node[candidate.request_id], rest_node[rest], candidate.priority)))
    while True:
        distances = [float("inf")] * len(graph)
        previous = [None] * len(graph)
        distances[0] = 0
        pending = deque([0])
        in_queue = {0}
        while pending:
            a = pending.popleft()
            in_queue.remove(a)
            for index, item in enumerate(graph[a]):
                b, _, capacity, cost = item
                if capacity and distances[b] > distances[a] + cost:
                    distances[b] = distances[a] + cost
                    previous[b] = (a, index)
                    if b not in in_queue:
                        pending.append(b)
                        in_queue.add(b)
        if previous[sink] is None:
            break
        node = sink
        while node != 0:
            a, index = previous[node]
            item = graph[a][index]
            item[2] -= 1
            graph[node][item[1]][2] += 1
            node = a
    adopted = {c.request_id: c for c, e in references if e[2] == 0}
    return [RequestResult(r, request_employee[r], adopted[r].day if r in adopted else None,
                          adopted[r].priority if r in adopted else None,
                          f"第{adopted[r].priority}希望を採用" if r in adopted else "未採用")
            for r in request_ids]


def _paid_results(data: InputData) -> list[RequestResult]:
    """有給の確認状態は公休希望の採用数・最適化とは別に表示する。"""
    rows = []
    for request in data.independent_requests:
        if request.leave_type != PAID_REQUEST:
            continue
        confirmed = isinstance(request.request_id, str) and request.request_id in data.confirmed_paid_request_ids
        rows.append(RequestResult(request.request_id, request.employee_id,
                                  request.day if confirmed else None, None,
                                  "有給確定" if confirmed else "確認待ち"))
    return rows


def evaluate_schedule(data: InputData, schedule: Schedule) -> Evaluation:
    try:
        data = configure_input(data)
    except InputError as exc:
        return Evaluation(exc.issues, [], {}, {}, paid_request_results=_paid_results(data))
    issues = validate_input(data)
    if issues:
        return Evaluation(issues, [], {}, {}, paid_request_results=_paid_results(data))
    days = data.days
    settings = data.settings
    mode = settings.mode
    advanced = settings.leave_enabled
    labels = SHIFT_MODES[mode]
    crosses = crossing_shifts(settings) if advanced else frozenset()
    expected = {(e.id, day) for e in data.employees for day in days}
    day_counts = {day: 0 for day in days}
    employee_counts = {e.id: 0 for e in data.employees}
    daily_shift_counts = {day: {label: 0 for label in labels} for day in days}
    employee_shift_counts = {e.id: {label: 0 for label in labels} for e in data.employees}
    public_counts = {e.id: 0 for e in data.employees}
    paid_counts = {e.id: 0 for e in data.employees}
    recovery_counts = {e.id: 0 for e in data.employees}
    actual_counts = {e.id: 0 for e in data.employees}
    approved_paid = {(paid.employee_id, paid.day) for paid in data.paid_leaves}
    input_file = next((r.source.file for rows in (data.employees, data.requirements) for r in rows if r.source.file), "")

    def cell_source(employee_index, day_index):
        number = day_index + 3
        column = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            column = chr(65 + remainder) + column
        cell = f"{column}{employee_index + 4}"
        return Source(input_file, "勤務表", employee_index + 4, cell)

    def cell_issue(kind, reason, employee_index, day_index):
        employee = data.employees[employee_index]
        day = days[day_index]
        source = cell_source(employee_index, day_index)
        issues.append(Issue(kind, f"セル{source.column}（職員{employee.id}・{day}）：{reason}", source, employee.id, day))

    for employee_index, employee in enumerate(data.employees):
        for day_index, day in enumerate(days):
            key = (employee.id, day)
            label = assignment_label(schedule.get(key), settings)
            if key not in schedule or label is None:
                cell_issue("区分違反", f"区分は{'／'.join(assignment_labels(settings))}のいずれか1つにしてください。", employee_index, day_index)
            elif label in labels:
                day_counts[day] += 1
                employee_counts[employee.id] += 1
                daily_shift_counts[day][label] += 1
                employee_shift_counts[employee.id][label] += 1
            if is_actual_working(schedule.get(key), settings):
                actual_counts[employee.id] += 1
            if advanced:
                public_counts[employee.id] += int(label == PUBLIC_OFF)
                paid_counts[employee.id] += int(label == PAID_LEAVE)
                recovery_counts[employee.id] += int(label == RECOVERY)
                if key in approved_paid and label != PAID_LEAVE:
                    cell_issue("有給違反", "承認済みの有給日には有給を割り当ててください。公休や明けで代用できません。", employee_index, day_index)
                elif key not in approved_paid and label == PAID_LEAVE:
                    cell_issue("未承認有給", "有給一覧にない日へ有給を自動生成できません。", employee_index, day_index)
    for key in schedule.keys() - expected:
        issues.append(Issue("範囲違反", f"対象外の職員・日付のセルがあります：{key}", Source(input_file, "勤務表")))
    for day_index, requirement in enumerate(sorted(data.requirements, key=lambda r: r.day)):
        count = day_counts[requirement.day]
        if count != requirement.required:
            issues.append(Issue("日別人数違反", f"{requirement.day}：日別人数は必要{requirement.required}人・現在{count}人。明け・公休・有給は勤務開始人数に含めません。",
                                requirement.source, day=requirement.day))
        shift_errors = []
        for label, required_count in requirement.shift_counts:
            actual = daily_shift_counts[requirement.day][label]
            if actual != required_count:
                shift_errors.append(f"{label}は必要{required_count}人・現在{actual}人")
        if shift_errors:
            detail = "、".join(shift_errors)
            issues.append(Issue("勤務区分人数違反", f"{requirement.day}：{detail}。該当日の区分を確認してください。", requirement.source, day=requirement.day))
            for employee_index, employee in enumerate(data.employees):
                cell_issue("セル確認", f"{detail}。区分を確認してください。", employee_index, day_index)
    use_work_target = not advanced or settings.day_count_mode != PUBLIC_DAYS
    for employee in data.employees:
        count = employee_counts[employee.id]
        if use_work_target and employee.work_days is not None and count != employee.work_days:
            issues.append(Issue("勤務日数違反", f"職員{employee.id}：勤務日数（勤務開始回数）は必要{employee.work_days}日・現在{count}日。該当職員の各セルを確認してください。",
                                employee.source, employee.id))
        target = public_holiday_target(employee, settings)
        if target is not None and public_counts[employee.id] != target:
            issues.append(Issue("公休日数違反", f"職員{employee.id}：公休は必要{target}日・現在{public_counts[employee.id]}日。有給・明けは公休に含めません。", employee.source, employee.id))
    required = {r.day: r.required for r in data.requirements}
    for employee_index, employee in enumerate(data.employees):
        if use_work_target and employee.work_days is not None and employee_counts[employee.id] != employee.work_days:
            for day_index, day in enumerate(days):
                if day_counts[day] != required[day]:
                    cell_issue("セル確認", "日別人数違反と勤務日数違反の交点です。区分を確認してください。", employee_index, day_index)
        if settings.max_consecutive_days is not None:
            run = employee.previous_work_days
            for day_index, day in enumerate(days):
                run = run + 1 if is_actual_working(schedule.get((employee.id, day)), settings) else 0
                if run > settings.max_consecutive_days:
                    history = f"（期間開始前の連勤{employee.previous_work_days}日を含む）" if employee.previous_work_days else ""
                    cell_issue("連続勤務違反", f"連続勤務{run}日{history}で上限{settings.max_consecutive_days}日を超えています。明けも実労働のある暦日として数えます。", employee_index, day_index)
    for employee_index, employee in enumerate(data.employees):
        for day_index, day in enumerate(days):
            previous = previous_shift_label(employee, settings) if day_index == 0 else assignment_label(schedule.get((employee.id, days[day_index - 1])), settings)
            current = assignment_label(schedule.get((employee.id, day)), settings)
            reasons = []
            if advanced:
                needs_recovery = previous in crosses
                if needs_recovery and current != RECOVERY:
                    reasons.append(("明け違反", "前日の勤務が日跨ぎのため明けが必要です。公休・有給や新しい勤務と重ねられません"))
                elif not needs_recovery and current == RECOVERY:
                    reasons.append(("明け違反", "前日に日跨ぎ勤務がないため明けを割り当てられません"))
            if settings.night_rest and previous in ("夜勤", "深夜勤") and current in labels:
                reasons.append(("夜勤翌日違反", f"{previous}の翌日は新しい勤務開始ができませんが{current}です"))
            if (previous, current) in settings.forbidden_transitions:
                reasons.append(("勤務順序違反", f"{previous}→{current}は禁止された勤務順序です"))
            for kind, reason in reasons:
                affected = (day_index,) if day_index == 0 else (day_index - 1, day_index)
                period = "期間開始前日の履歴から初日への境界" if day_index == 0 else f"{days[day_index - 1]}から{day}"
                for affected_index in affected:
                    cell_issue(kind, f"{period}で{reason}。", employee_index, affected_index)
    employee_indices = {employee.id: index for index, employee in enumerate(data.employees)}
    day_indices = {day: index for index, day in enumerate(days)}
    for fixed in data.fixed_off:
        current = assignment_label(schedule.get((fixed.employee_id, fixed.day)), settings)
        invalid = current != PUBLIC_OFF if advanced else is_working(schedule.get((fixed.employee_id, fixed.day)), settings)
        if invalid:
            cell_issue("確定休違反", "変更不可の確定公休には公休を割り当ててください。" if advanced else "変更不可の確定休に勤務が割り当てられています。", employee_indices[fixed.employee_id], day_indices[fixed.day])
    return Evaluation(issues, _requests(data, schedule), day_counts, employee_counts, daily_shift_counts, employee_shift_counts,
                      public_counts, paid_counts, recovery_counts, actual_counts, _paid_results(data))
