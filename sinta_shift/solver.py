"""CP-SATによる辞書式の二段階最適化。GUI非依存。"""
import math
from time import monotonic
from collections import defaultdict
from itertools import chain
from ortools.sat.python import cp_model
from .models import DAY_ONLY, OFF, PUBLIC_OFF, PAID_LEAVE, RECOVERY, PUBLIC_DAYS, SHIFT_MODES, InputData, InputError, SolveResult, SolverStatus
from .planning import configure_input, previous_shift_label, crossing_shifts, public_holiday_target
from .validation import validate_input


class _BudgetExpired(Exception):
    """探索だけでなくモデル構築も共通の時間予算を消費する。"""


def solve_schedule(data: InputData, time_limit: float = 30) -> SolveResult:
    started = monotonic()
    if not isinstance(time_limit, (int, float)) or isinstance(time_limit, bool) or not math.isfinite(time_limit) or time_limit <= 0:
        return SolveResult(SolverStatus.MODEL_ERROR, message="制限時間は0より大きい有限の秒数にしてください。")
    deadline = started + time_limit
    schedule = None
    count_optimal = False
    priority_optimal = False

    def check_budget():
        if monotonic() >= deadline:
            raise _BudgetExpired

    def checked(items):
        # 大きな行・列の式やヒント作成も128項目ごとに期限を確認する。
        for index, item in enumerate(items):
            if index % 128 == 0:
                check_budget()
            yield item

    def unknown_result():
        return SolveResult(SolverStatus.UNKNOWN, message="制限時間内に案が見つかりませんでした。案が存在しないという意味ではありません。時間を延ばして再実行してください。", elapsed_seconds=monotonic() - started)

    def feasible_result():
        optimal = count_optimal and priority_optimal
        message = ("採用申請数の最大化と順位合計の最小化を両方確認しました。" if optimal else
                   "必須条件を満たす案です。採用申請数は最大と確認済みですが、順位合計の最小化は未確認です。" if count_optimal else
                   "必須条件を満たす案です。採用申請数・順位合計の最適性は時間内に確認できていません。")
        return SolveResult(SolverStatus.OPTIMAL if optimal else SolverStatus.FEASIBLE, schedule, count_optimal,
                           priority_optimal, monotonic() - started, message)

    try:
        try:
            data = configure_input(data)
        except InputError as exc:
            return SolveResult(SolverStatus.INFEASIBLE, message=str(exc), issues=exc.issues,
                               elapsed_seconds=monotonic() - started)
        issues = validate_input(data)
        check_budget()
        if issues:
            return SolveResult(SolverStatus.INFEASIBLE, message="\n".join(i.message for i in issues), issues=issues,
                               elapsed_seconds=monotonic() - started)
        days = data.days
        model = cp_model.CpModel()
        x = {}
        for i, employee in enumerate(checked(data.employees)):
            for j, day in enumerate(checked(days)):
                x[employee.id, day] = model.new_bool_var(f"work_{i}_{j}")
        labels = SHIFT_MODES[data.settings.mode]
        shifts = {}
        extra_shift_variables = []
        for i, employee in enumerate(checked(data.employees)):
            for j, day in enumerate(checked(days)):
                if data.settings.mode == DAY_ONLY:
                    shifts[employee.id, day, labels[0]] = x[employee.id, day]
                else:
                    for shift_index, label in enumerate(labels):
                        variable = model.new_bool_var(f"shift_{i}_{j}_{shift_index}")
                        shifts[employee.id, day, label] = variable
                        extra_shift_variables.append(variable)
                    model.add(sum(shifts[employee.id, day, label] for label in labels) == x[employee.id, day])
        for requirement in checked(data.requirements):
            model.add(sum(x[e.id, requirement.day] for e in checked(data.employees)) == requirement.required)
            if data.settings.mode != DAY_ONLY:
                for label, count_required in requirement.shift_counts:
                    model.add(sum(shifts[e.id, requirement.day, label] for e in checked(data.employees)) == count_required)
        advanced = data.settings.leave_enabled
        state_variables = {}
        crosses = crossing_shifts(data.settings) if advanced else frozenset()
        paid_cells = {(paid.employee_id, paid.day) for paid in data.paid_leaves}
        if advanced:
            for i, employee in enumerate(checked(data.employees)):
                for j, day in enumerate(checked(days)):
                    for state_index, label in enumerate((PUBLIC_OFF, PAID_LEAVE, RECOVERY)):
                        variable = model.new_bool_var(f"state_{i}_{j}_{state_index}")
                        state_variables[employee.id, day, label] = variable
                        extra_shift_variables.append(variable)
                    model.add(x[employee.id, day] + sum(state_variables[employee.id, day, label] for label in (PUBLIC_OFF, PAID_LEAVE, RECOVERY)) == 1)
                    model.add(state_variables[employee.id, day, PAID_LEAVE] == int((employee.id, day) in paid_cells))
                    recovery = state_variables[employee.id, day, RECOVERY]
                    if j == 0:
                        model.add(recovery == int(previous_shift_label(employee, data.settings) in crosses))
                    else:
                        model.add(recovery == sum(shifts[employee.id, days[j - 1], label] for label in labels if label in crosses))
        for employee in checked(data.employees):
            if advanced and data.settings.day_count_mode == PUBLIC_DAYS:
                model.add(sum(state_variables[employee.id, d, PUBLIC_OFF] for d in checked(days)) == public_holiday_target(employee, data.settings))
            else:
                model.add(sum(x[employee.id, d] for d in checked(days)) == employee.work_days)
        for fixed in checked(data.fixed_off):
            model.add(state_variables[fixed.employee_id, fixed.day, PUBLIC_OFF] == 1 if advanced else x[fixed.employee_id, fixed.day] == 0)
        def actual_work(employee_id, day):
            return x[employee_id, day] + state_variables[employee_id, day, RECOVERY] if advanced else x[employee_id, day]
        limit = data.settings.max_consecutive_days
        if limit is not None:
            for employee in checked(data.employees):
                for start in checked(range(max(0, len(days) - limit))):
                    model.add(sum(actual_work(employee.id, day) for day in checked(days[start:start + limit + 1])) <= limit)
                if employee.previous_work_days:
                    boundary = limit - employee.previous_work_days + 1
                    if boundary <= len(days):
                        model.add(sum(actual_work(employee.id, day) for day in checked(days[:boundary])) <= boundary - 1)
        def label_expression(employee_id, day, label):
            if advanced and label in (PUBLIC_OFF, PAID_LEAVE, RECOVERY):
                return state_variables[employee_id, day, label]
            return 1 - x[employee_id, day] if label == OFF else shifts[employee_id, day, label]

        night_labels = tuple(label for label in labels if label in ("夜勤", "深夜勤"))
        for employee in checked(data.employees):
            for day_index in checked(range(len(days) - 1)):
                day, following = days[day_index:day_index + 2]
                if data.settings.night_rest:
                    for label in night_labels:
                        model.add(shifts[employee.id, day, label] + x[employee.id, following] <= 1)
                for earlier, later in data.settings.forbidden_transitions:
                    model.add(label_expression(employee.id, day, earlier) + label_expression(employee.id, following, later) <= 1)
            previous = previous_shift_label(employee, data.settings)
            if data.settings.night_rest and previous in ("夜勤", "深夜勤"):
                model.add(x[employee.id, days[0]] == 0)
            for earlier, later in data.settings.forbidden_transitions:
                if previous == earlier:
                    model.add(label_expression(employee.id, days[0], later) == 0)
        per_request = defaultdict(list)
        per_rest = defaultdict(list)
        ys = []
        for index, candidate in enumerate(checked(data.requests)):
            chosen = model.new_bool_var(f"request_{index}")
            if advanced:
                model.add(chosen <= state_variables[candidate.employee_id, candidate.day, PUBLIC_OFF])
            else:
                model.add(chosen + x[candidate.employee_id, candidate.day] <= 1)
            ys.append(chosen)
            per_request[candidate.request_id].append(chosen)
            per_rest[candidate.employee_id, candidate.day].append(chosen)
        for group in checked(chain(per_request.values(), per_rest.values())):
            model.add(sum(checked(group)) <= 1)
        count = sum(checked(ys))
        priority = sum(y * c.priority for y, c in checked(zip(ys, data.requests)))
        model.maximize(count)
        check_budget()
        error = model.validate()
        check_budget()
        if error:
            return SolveResult(SolverStatus.MODEL_ERROR, message=f"CP-SATモデルを検証できません：{error}", elapsed_seconds=monotonic() - started)

        def run():
            remaining = deadline - monotonic()
            if remaining <= 0:
                return None, cp_model.UNKNOWN
            engine = cp_model.CpSolver()
            engine.parameters.max_time_in_seconds = remaining
            engine.parameters.num_search_workers = 1
            engine.parameters.random_seed = 0
            return engine, engine.solve(model)

        def extract_schedule(engine):
            if advanced:
                all_labels = (*labels, PUBLIC_OFF, PAID_LEAVE, RECOVERY)
                return {(employee.id, day): next(label for label in all_labels if engine.value(label_expression(employee.id, day, label)))
                        for employee in data.employees for day in days}
            if data.settings.mode == DAY_ONLY:
                return {key: bool(engine.value(value)) for key, value in x.items()}
            return {(employee.id, day): next((label for label in labels if engine.value(shifts[employee.id, day, label])), OFF)
                    for employee in data.employees for day in days}

        engine, status = run()
        if status == cp_model.INFEASIBLE:
            return SolveResult(SolverStatus.INFEASIBLE, message="必須条件を同時に満たす案はありません。確定休・勤務日数・公休・有給・明け・配置人数と、設定した連勤上限・夜勤翌日休み・禁止順序を見直してください。", elapsed_seconds=monotonic() - started)
        if status == cp_model.MODEL_INVALID:
            return SolveResult(SolverStatus.MODEL_ERROR, message="CP-SATがモデルエラーを返しました。", elapsed_seconds=monotonic() - started)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return unknown_result()
        schedule = extract_schedule(engine)
        count_optimal = status == cp_model.OPTIMAL
        # 最大件数が未証明なら順位のために探索時間を消費しない。
        # 第1段階で得た実行可能案をそのまま表示する。
        if count_optimal:
            check_budget()
            best_count = int(round(engine.objective_value))
            model.add(count == best_count)
            model.minimize(priority)
            for variable in checked(chain(x.values(), extra_shift_variables, ys)):
                model.add_hint(variable, engine.value(variable))
            second, second_status = run()
            if second_status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                schedule = extract_schedule(second)
                priority_optimal = second_status == cp_model.OPTIMAL
            elif second_status == cp_model.MODEL_INVALID:
                return SolveResult(SolverStatus.MODEL_ERROR, message="第2段階でモデルエラーが発生しました。", elapsed_seconds=monotonic() - started)
        return feasible_result()
    except _BudgetExpired:
        # 第2段階の準備中に期限へ達しても、第1段階の完全な案は保持する。
        return feasible_result() if schedule is not None else unknown_result()
    except Exception as exc:
        return SolveResult(SolverStatus.MODEL_ERROR, message=f"計算に失敗しました：{type(exc).__name__}: {exc}", elapsed_seconds=monotonic() - started)


solve = solve_schedule
