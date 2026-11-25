"""時間切れ・最適性の表示を実解と独立に検証する回帰テスト。"""
from dataclasses import replace
from datetime import date, timedelta

import pytest
from ortools.sat.python import cp_model

import sinta_shift.solver as solver_module
from sinta_shift.evaluation import evaluate_schedule
from sinta_shift.models import DayRequirement, Employee, FixedOff, InputData, SolverStatus, Source
from sinta_shift.sample import sample_data
from sinta_shift.validation import validate_input


def _controlled_solvers(monkeypatch, statuses, after_solve=None):
    """実解を作る実エンジンに、再現したい終了状態だけを重ねる。"""
    original = cp_model.CpSolver
    remaining = iter(statuses)
    engines = []

    class ControlledSolver:
        def __init__(self):
            self.inner = original()
            self.parameters = self.inner.parameters
            self.status = next(remaining)
            engines.append(self)

        def solve(self, model):
            if self.status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                assert self.inner.solve(model) == cp_model.OPTIMAL
            if after_solve is not None:
                after_solve(self)
            return self.status

        def value(self, variable):
            return self.inner.value(variable)

        @property
        def objective_value(self):
            return self.inner.objective_value

    monkeypatch.setattr(cp_model, "CpSolver", ControlledSolver)
    return engines


def test_first_phase_feasible_preserves_valid_plan_without_claiming_optimum(monkeypatch):
    engines = _controlled_solvers(monkeypatch, [cp_model.FEASIBLE])
    data = sample_data()
    result = solver_module.solve_schedule(data)
    assert result.status is SolverStatus.FEASIBLE
    assert evaluate_schedule(data, result.schedule).valid
    assert not result.count_optimal and not result.priority_optimal
    assert len(engines) == 1


def test_second_phase_unknown_retains_first_plan_and_count_proof(monkeypatch):
    _controlled_solvers(monkeypatch, [cp_model.OPTIMAL, cp_model.UNKNOWN])
    data = sample_data()
    result = solver_module.solve_schedule(data)
    evaluation = evaluate_schedule(data, result.schedule)
    assert result.status is SolverStatus.FEASIBLE
    assert evaluation.valid and evaluation.adopted_count == 3
    assert result.count_optimal and not result.priority_optimal
    assert "順位合計の最小化は未確認" in result.message


def test_first_phase_unknown_does_not_claim_infeasibility(monkeypatch):
    _controlled_solvers(monkeypatch, [cp_model.UNKNOWN])
    result = solver_module.solve_schedule(sample_data())
    assert result.status is SolverStatus.UNKNOWN
    assert result.schedule is None
    assert "存在しないという意味ではありません" in result.message


@pytest.mark.parametrize("statuses", [
    [cp_model.MODEL_INVALID],
    [cp_model.OPTIMAL, cp_model.MODEL_INVALID],
])
def test_model_invalid_is_distinct_from_infeasible_and_timeout(monkeypatch, statuses):
    _controlled_solvers(monkeypatch, statuses)
    result = solver_module.solve_schedule(sample_data())
    assert result.status is SolverStatus.MODEL_ERROR
    assert result.schedule is None
    assert "モデルエラー" in result.message


def test_two_optimization_phases_share_one_total_budget(monkeypatch):
    now = [0.0]
    validation = solver_module.validate_input
    def validate_with_elapsed_second(data):
        issues = validation(data)
        now[0] = 1.0
        return issues
    def consume_search_time(engine):
        now[0] += 10.0 if len(engines) == 1 else 1.0
    engines = _controlled_solvers(monkeypatch, [cp_model.OPTIMAL, cp_model.OPTIMAL], consume_search_time)
    monkeypatch.setattr(solver_module, "monotonic", lambda: now[0])
    monkeypatch.setattr(solver_module, "validate_input", validate_with_elapsed_second)
    result = solver_module.solve_schedule(sample_data(), time_limit=30)
    assert result.status is SolverStatus.OPTIMAL
    assert [engine.parameters.max_time_in_seconds for engine in engines] == [29.0, 19.0]
    assert result.elapsed_seconds == 12.0


def test_cpsat_proves_conflict_that_per_day_and_per_person_checks_cannot_detect():
    days = tuple(date(2026, 11, 7) + timedelta(days=n) for n in range(4))
    # A・Bは前半2日しか勤務できず、Cは指定勤務日数0。
    # 各日の単純な勤務可能人数と各人の勤務可能日数は足りるが、
    # 前半2日の定員へA・Bの合計4勤務を入れることはできない。
    data = InputData(
        employees=(Employee("A", "職員A", 2), Employee("B", "職員B", 2), Employee("C", "職員C", 0)),
        requirements=tuple(DayRequirement(day, 1) for day in days),
        fixed_off=tuple(FixedOff(employee, day) for employee in ("A", "B") for day in days[2:]),
    )
    assert validate_input(data) == []
    result = solver_module.solve_schedule(data)
    assert result.status is SolverStatus.INFEASIBLE
    assert result.schedule is None


def test_model_construction_stops_before_all_cells_when_budget_expires(monkeypatch):
    days = tuple(date(2026, 11, 7) + timedelta(days=n) for n in range(500))
    data = InputData(
        (Employee("A", "職員A", 250), Employee("B", "職員B", 250)),
        tuple(DayRequirement(day, 1) for day in days),
    )
    now = [0.0]
    created = [0]
    original = cp_model.CpModel.new_bool_var
    def slow_variable(model, name):
        variable = original(model, name)
        created[0] += 1
        now[0] += 0.01
        return variable
    monkeypatch.setattr(solver_module, "monotonic", lambda: now[0])
    monkeypatch.setattr(cp_model.CpModel, "new_bool_var", slow_variable)
    result = solver_module.solve_schedule(data, time_limit=1)
    assert result.status is SolverStatus.UNKNOWN
    assert result.schedule is None
    assert 0 < created[0] < len(data.employees) * len(days)
    assert "存在しないという意味ではありません" in result.message


@pytest.mark.parametrize("method", ["new_bool_var", "validate"])
def test_model_construction_and_validation_exceptions_are_model_errors(monkeypatch, method):
    def broken(*args, **kwargs):
        raise RuntimeError("test library failure")
    monkeypatch.setattr(cp_model.CpModel, method, broken)
    result = solver_module.solve_schedule(sample_data())
    assert result.status is SolverStatus.MODEL_ERROR
    assert result.schedule is None
    assert "RuntimeError" in result.message


def test_expiry_while_preparing_second_phase_preserves_first_plan(monkeypatch):
    now = [0.0]
    engines = _controlled_solvers(monkeypatch, [cp_model.OPTIMAL])
    original = cp_model.CpModel.add_hint
    def slow_hint(model, variable, value):
        original(model, variable, value)
        now[0] = 31.0
    monkeypatch.setattr(solver_module, "monotonic", lambda: now[0])
    monkeypatch.setattr(cp_model.CpModel, "add_hint", slow_hint)
    data = sample_data()
    result = solver_module.solve_schedule(data, time_limit=30)
    assert result.status is SolverStatus.FEASIBLE
    assert result.count_optimal and not result.priority_optimal
    assert evaluate_schedule(data, result.schedule).valid
    assert len(engines) == 1


def test_aggregate_validation_issue_retains_input_filename():
    data = sample_data()
    employees = tuple(replace(e, source=Source("mapped_input.xlsx", "メンバー", i + 2)) for i, e in enumerate(data.employees))
    employees = (replace(employees[0], work_days=1), *employees[1:])
    data = replace(data, employees=employees)
    issues = validate_input(data)
    mismatch = next(issue for issue in issues if "合計" in issue.message)
    assert mismatch.source.file == "mapped_input.xlsx"
    assert mismatch.source.row == 0  # 合計の不一致には単一の原因行を断定しない。
