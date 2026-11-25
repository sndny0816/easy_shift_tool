"""実際の読込・条件検証・手修正Issueから、操作先と修正手順を確かめる。"""
from dataclasses import FrozenInstanceError, replace
from datetime import date, timedelta
from decimal import Decimal
import errno

import pytest
from openpyxl import load_workbook

from sinta_shift.error_guidance import (
    guidance_for_issue, issues_for_exception, issues_for_solve_result,
)
from sinta_shift.excel_io import export_result, export_template, read_input
from sinta_shift.evaluation import evaluate_schedule
from sinta_shift.models import (
    PUBLIC_DAYS, TWO_SHIFT, DayRequirement, Employee, FixedOff,
    IndependentLeaveRequest, InputData, InputError, Issue, PAID_REQUEST,
    PaidLeave, ShiftSettings, ShiftTime, SolveResult, SolverStatus, Source,
)
from sinta_shift.planning import configure_input
from sinta_shift.sample import sample_data
from sinta_shift.solver import solve_schedule
from sinta_shift.validation import validate_input


def _change(path, sheet, cell, value):
    book = load_workbook(path)
    book[sheet][cell] = value
    book.save(path)
    book.close()


def _all_text(guide):
    return "\n".join((guide.title, *guide.steps, guide.next_action))


@pytest.mark.parametrize("sheet,cell,value,category", [
    ("職員", "C2", "=1+1", "formula"),
    ("必要人数", "A2", "11-07", "dates"),
    ("職員", "C2", "不明", "counts"),
    ("職員", "C1", "別の見出し", "columns"),
])
def test_real_excel_issues_have_exact_reason_original_cell_and_reload(tmp_path, sheet, cell, value, category):
    path = export_template(tmp_path)
    _change(path, sheet, cell, value)
    with pytest.raises(InputError) as caught:
        read_input(path)
    issue = caught.value.issues[0]
    guide = guidance_for_issue(issue, operation="read")
    assert guide.reason == issue.message
    assert guide.target == "input"
    assert guide.category == category
    assert path.name in guide.location
    assert sheet in guide.location
    assert "元行" in guide.location
    assert "保存" in _all_text(guide) and "読み直" in _all_text(guide)
    if category == "formula":
        assert "値］" in _all_text(guide)
        assert "古い計算結果" in _all_text(guide)
        assert cell in guide.location


def test_real_unknown_id_matches_staff_table_and_preserves_na_zero_advice(tmp_path):
    path = export_template(tmp_path)
    _change(path, "希望休", "B2", "UNKNOWN")
    data = read_input(path)
    issue = next(i for i in validate_input(data) if "未知の職員ID" in i.message)
    guide = guidance_for_issue(issue, data=data)
    assert (guide.target, guide.category) == ("input", "ids")
    assert guide.reason == issue.message
    assert all(part in guide.location for part in (path.name, "希望休", "元行 2", "UNKNOWN"))
    assert "先頭ゼロ" in _all_text(guide) and "NA" in _all_text(guide)
    assert "推測で補わない" in _all_text(guide)


def test_actual_duplicate_header_and_sheet_error_use_mapping_not_generic_duplicate(tmp_path):
    path = export_template(tmp_path)
    _change(path, "職員", "B1", "職員ID")
    with pytest.raises(InputError) as caught:
        read_input(path)
    guides = [guidance_for_issue(i) for i in caught.value.issues]
    assert guides and all(g.category == "columns" for g in guides)
    assert all(g.target == "input" for g in guides)
    assert "対応" in guides[0].next_action


def test_real_totals_direct_uses_excel_work_start_counts():
    data = sample_data()
    data = replace(data, employees=(replace(data.employees[0], work_days=1), *data.employees[1:]))
    issue = next(i for i in validate_input(data) if "勤務日数合計" in i.message)
    guide = guidance_for_issue(issue, data=data)
    assert (guide.target, guide.category) == ("input", "totals")
    assert "勤務開始回数" in _all_text(guide)
    assert "誤記が確認できた側だけ" in _all_text(guide)
    assert "元行 0" not in guide.location


def test_totals_with_effective_attendance_overrides_targets_settings():
    data = replace(sample_data(), settings=ShiftSettings(weekday_ratios=(Decimal(100),) * 7))
    issue = next(i for i in validate_input(data) if "勤務日数合計" in i.message)
    guide = guidance_for_issue(issue, data=data)
    assert guide.target == "settings"
    assert "元Excelの人数ではなく" in _all_text(guide)
    assert "有効な必要人数" in _all_text(guide)
    assert "自動変更" not in guide.reason


def _derived_data(crossing):
    days = tuple(date(2026, 11, 1) + timedelta(days=i) for i in range(4))
    return InputData(
        (Employee("A", "職員A", None, public_holidays=2),),
        tuple(DayRequirement(day, int(i == 0), shift_counts=(("日勤", 0), ("夜勤", int(i == 0)))) for i, day in enumerate(days)),
        paid_leaves=(PaidLeave("A", days[-1]),),
        settings=ShiftSettings(mode=TWO_SHIFT, leave_enabled=True, day_count_mode=PUBLIC_DAYS,
            shift_times=(ShiftTime("日勤", 480, 1020), ShiftTime("夜勤", 960, 1980 if crossing else 1440)),
            shift_times_confirmed=True),
    )


@pytest.mark.parametrize("crossing", [False, True])
def test_actual_derived_totals_distinguish_recovery_days_and_no_independent_work_target(crossing):
    data = _derived_data(crossing)
    # Without crossing, public2+paid1+start1 balances; force a true mismatch.
    if not crossing:
        data = replace(data, employees=(replace(data.employees[0], public_holidays=1),))
    issue = next(i for i in validate_input(data) if "日数が一致しません" in i.message)
    guide = guidance_for_issue(issue, data=data)
    assert (guide.target, guide.category) == ("settings", "totals")
    text = _all_text(guide)
    assert "入力勤務日数を独立した基準として直しません" in text
    assert "日跨ぎがあるため" in text if crossing else "対象期間の日数−公休目標−確定有給" in text
    assert guide.reason == issue.message


def test_real_daily_capacity_preserves_reason_without_recommending_revoking_fixed_off():
    data = sample_data()
    data = replace(data, fixed_off=tuple(FixedOff(e, data.days[0]) for e in ("A", "B")))
    issue = next(i for i in validate_input(data) if "必要2人・勤務可能1人" in i.message)
    guide = guidance_for_issue(issue, data=data)
    assert guide.category == "capacity"
    assert guide.reason == issue.message
    assert "見直してください" in guide.reason  # Existing Issue is kept verbatim.
    assert "誤記" in _all_text(guide) and "要員確保" in _all_text(guide)
    assert "自動解除" in _all_text(guide)
    assert "確定休を見直す" not in "\n".join(guide.steps)


def test_actual_combined_infeasible_does_not_invent_a_specific_cause():
    days = tuple(date(2026, 11, 1) + timedelta(days=i) for i in range(4))
    data = InputData((Employee("A", "A", 2), Employee("B", "B", 2), Employee("C", "C", 0)),
        tuple(DayRequirement(day, 1) for day in days),
        fixed_off=tuple(FixedOff(employee, day) for employee in ("A", "B") for day in days[2:]))
    assert not validate_input(data)
    result = solve_schedule(data)
    assert result.status is SolverStatus.INFEASIBLE
    guide = guidance_for_issue(issues_for_solve_result(result)[0], data=data)
    assert guide.category == "infeasible"
    assert "組合せが原因かは特定できません" in _all_text(guide)
    assert guide.reason == result.message


def test_real_manual_errors_route_to_app_schedule_not_input_workbook_cells(tmp_path):
    data = read_input(export_template(tmp_path))
    result = solve_schedule(data)
    schedule = dict(result.schedule)
    schedule["A", data.days[0]] = True
    issues = evaluate_schedule(data, schedule).issues
    assert {"日別人数違反", "勤務日数違反", "セル確認"} <= {i.kind for i in issues}
    for issue in issues:
        guide = guidance_for_issue(issue, data=data, has_schedule=True, input_path="original.xlsx")
        assert guide.target == "schedule"
        assert guide.reason == issue.message
        assert "アプリ内の勤務表" in guide.location
        assert data.employees[0].source.file not in guide.location
        assert "元行" not in guide.location and "original.xlsx" not in guide.location
        assert "全必須条件" in guide.next_action
        if issue.kind == "セル確認":
            assert f"Excel出力時のセル {issue.source.column}（参考）" in guide.location
    assert guidance_for_issue(issues[0], has_schedule=False).target == "run"


def test_real_unconfirmed_paid_request_goes_to_explicit_leave_decision():
    data = sample_data()
    data = replace(data, independent_requests=(IndependentLeaveRequest("P1", "A", data.days[0], PAID_REQUEST,
        Source("input.xlsx", "希望休", 2, "C2")),))
    issue = next(i for i in validate_input(data) if i.kind == "有給申請未確認")
    guide = guidance_for_issue(issue, data=data)
    assert (guide.target, guide.category) == ("leave", "leave_confirmation")
    assert "明示" in _all_text(guide) and "保留" in _all_text(guide)
    assert "有給残高" in _all_text(guide)


@pytest.mark.parametrize("settings,category", [
    (ShiftSettings(max_consecutive_days=0), "consecutive_settings"),
    (ShiftSettings(date_ratios=((date(2026, 12, 31), Decimal(50)),)), "attendance_settings"),
    (ShiftSettings(weekly_ratios=((date(2026, 11, 2), Decimal(50)), (date(2026, 11, 2), Decimal(60)))), "attendance_settings"),
    (ShiftSettings(shift_times=(ShiftTime("日勤", 480, 1020),) * 2), "time_settings"),
    (ShiftSettings(mode=TWO_SHIFT, leave_enabled=True), "time_settings"),
    (ShiftSettings(forbidden_transitions=(("休み", "日勤"),) * 2), "sequence_settings"),
])
def test_specific_real_setting_issues_precede_generic_duplicate_or_date_matching(settings, category):
    data = replace(sample_data(), settings=settings)
    with pytest.raises(InputError) as caught:
        configure_input(data)
    issue = caught.value.issues[0]
    guide = guidance_for_issue(issue, data=data, operation="settings")
    assert (guide.target, guide.category) == ("settings", category)
    assert guide.reason == issue.message


def test_setting_source_row_zero_does_not_claim_an_excel_cell_or_force_reload():
    issue = Issue("勤務条件", "連続勤務上限は1以上の整数にしてください。",
                  Source("input.xlsx", "勤務条件", 0, "連続勤務上限"))
    guide = guidance_for_issue(issue)
    assert guide.target == "settings"
    assert "元行 0" not in guide.location and "セル" not in guide.location
    assert "ファイルから復元" in _all_text(guide)
    excel_issue = replace(issue, source=Source("input.xlsx", "勤務条件", 4, "B4"))
    excel_guide = guidance_for_issue(excel_issue)
    assert excel_guide.target == "input" and "読み直" in excel_guide.next_action


def test_actual_timeout_is_not_proof_of_infeasibility_and_can_retry_with_longer_limit():
    result = solve_schedule(sample_data(), time_limit=1e-12)
    assert result.status is SolverStatus.UNKNOWN
    guide = guidance_for_issue(issues_for_solve_result(result)[0])
    assert (guide.target, guide.category) == ("run", "timeout")
    assert "案がないと確定した状態ではありません" in _all_text(guide)
    assert "最大300秒" in _all_text(guide)
    assert "人数不足" not in _all_text(guide)


def test_solver_specific_issues_take_precedence_and_success_has_no_synthetic_issue():
    issue = Issue("入力不備", "actual reason")
    result = SolveResult(SolverStatus.INFEASIBLE, message="summary", issues=[issue])
    assert issues_for_solve_result(result) == (issue,)
    assert not issues_for_solve_result(SolveResult(SolverStatus.FEASIBLE))
    assert not issues_for_solve_result(SolveResult(SolverStatus.OPTIMAL))


def test_model_error_keeps_full_technical_detail_and_does_not_claim_infeasibility():
    message = "計算に失敗しました：RuntimeError: detail\nline2\nline3"
    guide = guidance_for_issue(issues_for_solve_result(SolveResult(SolverStatus.MODEL_ERROR, message=message))[0])
    assert (guide.target, guide.category) == ("support", "model_error")
    assert guide.reason == message
    assert "条件不成立の判定ではありません" in _all_text(guide)


def test_save_permission_exception_changes_save_location_not_work_conditions(tmp_path):
    exc = PermissionError(errno.EACCES, "Permission denied", str(tmp_path / "result.xlsx"))
    issue = issues_for_exception(exc, operation="save")[0]
    guide = guidance_for_issue(issue, operation="save")
    assert issue.message == str(exc)
    assert (guide.target, guide.category) == ("save", "save_path")
    assert "result.xlsx" in guide.location
    assert "人数・日数は変更しません" in guide.next_action
    assert "権限" in _all_text(guide)


def test_actual_same_input_folder_guard_has_save_only_remedy(tmp_path):
    path = export_template(tmp_path)
    data = read_input(path)
    result = solve_schedule(data)
    with pytest.raises(ValueError) as caught:
        export_result(data, result.schedule, result, tmp_path, input_path=path)
    issue = issues_for_exception(caught.value, operation="save")[0]
    guide = guidance_for_issue(issue, operation="save")
    assert guide.target == "save"
    assert "原本と別" in _all_text(guide)
    assert "読み直" not in _all_text(guide)


@pytest.mark.parametrize("operation,target", [
    ("mapping_load", "input"), ("mapping_save", "input"),
    ("settings_load", "settings"), ("settings_save", "settings"),
])
def test_json_operations_never_direct_user_to_reload_or_modify_input_excel(operation, target):
    exc = ValueError("JSON parse error\nfull detail")
    issue = issues_for_exception(exc, operation=operation)[0]
    guide = guidance_for_issue(issue, operation=operation)
    assert (guide.target, guide.category) == (target, operation)
    assert guide.reason == str(exc)
    assert "JSON" in guide.title and "JSON" in _all_text(guide)
    assert "入力Excel" not in guide.next_action
    assert "再実行" not in guide.next_action
    assert "保存先" in guide.action_label if operation.endswith("save") else "選び直す" in guide.action_label


def test_unknown_input_and_technical_errors_keep_reason_without_guessing():
    reason = "未対応の事情\n詳細は失わない"
    input_issue = Issue("入力不備", reason, Source("input.xlsx", "職員", 3))
    guide = guidance_for_issue(input_issue)
    assert guide.target == "input" and guide.reason == reason
    technical = guidance_for_issue(Issue("未知の例外", reason))
    assert technical.target == "support" and technical.reason == reason
    assert "特定できません" in _all_text(technical)
    assert "元行" not in technical.location
    with pytest.raises(FrozenInstanceError):
        technical.target = "settings"
    error = InputError([input_issue])
    assert issues_for_exception(error) == (input_issue,)


def test_template_save_is_distinct_from_result_and_json_save():
    exc = PermissionError(errno.EACCES, "Permission denied", "sample_input.xlsx")
    issue = issues_for_exception(exc, operation="template_save")[0]
    guide = guidance_for_issue(issue, operation="template_save", input_path="selected_original.xlsx")
    assert (guide.target, guide.category) == ("save", "template_save")
    assert guide.action_label == "入力ひな形の保存先を選び直す"
    assert "入力ひな形を保存" in guide.next_action
    assert "結果を保存" not in _all_text(guide)
    assert "selected_original.xlsx" not in guide.location
    assert guide.reason == str(exc)


def test_json_error_does_not_display_unrelated_selected_input_location():
    issue = issues_for_exception(ValueError("broken json"), operation="mapping_load")[0]
    guide = guidance_for_issue(issue, operation="mapping_load", input_path="unrelated.xlsx")
    assert "unrelated.xlsx" not in guide.location


def test_actual_invalid_individual_public_days_routes_to_employee_excel_row():
    data = sample_data()
    data = replace(data, employees=(replace(data.employees[0], public_holidays=-1,
        source=Source("input.xlsx", "職員", 2)), *data.employees[1:]),
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, default_public_holidays=1))
    issue = validate_input(data)[0]
    guide = guidance_for_issue(issue, data=data)
    assert (guide.target, guide.category) == ("input", "leave_counts")
    assert "元行 2" in guide.location and "読み直" in guide.next_action


def test_actual_bad_date_value_in_issue_does_not_crash_location_formatter():
    data = replace(sample_data(), fixed_off=(FixedOff("UNKNOWN", "2026-11-07"),))
    issues = validate_input(data)
    assert issues
    for issue in issues:
        guide = guidance_for_issue(issue, data=data)
        assert guide.reason == issue.message
        assert "2026-11-07" in guide.location


@pytest.mark.parametrize("operation,category,target", [
    ("mapping", "columns", "input"), ("settings", "settings", "settings"),
])
def test_generic_gui_value_error_goes_back_to_the_operation(operation, category, target):
    issue = issues_for_exception(ValueError("入力欄の整数が空欄です"), operation=operation)[0]
    guide = guidance_for_issue(issue, operation=operation)
    assert (guide.target, guide.category) == (target, category)
    assert "別名" not in _all_text(guide)
    assert "読み直" not in guide.next_action


@pytest.mark.parametrize("entered", ["abc", "0", "301"])
def test_invalid_gui_solve_time_limit_returns_to_run_with_input_correction(entered):
    # 実start_solveの入力検証だけを動かす。Tkルートやworkerは不要。
    from types import SimpleNamespace
    from sinta_shift.gui import ShiftApp
    errors = []
    app = SimpleNamespace(busy=False, data=sample_data(),
        time_var=SimpleNamespace(get=lambda: entered),
        _show_exception=lambda error, **context: errors.append((error, context)))
    ShiftApp.start_solve(app)
    assert len(errors) == 1
    error, context = errors[0]
    assert context == {"preserve_result": True, "operation": "solve"}
    issue = issues_for_exception(error, operation=context["operation"])[0]
    guide = guidance_for_issue(issue, operation=context["operation"])
    assert (guide.target, guide.category) == ("run", "time_limit")
    assert guide.reason == str(error)
    assert "0より大きく300以下" in _all_text(guide)
    assert "初期値は30秒" in _all_text(guide)
    assert "計算を開始できていません" in _all_text(guide)
    assert "時間内に案を発見" not in guide.title
    assert "案がない" not in _all_text(guide)
    assert "シフト案を作る" in guide.next_action


def test_solve_operation_does_not_reclassify_real_timeout_as_invalid_input():
    result = solve_schedule(sample_data(), time_limit=1e-12)
    issue = issues_for_solve_result(result)[0]
    guide = guidance_for_issue(issue, operation="solve")
    assert guide.category == "timeout"
    assert "計算を開始できていません" not in _all_text(guide)
