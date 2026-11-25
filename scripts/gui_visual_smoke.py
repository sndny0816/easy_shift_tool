"""実Tkで4ステップ、週/日付出勤率、公休有給配分、週表示後のセル編集を検証。"""
from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta
from decimal import Decimal
import json
from pathlib import Path
import platform
import sys
import time
import tkinter as tk
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sinta_shift.gui import ShiftApp
from sinta_shift.models import PUBLIC_OFF, ShiftSettings, assignment_label
from sinta_shift.planning import attendance_ratio_for_day, attendance_source_for_day, week_start
from sinta_shift.runtime import prepare_tk_runtime
from sinta_shift.excel_io import export_template
from openpyxl import load_workbook


def wait_for(root, predicate, timeout=50):
    started = time.monotonic()
    while time.monotonic() - started < timeout:
        root.update()
        if predicate():
            return
        time.sleep(0.02)
    raise AssertionError("GUI処理が制限時間内に完了しませんでした。")


def idle(root, app):
    wait_for(root, lambda: not app.busy)


def geometry(root, widget):
    root.update_idletasks()
    x, y = widget.winfo_rootx() - root.winfo_rootx(), widget.winfo_rooty() - root.winfo_rooty()
    width, height = widget.winfo_width(), widget.winfo_height()
    mapped = bool(widget.winfo_ismapped())
    inside = mapped and x >= 0 and y >= 0 and width > 0 and height > 0 and x + width <= root.winfo_width() and y + height <= root.winfo_height()
    return {"mapped": mapped, "inside_root": inside, "bounds": [x, y, width, height]}


def solve(root, app):
    app.start_solve()
    idle(root, app)
    assert app.schedule is not None, app.status_var.get()
    assert app.evaluation.valid, [issue.message for issue in app.evaluation.issues]
    app.show_result_view("schedule")
    root.update()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/gui_visual_verification"))
    parser.add_argument("--hold", type=float, default=0)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    prepare_tk_runtime()
    root = tk.Tk()
    callback_errors = []
    root.report_callback_exception = lambda kind, value, traceback: callback_errors.append(f"{kind.__name__}: {value}")
    app = ShiftApp(root)
    root.update()
    assert root.winfo_ismapped()
    assert len(app.step_buttons) == 4
    assert not app.import_details.expanded
    assert all(geometry(root, button)["inside_root"] for button in app.step_buttons.values())
    app.load_independent_sample()
    wait_for(root, lambda: app.base_data is not None and not app.busy)
    assert len(app._paid_requests()) == 3
    assert not app.confirmed_paid_request_ids and not app.converted_public_request_ids
    app.go_step(2)
    root.update()
    target = app.base_data.days[3]
    monday = week_start(target)
    assert len(app.attendance_calendar.day_buttons) == 30
    assert attendance_ratio_for_day(target, app._get_settings()) is None
    app.set_attendance_ratio(monday, Decimal(50), scope="week")
    app.set_attendance_ratio(target, Decimal(80))
    assert attendance_ratio_for_day(target, app._get_settings()) == Decimal(80)
    assert attendance_source_for_day(target, app._get_settings()) == "日付指定"
    assert "3人" in app.attendance_calendar.day_buttons[target].cget("text")
    app.set_attendance_ratio(target, None)
    assert attendance_ratio_for_day(target, app._get_settings()) == Decimal(50)
    assert attendance_source_for_day(target, app._get_settings()) == "週指定"
    app.set_attendance_ratio(monday, None, scope="week")
    assert attendance_source_for_day(target, app._get_settings()) == "Excel人数"
    # 合計人数を変えない率でも、週/日付の優先順位とJSON往復を確認する。
    app.set_attendance_ratio(monday, Decimal(75), scope="week")
    app.set_attendance_ratio(monday, Decimal(25))
    def entry_values(widget):
        values = [widget.get()] if widget.winfo_class() == "TEntry" else []
        for child in widget.winfo_children():
            values.extend(entry_values(child))
        return values
    app._attendance_dialog(monday, "week")
    root.update()
    dialog = next(widget for widget in root.winfo_children() if isinstance(widget, tk.Toplevel))
    assert "75" in entry_values(dialog)
    dialog.destroy()
    app.set_attendance_ratio(monday, None)
    app.set_attendance_ratio(monday, Decimal(50), scope="week")
    app.set_attendance_ratio(target, Decimal("66.66"))
    app._attendance_dialog(target, "day")
    root.update()
    dialog = next(widget for widget in root.winfo_children() if isinstance(widget, tk.Toplevel))
    assert "66.66" in entry_values(dialog)
    dialog.destroy()
    app.set_attendance_ratio(target, Decimal(60))
    app.max_enabled_var.set(True)
    app._cap_scale_changed("5.1")
    assert app.max_days_var.get() == "5"
    assert "最大5日" in app.visual_cap_var.get()
    settings_path = args.output_dir / "visual_settings.json"
    app.save_settings(settings_path)
    idle(root, app)
    saved_settings = app._get_settings()
    app.clear_attendance_overrides()
    app.load_settings(settings_path)
    idle(root, app)
    assert app._get_settings().weekly_ratios == saved_settings.weekly_ratios
    assert app._get_settings().date_ratios == saved_settings.date_ratios
    app.read_conditions(show_settings=True)
    idle(root, app)
    calendar_geometry = geometry(root, app.attendance_calendar)
    assert calendar_geometry["inside_root"]
    app.go_step(3)
    root.update()
    original_paid = set(app.confirmed_paid_request_ids)
    original_public = set(app.converted_public_request_ids)
    app.mix_paid_count_var.set("1")
    root.update()
    assert app.confirmed_paid_request_ids == original_paid
    assert app.converted_public_request_ids == original_public
    assert len(app._planned_paid_ids) == 1 and len(app._planned_public_ids) == 2
    app.apply_paid_mix()
    idle(root, app)
    assert len(app.confirmed_paid_request_ids) == 1
    assert len(app.converted_public_request_ids) == 2
    assert app.data is not None, app.status_var.get()
    assert sum(employee.public_holiday_adjustment for employee in app.data.employees) == 2
    solve(root, app)
    assert sum(app.evaluation.employee_public_counts.values()) == 29
    assert sum(app.evaluation.employee_paid_counts.values()) == 1
    assert all(count == 20 for count in app.evaluation.employee_counts.values())
    results = (args.output_dir / "results").resolve()
    results.mkdir(exist_ok=True)
    app.save_result(results)
    wait_for(root, lambda: app.last_saved is not None and not app.busy)
    saved = app.last_saved
    # 2週目の表示位置から、実際の日付列へ正しく対応付けて編集する。
    app.go_step(4)
    app.show_result_week(1)
    root.update()
    display = tuple(app.result_tree.cget("displaycolumns"))
    assert len(display) <= 9
    column = next(value for value in display if value not in ("職員ID", "氏名"))
    day = date.fromisoformat(column)
    employee = app.data.employees[0]
    before = dict(app.schedule)
    bbox = app.result_tree.bbox("emp0", column)
    assert bbox, (display, app.result_week_var.get())
    x, y, width, height = bbox
    app._remember_cell(SimpleNamespace(x=x + width // 2, y=y + height // 2))
    assert app._selected_employee_day() == (employee.id, day)
    app._toggle_selected()
    idle(root, app)
    changed = {key for key in before if before[key] != app.schedule[key]}
    assert changed == {(employee.id, day)}, changed
    assert not app.evaluation.valid
    assert str(app.save_button.cget("state")) == "disabled"
    app.toggle_cell(employee.id, day, assignment_label(before[(employee.id, day)], app.data.settings))
    idle(root, app)
    assert app.evaluation.valid
    assert app.result_week_index == 1
    current = dict(app.schedule)
    app.show_result_week(2)
    app._toggle_selected()
    assert app.schedule == current
    assert app._selected_employee_day() is None
    app.show_result_week(1)
    # 配分の個別変更は既存固定有給へ触れず、保留に戻すと処理を止める。
    paid_id = next(iter(app.confirmed_paid_request_ids))
    app.set_leave_choice([paid_id], choice="pending")
    idle(root, app)
    assert paid_id not in app.confirmed_paid_request_ids
    assert app.data is None and app.schedule is None
    assert len(app.converted_public_request_ids) == 2
    app.set_leave_choice([paid_id], choice="paid")
    idle(root, app)
    assert app.data is not None
    solve(root, app)
    # 前の表示週と重ならない別期間の入力でも、古い表示列IDを持ち越さない。
    inputs = args.output_dir / "inputs"
    inputs.mkdir(exist_ok=True)
    future = export_template(inputs)
    workbook = load_workbook(future)
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (date, datetime)):
                    cell.value += timedelta(days=80)
    workbook.save(future)
    app.show_result_week(1)
    app.set_settings(ShiftSettings(max_consecutive_days=5))
    app.load_file(future)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is not None, app.status_var.get()
    solve(root, app)
    assert app.data.days[0].year == 2027
    assert all(column in ("職員ID", "氏名") or column.startswith("2027-") for column in app.result_tree.cget("displaycolumns"))
    # 元の月例へ戻し、確認済みでない配分をもう一度明示確定して表示する。
    app.load_independent_sample()
    wait_for(root, lambda: app.base_data is not None and not app.busy)
    app.mix_paid_count_var.set("1")
    app.apply_paid_mix()
    idle(root, app)
    solve(root, app)
    # 最小サイズでも主ボタンとカレンダーが画面内にある。
    root.geometry("1120x800")
    root.update()
    minimum_main_button = geometry(root, app.workflow_button)
    assert minimum_main_button["inside_root"]
    app.go_step(2)
    control_days = app.data.days
    app.set_attendance_ratio(week_start(control_days[1]), Decimal(50), scope="week")
    app.set_attendance_ratio(control_days[2], Decimal(60))
    app.saved_var.set("保存先：" + "/長い保存先のフォルダ" * 15 + "/sinta_shift_result.xlsx")
    root.update()
    minimum_calendar = geometry(root, app.attendance_calendar)
    assert minimum_calendar["inside_root"]
    app.settings_controls_panel.canvas.yview_moveto(1.0)
    root.update()
    minimum_clear_button = geometry(root, app.clear_attendance_button)
    assert minimum_clear_button["inside_root"]
    assert minimum_clear_button["bounds"][3] >= 28
    canvas = app.settings_controls_panel.canvas
    button_top = app.clear_attendance_button.winfo_rooty()
    assert canvas.winfo_rooty() <= button_top
    assert button_top + app.clear_attendance_button.winfo_height() <= canvas.winfo_rooty() + canvas.winfo_height()
    # 下までスクロールして到達した解除ボタンを実際に押す。
    app.clear_attendance_button.invoke()
    root.update()
    assert not app.weekly_ratio_overrides and not app.date_ratio_overrides
    assert app.schedule is None
    app.read_conditions(show_manager=True)
    idle(root, app)
    solve(root, app)
    app.saved_var.set("")
    app.go_step(4)
    app.show_result_week(1)
    root.geometry("1220x850")
    root.update()
    save_geometry = geometry(root, app.save_button)
    assert save_geometry["inside_root"]
    assert not callback_errors, callback_errors
    report = {"python": sys.version, "os": platform.platform(), "tk": root.tk.call("info", "patchlevel"),
        "saved_excel": str(saved), "four_step_navigation": True, "collapsed_mapping_on_start": True,
        "week_date_weekday_excel_priority": True, "override_json_round_trip": True,
        "week_editor_ignores_day_override": True, "fractional_ratio_not_rounded_on_open": True,
        "mix_preview_does_not_confirm": True, "mix_confirmed": {"public": 2, "paid": 1},
        "public_days_total": sum(app.evaluation.employee_public_counts.values()),
        "paid_days_total": sum(app.evaluation.employee_paid_counts.values()),
        "weekly_cell_edit_actual_date": day.isoformat(), "weekly_edit_changed_one_cell": True, "hidden_previous_week_cannot_be_edited": True,
        "nonoverlapping_period_replaces_displaycolumns": True,
        "calendar_geometry": calendar_geometry, "minimum_calendar_geometry": minimum_calendar,
        "minimum_main_button_geometry": minimum_main_button, "minimum_clear_button_geometry": minimum_clear_button,
        "long_saved_path_controls_scroll_reachable": True, "save_button_geometry": save_geometry,
        "callback_errors": callback_errors, "result_summary": app.result_summary.get(),
        "visual_review": "実画面の目視確認は別途必要"}
    (args.output_dir / "gui_visual_smoke.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.hold:
        deadline = time.monotonic() + args.hold
        while time.monotonic() < deadline:
            root.update()
            time.sleep(0.04)
    app.close()


if __name__ == "__main__":
    main()
