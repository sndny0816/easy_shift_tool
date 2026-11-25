"""実Tkで公休・有給・明け、時刻確認、期間別公休日数を検証する。"""
from __future__ import annotations

import argparse
from dataclasses import replace
from datetime import date, timedelta
import json
from pathlib import Path
import platform
import sys
import time
import tkinter as tk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from sinta_shift.gui import ShiftApp, UNUSED_COLUMN
from sinta_shift.models import (DAY_ONLY, TWO_SHIFT, THREE_SHIFT, PUBLIC_DAYS, PUBLIC_OFF, PAID_LEAVE, RECOVERY,
                                ShiftSettings, ShiftTime, assignment_label)
from sinta_shift.excel_io import export_template, export_leave_sample
from sinta_shift.runtime import prepare_tk_runtime


def wait_for(root, predicate, timeout=45):
    started = time.monotonic()
    while time.monotonic() - started < timeout:
        root.update()
        if predicate():
            return
        time.sleep(0.02)
    raise AssertionError("GUI処理が制限時間内に完了しませんでした。")


def idle(root, app):
    wait_for(root, lambda: not app.busy)


def solve(root, app):
    app.start_solve()
    idle(root, app)
    assert app.schedule is not None, app.status_var.get()
    assert app.evaluation.valid, [i.message for i in app.evaluation.issues]
    app.show_result_view("schedule")
    root.update()


def crossing_fixture(folder):
    settings = ShiftSettings(mode=TWO_SHIFT, max_consecutive_days=5, leave_enabled=True,
        day_count_mode=PUBLIC_DAYS, default_public_holidays=1,
        shift_times=(ShiftTime("日勤", 480, 1020), ShiftTime("夜勤", 960, 1980)), shift_times_confirmed=True)
    path = export_template(folder, settings=settings)
    workbook = load_workbook(path)
    start = date(2026, 11, 7)
    content = {
        "職員": (["職員ID", "氏名", "対象期間の勤務日数", "前期間の連続勤務日数", "前日の勤務区分", "対象期間の公休日数"],
                 [["A", "職員A", None, 0, PUBLIC_OFF, 1], ["B", "職員B", None, 0, PUBLIC_OFF, 2]]),
        "必要人数": (["日付", "配置人数", "日勤人数", "夜勤人数"],
                     [[start + timedelta(days=i), sum(pair), *pair] for i, pair in enumerate(((0, 1), (1, 0), (0, 1), (1, 0), (0, 0)))]),
        "確定休": (["職員ID", "日付"], [["A", start + timedelta(days=2)]]),
        "希望休": (["申請ID", "職員ID", "候補日", "優先順位"],
                   [["R1", "A", start + timedelta(days=2), 1], ["R2", "A", start + timedelta(days=4), 1]]),
        "有給": (["職員ID", "日付"], [["A", start + timedelta(days=4)]]),
    }
    for name, (headers, rows) in content.items():
        sheet = workbook[name] if name in workbook.sheetnames else workbook.create_sheet(name)
        sheet.delete_rows(1, sheet.max_row)
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
    workbook.save(path)
    return path, settings


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/gui_leave_verification"))
    parser.add_argument("--hold", type=float, default=0)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    inputs = args.output_dir / "inputs"
    inputs.mkdir(exist_ok=True)
    prepare_tk_runtime()
    root = tk.Tk()
    app = ShiftApp(root)
    root.update()
    assert root.winfo_ismapped()
    assert not app.leave_enabled_var.get()
    app.load_leave_sample()
    wait_for(root, lambda: app.data is not None and not app.busy)
    assert app.data.settings.leave_enabled
    assert app.data.settings.day_count_mode == PUBLIC_DAYS
    assert "有給" in app.mapping_vars
    assert app.mapping_vars["有給"]["sheet"].get() == "有給"
    assert app.mapping_vars["職員"]["columns"]["対象期間の公休日数"].get() != UNUSED_COLUMN
    solve(root, app)
    assert sum(app.evaluation.employee_paid_counts.values()) == len(app.data.paid_leaves)
    assert sum(app.evaluation.employee_recovery_counts.values()) == 0
    assert all(isinstance(value, str) for value in app.schedule.values())
    assert {"公休", "有給", "明け", "実勤務日数"} <= set(app.result_tree.cget("columns"))
    day_output = (args.output_dir / "day_only").resolve()
    day_output.mkdir(exist_ok=True)
    app.save_result(day_output)
    wait_for(root, lambda: app.last_saved is not None and not app.busy)
    day_saved = app.last_saved
    paid = app.data.paid_leaves[0]
    app.toggle_cell(paid.employee_id, paid.day, PUBLIC_OFF)
    idle(root, app)
    assert not app.evaluation.valid
    assert str(app.save_button.cget("state")) == "disabled"
    app.toggle_cell(paid.employee_id, paid.day, PAID_LEAVE)
    idle(root, app)
    assert app.evaluation.valid
    settings_path = args.output_dir / "leave_settings.json"
    app.save_settings(settings_path)
    idle(root, app)
    assert settings_path.is_file()
    saved_settings = app._get_settings()
    app.set_settings(ShiftSettings())
    app.load_settings(settings_path)
    idle(root, app)
    assert app._get_settings() == saved_settings
    fixture, settings = crossing_fixture(inputs)
    # ファイルが確認済みでも、有効区分の時刻が不足していれば記入例を確認済み扱いしない。
    partial_path = fixture.with_name("missing_shift_time.xlsx")
    partial = load_workbook(fixture)
    times_sheet = partial["勤務時刻"]
    for row in range(times_sheet.max_row, 1, -1):
        if times_sheet.cell(row, 1).value == "夜勤":
            times_sheet.delete_rows(row)
    partial.save(partial_path)
    app.load_file(partial_path)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is None
    assert not app.times_confirmed_var.get()
    assert app.time_start_vars["夜勤"].get() == "16:00"
    assert app.time_end_vars["夜勤"].get() == "09:00"
    app.read_conditions()
    idle(root, app)
    assert app.data is None
    assert not app.times_confirmed_var.get()
    assert str(app.save_button.cget("state")) == "disabled"
    assert any("確認" in issue.message for issue in app._issue_by_iid.values())
    app.load_file(fixture)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is not None, app.status_var.get()
    assert app.data.settings.shift_times_confirmed
    assert app.data.settings.mode == TWO_SHIFT
    assert app._get_settings().shift_times == settings.shift_times
    # 時刻確認を外すと処理を止め、元ファイルの確認済み値へ黙って戻さない。
    app.times_confirmed_var.set(False)
    app.read_conditions()
    idle(root, app)
    assert app.data is None
    assert str(app.save_button.cget("state")) == "disabled"
    assert any("確認" in issue.message for issue in app._issue_by_iid.values())
    app.times_confirmed_var.set(True)
    app.read_conditions()
    idle(root, app)
    assert app.data is not None, app.status_var.get()
    solve(root, app)
    assert sum(app.evaluation.employee_recovery_counts.values()) == 2
    assert sum(app.evaluation.employee_paid_counts.values()) == 1
    assert app.evaluation.employee_public_counts == {"A": 1, "B": 2}
    assert all(employee.work_days is None for employee in app.data.employees)
    assert app.evaluation.employee_actual_work_days == {key: app.evaluation.employee_counts[key] + app.evaluation.employee_recovery_counts[key]
                                                        for key in app.evaluation.employee_counts}
    assert next(r for r in app.evaluation.request_results if r.request_id == "R2").adopted_day is None
    original = dict(app.schedule)
    recovery_key = next(key for key, value in original.items() if value == RECOVERY)
    app.toggle_cell(*recovery_key, PUBLIC_OFF)
    idle(root, app)
    assert not app.evaluation.valid
    assert str(app.save_button.cget("state")) == "disabled"
    reasons = [i.message for i in app.evaluation.issues]
    app.toggle_cell(*recovery_key, RECOVERY)
    idle(root, app)
    assert app.evaluation.valid
    crossing_output = (args.output_dir / "two_shift").resolve()
    crossing_output.mkdir(exist_ok=True)
    app.save_result(crossing_output)
    wait_for(root, lambda: app.last_saved is not None and not app.busy)
    crossing_saved = app.last_saved
    crossing_counts = {"public_counts": dict(app.evaluation.employee_public_counts),
                       "paid_counts": dict(app.evaluation.employee_paid_counts),
                       "recovery_counts": dict(app.evaluation.employee_recovery_counts),
                       "actual_work_days": dict(app.evaluation.employee_actual_work_days)}
    # 高度３交代の30日実ファイルでも、記入例を未確認のまま処理しない。
    three_settings = ShiftSettings(mode=THREE_SHIFT, max_consecutive_days=5, leave_enabled=True,
        day_count_mode=PUBLIC_DAYS, default_public_holidays=8,
        forbidden_transitions=(("準夜勤", "日勤"),))
    three_fixture = export_leave_sample(inputs, three_settings)
    app.load_file(three_fixture)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is None
    assert app.mode_var.get() == THREE_SHIFT
    assert not app.times_confirmed_var.get()
    assert any("確認" in issue.message for issue in app._issue_by_iid.values())
    assert str(app.save_button.cget("state")) == "disabled"
    app.times_confirmed_var.set(True)
    assert app._get_settings().shift_times_confirmed
    app.read_conditions()
    idle(root, app)
    assert app.data is not None, app.status_var.get()
    assert len(app.data.employees) == 6 and len(app.data.days) == 30
    solve(root, app)
    assert all(count == 8 for count in app.evaluation.employee_public_counts.values())
    assert all(count == 2 for count in app.evaluation.employee_paid_counts.values())
    assert all(count == 0 for count in app.evaluation.employee_recovery_counts.values())
    assert all(count == 20 for count in app.evaluation.employee_counts.values())
    three_output = (args.output_dir / "three_shift").resolve()
    three_output.mkdir(exist_ok=True)
    app.save_result(three_output)
    wait_for(root, lambda: app.last_saved is not None and not app.busy)
    three_saved = app.last_saved
    assert three_saved.is_file()
    three_counts = {"public_counts": dict(app.evaluation.employee_public_counts),
                    "paid_counts": dict(app.evaluation.employee_paid_counts),
                    "recovery_counts": dict(app.evaluation.employee_recovery_counts)}
    app.notebook.select(app.result_tab)
    root.update()
    button = app.save_button
    save_x = button.winfo_rootx() - root.winfo_rootx()
    save_y = button.winfo_rooty() - root.winfo_rooty()
    save_width, save_height = button.winfo_width(), button.winfo_height()
    save_mapped = bool(button.winfo_ismapped())
    save_inside = (save_mapped and save_width > 0 and save_height > 0 and save_x >= 0 and save_y >= 0
                   and save_x + save_width <= root.winfo_width()
                   and save_y + save_height <= root.winfo_height())
    assert save_mapped and save_inside, "高度３交代の結果保存ボタンが画面内に表示されていません。"
    save_geometry = {"mapped": save_mapped, "inside_root": save_inside,
                     "bounds": [save_x, save_y, save_width, save_height]}
    report = {"python": sys.version, "os": platform.platform(), "tk": root.tk.call("info", "patchlevel"),
              "window": [root.winfo_width(), root.winfo_height()], "day_sample_saved": str(day_saved),
              "crossing_sample_saved": str(crossing_saved), "three_shift_sample_saved": str(three_saved),
              "three_shift_unconfirmed_blocked": True, "three_shift_counts": three_counts,
              "save_button_geometry": save_geometry, "paid_public_separated": True,
              "paid_cannot_become_public": True, "recovery_cannot_become_public": True,
              "unconfirmed_shift_times_blocked": True, "incomplete_times_not_auto_confirmed": True,
              "settings_json_round_trip": True,
              **crossing_counts, "edit_violations": reasons,
              "result_summary": app.result_summary.get(), "visual_review": "実画面の目視確認は別途必要"}
    (args.output_dir / "gui_leave_smoke.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.hold:
        deadline = time.monotonic() + args.hold
        while time.monotonic() < deadline:
            root.update()
            time.sleep(0.04)
    app.close()


if __name__ == "__main__":
    main()
