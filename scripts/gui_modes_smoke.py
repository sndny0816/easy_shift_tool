"""実Tkで勤務3モード、曜日割合、連勤・手修正・条件保存を通す追加検証。"""
from __future__ import annotations

import argparse
from datetime import date, timedelta
from decimal import Decimal
import json
from pathlib import Path
import platform
import sys
import time
import tkinter as tk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
from sinta_shift.gui import ShiftApp
from sinta_shift.models import DAY_ONLY, TWO_SHIFT, THREE_SHIFT, SHIFT_MODES, OFF, ShiftSettings, SolverStatus, assignment_label
from sinta_shift.excel_io import REQUIRED_FIELDS
from sinta_shift.runtime import prepare_tk_runtime


def wait_for(root, predicate, timeout=45):
    started = time.monotonic()
    while time.monotonic() - started < timeout:
        root.update()
        if predicate():
            return
        time.sleep(0.02)
    raise AssertionError("GUI操作が制限時間内に完了しませんでした。")


def idle(root, app):
    wait_for(root, lambda: not app.busy)


def make_input(path):
    """6人・4日・1人2日勤務。計12人日を毎日3人へ配置する架空データ。"""
    workbook = Workbook()
    workbook.remove(workbook.active)
    start = date(2026, 11, 7)
    rows = {
        "職員": [(f"{i:03}", f"職員{i}", 2) for i in range(1, 7)],
        "必要人数": [(start + timedelta(days=i), 3) for i in range(4)],
        "確定休": [],
        "希望休": [("R1", "001", start, 1), ("R2", "002", start + timedelta(days=1), 1)],
    }
    for role, fields in REQUIRED_FIELDS.items():
        sheet = workbook.create_sheet(role)
        sheet.append(list(fields))
        for row in rows[role]:
            sheet.append(row)
    workbook.save(path)


def make_renamed_history_input(path):
    """改名された任意履歴列をGUIで対応すると初日夜勤後休みと配置が衝突する。"""
    workbook = Workbook()
    workbook.remove(workbook.active)
    for role, fields in REQUIRED_FIELDS.items():
        sheet = workbook.create_sheet(role)
        sheet.append([*fields, *(["直前の連勤日数", "直前の区分"] if role == "職員" else [])])
        if role == "職員":
            sheet.append(["001", "職員1", 1, 1, "夜勤"])
        elif role == "必要人数":
            sheet.append([date(2026, 11, 7), 1])
    workbook.save(path)


def read_and_solve(root, app):
    app.read_conditions()
    idle(root, app)
    assert app.data is not None, app.status_var.get()
    app.start_solve()
    wait_for(root, lambda: app.schedule is not None and not app.busy)
    assert app.evaluation.valid, app.status_var.get()
    app.show_result_view("schedule")
    root.update()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/gui_modes_verification"))
    parser.add_argument("--hold", type=float, default=0)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    input_dir = args.output_dir / "inputs"
    input_dir.mkdir(exist_ok=True)
    fixture = (input_dir / "six_staff.xlsx").resolve()
    make_input(fixture)
    prepare_tk_runtime()
    root = tk.Tk()
    app = ShiftApp(root)
    root.update()
    assert root.winfo_ismapped()
    assert app.max_enabled_var.get() and app.max_days_var.get() == "5"
    assert app.night_rest_var.get()
    app.mode_var.set(THREE_SHIFT)
    assert ("準夜勤", "日勤") in app.forbidden_transitions
    app.mode_var.set(DAY_ONLY)
    records = []
    # 旧受入サンプルを、曜日割合と連勤上限2日で実ファイルから計算する。
    app.set_settings(ShiftSettings(DAY_ONLY, 2, (Decimal(50),) * 7, (1,)))
    app.load_sample()
    wait_for(root, lambda: app.data is not None and not app.busy)
    app.start_solve()
    wait_for(root, lambda: app.schedule is not None and not app.busy)
    assert app.evaluation.valid
    assert app.evaluation.adopted_count == 3
    assert app.evaluation.priority_sum == 4
    assert app.data.settings.max_consecutive_days == 2
    assert all(r.required == 2 for r in app.data.requirements)
    settings_path = args.output_dir / "勤務条件.json"
    app.save_settings(settings_path)
    idle(root, app)
    assert settings_path.is_file()
    app.set_settings(ShiftSettings(DAY_ONLY, 1))
    assert app.schedule is None and app.data is None
    app.load_settings(settings_path)
    idle(root, app)
    assert app._get_settings().max_consecutive_days == 2
    assert app._get_settings().weekday_ratios == (Decimal(50),) * 7
    read_and_solve(root, app)
    day_out = (args.output_dir / "day_only").resolve()
    day_out.mkdir(exist_ok=True)
    app.save_result(day_out)
    wait_for(root, lambda: app.last_saved is not None and not app.busy)
    records.append({"mode": DAY_ONLY, "saved": str(app.last_saved), "adopted": app.evaluation.adopted_count,
                    "ratio_round_up": "3人×50％→2人", "settings_round_trip": True})
    for mode in (TWO_SHIFT, THREE_SHIFT):
        forbidden = (("準夜勤", "日勤"),) if mode == THREE_SHIFT else ()
        app.set_settings(ShiftSettings(mode, 2, (Decimal(50),) * 7, (1,) * len(SHIFT_MODES[mode]), True, forbidden))
        app.load_file(fixture)
        idle(root, app)
        read_and_solve(root, app)
        assert app.data.settings.mode == mode
        assert all(r.required == 3 for r in app.data.requirements)
        assert all(e.previous_work_days == 0 for e in app.data.employees)
        assert all(set(counts) == set(SHIFT_MODES[mode]) for counts in app.evaluation.daily_shift_counts.values())
        assert len(app.result_tree.get_children()) == len(app.data.employees) + 1 + len(SHIFT_MODES[mode])
        for req in app.data.requirements:
            assert app.evaluation.daily_shift_counts[req.day] == dict(req.shift_counts)
        original = dict(app.schedule)
        # 勤務人数は保ちながら区分だけ変更しても、区分人数違反を検出する。
        key = next(key for key, value in original.items() if assignment_label(value, mode) != OFF)
        old_label = assignment_label(original[key], mode)
        other_label = next(label for label in SHIFT_MODES[mode] if label != old_label)
        app.toggle_cell(*key, other_label)
        idle(root, app)
        assert not app.evaluation.valid
        assert any(i.kind == "勤務区分人数違反" for i in app.evaluation.issues)
        assert str(app.save_button.cget("state")) == "disabled"
        app.toggle_cell(*key, old_label)
        idle(root, app)
        assert app.evaluation.valid
        # 夜勤（3交代は深夜勤）の翌日を勤務へ変更すると、翌日休み違反を検出する。
        night_label = "夜勤" if mode == TWO_SHIFT else "深夜勤"
        night_key = next(key for key, value in original.items()
                         if value == night_label and key[1] != app.data.days[-1])
        following = app.data.days[app.data.days.index(night_key[1]) + 1]
        app.toggle_cell(night_key[0], following, "日勤")
        idle(root, app)
        assert any(i.kind == "夜勤翌日違反" for i in app.evaluation.issues)
        assert str(app.save_button.cget("state")) == "disabled"
        app.toggle_cell(night_key[0], following, assignment_label(original[(night_key[0], following)], mode))
        idle(root, app)
        assert app.evaluation.valid
        if mode == THREE_SHIFT:
            evening_key = next(key for key, value in original.items()
                               if value == "準夜勤" and key[1] != app.data.days[-1])
            following = app.data.days[app.data.days.index(evening_key[1]) + 1]
            app.toggle_cell(evening_key[0], following, "日勤")
            idle(root, app)
            assert any(i.kind == "勤務順序違反" for i in app.evaluation.issues)
            assert str(app.save_button.cget("state")) == "disabled"
            app.toggle_cell(evening_key[0], following, assignment_label(original[(evening_key[0], following)], mode))
            idle(root, app)
            assert app.evaluation.valid
        # 3日連続の勤務へ手修正し、上限2日違反を検出・保存停止する。
        employee = app.data.employees[0]
        for day in app.data.days[:3]:
            app.toggle_cell(employee.id, day, SHIFT_MODES[mode][0])
            idle(root, app)
        assert any(i.kind == "連続勤務違反" for i in app.evaluation.issues)
        assert str(app.save_button.cget("state")) == "disabled"
        limit_reasons = [i.message for i in app.evaluation.issues if i.kind == "連続勤務違反"]
        for day in app.data.days[:3]:
            app.toggle_cell(employee.id, day, assignment_label(original[(employee.id, day)], mode))
            idle(root, app)
        assert app.evaluation.valid
        assert "必須条件：すべて適合" in app.result_summary.get()
        mode_out = (args.output_dir / ("two_shift" if mode == TWO_SHIFT else "three_shift")).resolve()
        mode_out.mkdir(exist_ok=True)
        app.save_result(mode_out)
        wait_for(root, lambda: app.last_saved is not None and not app.busy)
        saved = app.last_saved
        assert saved.is_file()
        records.append({"mode": mode, "saved": str(saved), "shift_counts": dict(app.data.requirements[0].shift_counts),
                        "shift_edit_violation_blocked": True, "night_rest_violation_blocked": True,
                        "forbidden_transition_violation_blocked": mode == THREE_SHIFT, "consecutive_edit_violation_blocked": True,
                        "consecutive_reasons": limit_reasons})
        # 人数だけを増やした場合、各人の勤務日数を自動補正せず再確認で停止する。
        app.set_settings(ShiftSettings(mode, 2, (Decimal(67),) * 7, (1,) * len(SHIFT_MODES[mode])))
        assert app.schedule is None
        app.read_conditions()
        idle(root, app)
        assert app.data is None
        assert app.error_tree.get_children()
        assert any("合計" in issue.message for issue in app._issue_by_iid.values())
        assert str(app.save_button.cget("state")) == "disabled"
    # 改名された任意列もGUI対応を通じて読み、期間境界の違反を見落とさない。
    renamed_history = (input_dir / "renamed_history.xlsx").resolve()
    make_renamed_history_input(renamed_history)
    app.set_settings(ShiftSettings(DAY_ONLY, 5, night_rest=True))
    app.load_file(renamed_history)
    idle(root, app)
    assert str(app.night_rest_check.cget("state")) == "normal"
    app.mapping_vars["職員"]["columns"]["前期間の連続勤務日数"].set("直前の連勤日数")
    app.mapping_vars["職員"]["columns"]["前日の勤務区分"].set("直前の区分")
    mapped_history = app._get_mappings()["職員"]["columns"]
    assert mapped_history["前期間の連続勤務日数"] == "直前の連勤日数"
    assert mapped_history["前日の勤務区分"] == "直前の区分"
    app.read_conditions()
    idle(root, app)
    assert app.base_data is not None, app.status_var.get()
    assert app.base_data.employees[0].previous_work_days == 1
    assert app.base_data.employees[0].previous_shift == "夜勤"
    if app.data is not None:
        app.start_solve()
        idle(root, app)
        assert app.result.status == SolverStatus.INFEASIBLE
    assert app.schedule is None
    assert str(app.save_button.cget("state")) == "disabled"
    # 最後は正常な3交代案を表示して終了する。
    app.load_file(fixture)
    idle(root, app)
    app.set_settings(ShiftSettings(THREE_SHIFT, 2, (Decimal(50),) * 7, (1, 1, 1), True, (("準夜勤", "日勤"),)))
    read_and_solve(root, app)
    root.update()
    report = {"python": sys.version, "os": platform.platform(), "tk": root.tk.call("info", "patchlevel"),
              "mapped": bool(root.winfo_ismapped()), "window": [root.winfo_width(), root.winfo_height()],
              "modes": records, "ratio_mismatch_not_autocorrected": True,
              "renamed_optional_history_mapped": True, "mapped_night_history_infeasible": True,
              "result_summary": app.result_summary.get(),
              "visual_review": "表示の目視確認は別途必要"}
    report_path = args.output_dir / "gui_modes_smoke.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.hold:
        deadline = time.monotonic() + args.hold
        while time.monotonic() < deadline:
            root.update()
            time.sleep(0.04)
    app.close()


if __name__ == "__main__":
    main()
