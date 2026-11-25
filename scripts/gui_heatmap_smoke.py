"""実Tkで希望集中・代替候補・不成立・手修正・別期間と表示を検証する。"""
from __future__ import annotations

import argparse
from dataclasses import replace
from datetime import date, timedelta
import hashlib
import json
from pathlib import Path
import platform
import sys
import time
import tkinter as tk
from tkinter import font as tkfont, ttk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sinta_shift.gui import ShiftApp
from sinta_shift.excel_io import _write_input_template
from sinta_shift.models import (
    DayRequirement, Employee, FixedOff, IndependentLeaveRequest, InputData,
    LeaveCandidate, PUBLIC_DAYS, PUBLIC_OFF, PUBLIC_REQUEST, ShiftSettings,
)
from sinta_shift.sample import sample_data
from sinta_shift.runtime import prepare_tk_runtime


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
    x = widget.winfo_rootx() - root.winfo_rootx()
    y = widget.winfo_rooty() - root.winfo_rooty()
    width, height = widget.winfo_width(), widget.winfo_height()
    inside = bool(widget.winfo_ismapped()) and x >= 0 and y >= 0 and width > 0 and height > 0 and x + width <= root.winfo_width() and y + height <= root.winfo_height()
    return {"inside_root": inside, "bounds": [x, y, width, height]}


def summaries(app):
    return {item.day: item for item in app.request_heatmap.summaries}


def load(root, app, path):
    app.load_file(path)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.base_data is not None, app.status_var.get()


def solve(root, app):
    assert app.data is not None, app.status_var.get()
    app.start_solve()
    idle(root, app)
    assert app.schedule is not None, app.status_var.get()
    assert app.evaluation.valid
    assert app.current_result_view == "requests"


def cluster_data():
    days = tuple(date(2026, 11, 1) + timedelta(days=i) for i in range(30))
    people = tuple(Employee(f"{i+1:03}", f"架空職員{i+1:02}", None, previous_shift=PUBLIC_OFF, public_holidays=10) for i in range(30))
    requests, metadata = [], []
    for i, employee in enumerate(people):
        for day in (date(2026, 11, 7), date(2026, 11, 8 + i % 10), date(2026, 11, 18 + i % 10)):
            rid = f"W:{employee.id}:{day.isoformat()}"
            requests.append(LeaveCandidate(rid, employee.id, day, 1))
            metadata.append(IndependentLeaveRequest(rid, employee.id, day, PUBLIC_REQUEST))
    return InputData(people, tuple(DayRequirement(day, 20, shift_counts=(("日勤", 20),)) for day in days),
        requests=tuple(requests), independent_requests=tuple(metadata), request_format="independent",
        settings=ShiftSettings(leave_enabled=True, day_count_mode=PUBLIC_DAYS, default_public_holidays=10))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/heatmap_20261123/gui"))
    parser.add_argument("--hold", type=float, default=0)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    inputs = args.output_dir / "inputs"
    inputs.mkdir(exist_ok=True)
    source = _write_input_template(inputs, cluster_data(), True, "independent")
    before_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    prepare_tk_runtime()
    root = tk.Tk()
    callback_errors = []
    root.report_callback_exception = lambda kind, value, traceback: callback_errors.append(f"{kind.__name__}: {value}")
    app = ShiftApp(root)
    root.update()
    load(root, app, source)
    app.go_step(4)
    app.show_result_view("requests")
    root.update()
    day = date(2026, 11, 7)
    before = summaries(app)[day]
    assert before.request_employee_count == 30
    assert before.adopted_employee_count is None
    solve(root, app)
    current = summaries(app)[day]
    assert current.request_employee_count == 30
    assert current.adopted_employee_count == 10
    assert current.unadopted_request_count == 20
    assert app.evaluation.adopted_count == 70
    app.request_heatmap.select_day(day)
    root.update()
    assert len(app.request_heatmap.detail_tree.get_children()) == 30
    assert "30" in app.request_heatmap.day_buttons[day].cget("text")
    assert "20" in app.request_heatmap.day_buttons[day].cget("text")
    canvas = app.request_heatmap.calendar_panel.canvas
    # 生成直後の初期画面では、6週ある月でも全日を見渡せる。
    visible_days = [button for value, button in app.request_heatmap.day_buttons.items()
                    if value.month == app.request_heatmap.month.month]
    all_visible = all(canvas.winfo_rooty() <= button.winfo_rooty() and
               button.winfo_rooty() + button.winfo_height() <= canvas.winfo_rooty() + canvas.winfo_height()
               for button in visible_days)
    assert all_visible, {"message": "初期画面で月全体が見えていません。",
        "canvas_height": canvas.winfo_height(), "calendar_body_height": app.request_heatmap.calendar_panel.body.winfo_height(),
        "top": min(b.winfo_rooty() for b in visible_days)-canvas.winfo_rooty(),
        "bottom": max(b.winfo_rooty()+b.winfo_height() for b in visible_days)-canvas.winfo_rooty(),
        "heatmap_bounds": geometry(root, app.request_heatmap)}
    style = ttk.Style(root)
    for button in visible_days:
        used_font = tkfont.Font(root=root, font=style.lookup(button.cget("style"), "font"))
        assert max(used_font.measure(line) for line in button.cget("text").splitlines()) <= button.winfo_width() - 8, "日付セルの文字が横に切れています。"
    outcome_box = app.request_heatmap.detail_tree.bbox("entry0", "採否")
    assert outcome_box and outcome_box[0] + outcome_box[2] <= app.request_heatmap.detail_tree.winfo_width(), "氏名・区分・採否が初期表示幅に収まっていません。"
    # 同じ入力へ戻した際にも濃さの基準を保持する。
    color_style = app.request_heatmap.day_buttons[day].cget("style")
    root.geometry("1120x800")
    root.update()
    minimum_calendar = geometry(root, app.request_heatmap)
    assert minimum_calendar["inside_root"]
    assert geometry(root, app.request_heatmap.detail_tree)["inside_root"]
    # 作成済み案を手修正すると希望の採否も更新し、不正案は参考と表示する。
    chosen = next(e for e in app.data.employees if app.schedule[e.id, day] == PUBLIC_OFF)
    app.show_result_view("schedule")
    app.toggle_cell(chosen.id, day, "日勤")
    idle(root, app)
    assert not app.evaluation.valid
    changed = summaries(app)[day]
    assert changed.adopted_employee_count == 9
    assert changed.unadopted_request_count == 21 and changed.is_reference
    app.show_result_view("requests")
    app.request_heatmap.select_day(day)
    root.update()
    assert str(app.save_button.cget("state")) == "disabled"
    app.toggle_cell(chosen.id, day, PUBLIC_OFF)
    idle(root, app)
    assert app.evaluation.valid and summaries(app)[day].adopted_employee_count == 10
    outputs = (args.output_dir / "results").resolve()
    outputs.mkdir(exist_ok=True)
    app.save_result(outputs)
    idle(root, app)
    assert app.last_saved and app.last_saved.is_file()
    saved = app.last_saved
    assert hashlib.sha256(source.read_bytes()).hexdigest() == before_hash

    # 代替候補の第2希望採用は、第1候補日で未採用へ落とさない。
    legacy = _write_input_template(inputs, sample_data(), False, "alternatives")
    app.set_settings(ShiftSettings())
    load(root, app, legacy)
    solve(root, app)
    legacy_days = summaries(app)
    assert legacy_days[day].request_employee_count == 2
    assert legacy_days[day].adopted_employee_count == 1
    assert legacy_days[day].other_candidate_adopted_count == 1
    assert legacy_days[day].unadopted_request_count == 0
    app.request_heatmap.select_day(day)
    assert any("他候補で採用" in str(app.request_heatmap.detail_tree.item(iid, "values"))
               for iid in app.request_heatmap.detail_tree.get_children())

    # 計算できない入力でも集中だけ確認でき、採否は未計算とする。
    conflict = replace(sample_data(), fixed_off=(FixedOff("A", day), FixedOff("B", day)))
    conflict_path = _write_input_template(inputs, conflict, False, "alternatives")
    load(root, app, conflict_path)
    assert app.schedule is None and app.data is None
    app.go_step(4)
    app.show_result_view("requests")
    root.update()
    blocked = summaries(app)[day]
    assert blocked.request_employee_count == 2 and blocked.adopted_employee_count is None
    assert blocked.fixed_off_employee_count == 2
    assert str(app.save_button.cget("state")) == "disabled"

    # 年月が重ならない新入力で、旧月・職員・採用状態を持ち越さない。
    future = sample_data()
    future = replace(future,
        requirements=tuple(replace(r, day=r.day + timedelta(days=85)) for r in future.requirements),
        requests=tuple(replace(r, day=r.day + timedelta(days=85)) for r in future.requests))
    future_path = _write_input_template(inputs, future, False, "alternatives")
    load(root, app, future_path)
    solve(root, app)
    assert all(item.day.year == 2027 for item in app.request_heatmap.summaries)
    assert app.request_heatmap.month.year == 2027
    assert day not in app.request_heatmap.day_buttons
    assert app.request_heatmap.month.month == 1
    app.request_heatmap.move_month(1)
    root.update()
    assert app.request_heatmap.month == date(2027, 2, 1)
    assert date(2027, 2, 1) in app.request_heatmap.day_buttons
    app.request_heatmap.move_month(-1)
    root.update()
    assert app.request_heatmap.month == date(2027, 1, 1)

    # 列対応を変えた直後に、旧入力を再び最新集計として表示しない。
    app.go_step(1)
    app.import_details.show()
    app.role_notebook.select(app.role_notebook.tabs()[app._roles.index("希望休")])
    idle(root, app)
    app.mapping_vars["希望休"]["columns"]["候補日"].set("職員ID")
    app.mapping_boxes["希望休"]["columns"]["候補日"].event_generate("<<ComboboxSelected>>")
    root.update()
    assert app.base_data is None
    until = time.monotonic() + 0.4
    while time.monotonic() < until:
        root.update()
        time.sleep(0.02)
    assert not app.request_heatmap.summaries

    # 最後は30人の集中カレンダーを表示して目視できる状態にする。
    load(root, app, source)
    solve(root, app)
    app.request_heatmap.select_day(day)
    root.geometry("1220x850")
    root.update()
    assert app.request_heatmap.day_buttons[day].cget("style") == color_style
    assert not callback_errors, callback_errors
    report = {"python": sys.version, "os": platform.platform(), "tk": root.tk.call("info", "patchlevel"),
        "input": str(source.resolve()), "saved_excel": str(saved),
        "auto_open_heatmap": True, "cluster_day": day.isoformat(),
        "requested_people": 30, "adopted_people": 10, "unadopted_requests": 20,
        "total_adopted_requests": 70, "detail_rows": 30,
        "legacy_other_candidate_distinguished": True, "manual_edit_recomputed": True,
        "invalid_save_blocked": True, "infeasible_input_heatmap_visible": True,
        "new_period_clears_old_calendar": True, "month_navigation": True,
        "mapping_change_clears_old_input": True, "whole_month_visible_on_start": True,
        "essential_detail_columns_visible": True, "calendar_cell_text_fits": True,
        "minimum_calendar_geometry": minimum_calendar,
        "input_sha256": before_hash, "input_unchanged": True, "callback_errors": callback_errors,
        "visual_review": "実画面の目視確認は別途必要"}
    (args.output_dir / "gui_heatmap_smoke.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.hold:
        until = time.monotonic() + args.hold
        while time.monotonic() < until:
            root.update()
            time.sleep(0.04)
    app.close()


if __name__ == "__main__":
    main()
