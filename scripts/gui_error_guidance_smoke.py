"""実Tkでエラーの案内から修正・再検証・保存までを確認する。"""
from __future__ import annotations

import argparse
from datetime import date
import hashlib
import json
from pathlib import Path
import platform
import sys
import time
import tkinter as tk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from sinta_shift import __version__
from sinta_shift.excel_io import _write_input_template
from sinta_shift.gui import ShiftApp
from sinta_shift.models import ShiftSettings, SolverStatus
from sinta_shift.runtime import prepare_tk_runtime
from sinta_shift.sample import sample_data


def idle(root, app, timeout=45):
    started = time.monotonic()
    while time.monotonic() - started < timeout:
        root.update()
        if not app.busy:
            return
        time.sleep(0.02)
    raise AssertionError("画面処理が完了しませんでした。")


def bounds(root, widget):
    root.update_idletasks()
    x, y = widget.winfo_rootx() - root.winfo_rootx(), widget.winfo_rooty() - root.winfo_rooty()
    w, h = widget.winfo_width(), widget.winfo_height()
    inside = bool(widget.winfo_ismapped()) and min(x, y) >= 0 and min(w, h) > 0 and x + w <= root.winfo_width() and y + h <= root.winfo_height()
    return {"inside_root": inside, "bounds": [x, y, w, h]}


def select_issue(root, app, predicate):
    iid = next(key for key, issue in app._issue_by_iid.items() if predicate(issue))
    app.error_tree.selection_set(iid)
    app.error_tree.event_generate("<<TreeviewSelect>>")
    root.update()
    assert app.notebook.select() == str(app.errors_tab), "行選択だけで説明画面を離れています。"
    assert app.selected_error_guidance is not None
    return app.selected_error_guidance


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/errors_20261125/gui"))
    parser.add_argument("--hold", type=float, default=0)
    args = parser.parse_args()
    output = args.output_dir.resolve()
    inputs = output / "inputs"
    inputs.mkdir(parents=True, exist_ok=True)
    (output / "results").mkdir(exist_ok=True)
    original = _write_input_template(inputs, sample_data(), False, "alternatives")
    wb = load_workbook(original)
    wb["職員"]["C2"] = "=1+1"
    formula = inputs / "formula.xlsx"
    wb.save(formula)
    wb.close()
    hashes = {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in (original, formula)}
    prepare_tk_runtime()
    root = tk.Tk()
    callbacks = []
    root.report_callback_exception = lambda kind, value, trace: callbacks.append(f"{kind.__name__}: {value}")
    app = ShiftApp(root)
    root.update()
    assert root.winfo_ismapped()
    app.load_file(formula)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is None and app.schedule is None
    guide = select_issue(root, app, lambda item: item.kind == "数式")
    assert "C2" in app.error_location_var.get()
    assert "職員" in app.error_location_var.get()
    assert guide.target == "input"
    assert "値" in app.error_steps_var.get() and "読" in app.error_next_var.get()
    assert str(app.save_button.cget("state")) == "disabled"
    formula_guidance = {"title": app.error_title_var.get(), "location": app.error_location_var.get(),
                        "steps": app.error_steps_var.get(), "next": app.error_next_var.get()}
    root.geometry("1120x800")
    root.update()
    layout = {name: bounds(root, getattr(app, name)) for name in
              ("error_tree", "error_detail_text", "error_raw_text", "error_action_button", "error_retry_button",
               "status_label", "error_alert_button", "workflow_button")}
    assert all(item["inside_root"] for item in layout.values()), layout
    assert int(app.status_label.cget("wraplength")) <= app.status_label.winfo_width()
    first_step = app.error_detail_text.search("1. ", "1.0", stopindex="end")
    line = app.error_detail_text.dlineinfo(first_step)
    assert line and line[1] + line[3] <= app.error_detail_text.winfo_height(), "直し方の冒頭が初期表示に見えません。"
    app.error_action()
    root.update()
    assert app.notebook.select() == str(app.import_tab)
    # 原本の数式は変更せず、値だけにした別コピーを読み直す。
    wb = load_workbook(formula)
    wb["職員"]["C2"] = 2
    corrected = inputs / "corrected_values.xlsx"
    wb.save(corrected)
    wb.close()
    hashes[str(corrected)] = hashlib.sha256(corrected.read_bytes()).hexdigest()
    # 同じ作業コピーを直した場合、再読込ボタンが設定シートも読み直す。
    working = inputs / "editable_working_copy.xlsx"
    wb = load_workbook(formula)
    wb.save(working)
    wb.close()
    app.load_file(working)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.selected_error_guidance.target == "input"
    wb = load_workbook(working)
    wb["職員"]["C2"] = 2
    settings_sheet = wb.create_sheet("勤務条件")
    settings_sheet.append(["項目", "値"])
    settings_sheet.append(["連続勤務上限", 2])
    wb.save(working)
    wb.close()
    app.retry_after_error()
    started = time.monotonic()
    while time.monotonic() - started < 30:
        root.update()
        if app.data is not None and not app.busy:
            break
        time.sleep(0.02)
    assert app.data is not None, app.status_var.get()
    assert app.data.settings.max_consecutive_days == 2
    app.set_settings(ShiftSettings())
    app.load_file(corrected)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is not None, app.status_var.get()
    app.start_solve()
    idle(root, app)
    assert app.evaluation.valid and app.evaluation.adopted_count == 3
    # 日別人数・個人日数の違反を実セル変更から検出し、案内から戻す。
    day = date(2026, 11, 7)
    app.toggle_cell("A", day, "日勤")
    idle(root, app)
    assert not app.evaluation.valid
    app.show_error_guidance()
    root.update()
    guide = select_issue(root, app, lambda item: item.employee_id == "A" and item.day == day)
    assert guide.target == "schedule"
    assert "corrected_values.xlsx" not in app.error_location_var.get()
    assert "Excelを修正" not in app.error_steps_var.get()
    manual_guidance = {"title": app.error_title_var.get(), "location": app.error_location_var.get(),
                       "steps": app.error_steps_var.get(), "next": app.error_next_var.get()}
    app.error_action()
    root.update()
    assert app.current_result_view == "schedule"
    assert app._selected_cell == ("emp0", day.isoformat())
    before = set(output.glob("results/*.xlsx"))
    app.save_result(output / "results")
    assert set(output.glob("results/*.xlsx")) == before
    app.toggle_cell("A", day, "休み")
    idle(root, app)
    assert app.evaluation.valid
    app.save_result(output / "results")
    idle(root, app)
    assert app.last_saved is not None and app.last_saved.is_file()
    saved = app.last_saved
    # 保存先の不備は入力Excelではなく保存先の説明へ。案は保持する。
    non_folder = output / "not_a_folder.txt"
    non_folder.write_text("save target fixture", encoding="utf-8")
    app.save_result(non_folder)
    idle(root, app)
    assert app.schedule is not None and app.evaluation.valid
    assert app.selected_error_guidance.target == "save"
    save_guidance = app.error_steps_var.get()
    # A・Bの土曜が確定休なら、必要2人・勤務可能1人を案内し出力を止める。
    wb = load_workbook(original)
    wb["確定休"].append(["A", day])
    wb["確定休"].append(["B", day])
    blocked = inputs / "fixed_off_conflict.xlsx"
    wb.save(blocked)
    wb.close()
    hashes[str(blocked)] = hashlib.sha256(blocked.read_bytes()).hexdigest()
    app.load_file(blocked)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is None and app.schedule is None
    guide = select_issue(root, app, lambda item: item.kind == "条件不成立" and item.day == day)
    assert "必要2人・勤務可能1人" in guide.reason
    conflict_guidance = {"reason": guide.reason, "steps": app.error_steps_var.get()}
    assert str(app.save_button.cget("state")) == "disabled"
    app.load_file(corrected)
    idle(root, app)
    app.read_conditions()
    idle(root, app)
    assert app.data is not None
    app.time_var.set("abc")
    app.start_solve()
    root.update()
    assert app.selected_error_guidance.category == "time_limit"
    assert app.data is not None and app.schedule is None
    app.error_action()
    root.update()
    assert app.notebook.select() == str(app.conditions_tab)
    # 実際の極短制限でUNKNOWNを発生させ、不成立と表示しない。
    app.time_var.set("0.000000001")
    app.start_solve()
    idle(root, app)
    assert app.result.status == SolverStatus.UNKNOWN and app.schedule is None
    assert app.error_tree.get_children(), "時間切れ時の説明が空です。"
    app.show_error_guidance()
    root.update()
    assert "時間" in app.error_title_var.get() or "時間" in app.error_reason_var.get()
    assert "存在しません" not in app.error_reason_var.get()
    timeout_guidance = {"reason": app.error_reason_var.get(), "steps": app.error_steps_var.get()}
    # 原因不明の処理エラーは推測せず、長い技術詳細も保持する。
    detail = "確認用の技術詳細：" + "長いエラーの原文／" * 300
    app._show_exception(RuntimeError(detail), preserve_result=True)
    root.update()
    assert detail in app.error_raw_text.get("1.0", "end-1c")
    assert "再インストール" not in app.error_steps_var.get()
    assert all(hashlib.sha256(Path(path).read_bytes()).hexdigest() == digest for path, digest in hashes.items())
    assert callbacks == [], callbacks
    report = {"version": __version__, "python": sys.version, "os": platform.platform(),
              "formula_input": str(formula), "corrected_input": str(corrected), "output": str(saved),
              "formula_guidance": formula_guidance, "manual_guidance": manual_guidance,
              "save_guidance": save_guidance, "timeout_guidance": timeout_guidance,
              "fixed_off_conflict_guidance": conflict_guidance,
              "minimum_window_layout": layout, "selection_stays_on_guidance": True,
              "first_fix_step_visible_without_scroll": True, "retry_reloads_excel_settings": True,
              "invalid_save_blocked": True, "long_details_retained": True,
              "originals_unchanged": True, "input_sha256": hashes, "callback_errors": callbacks}
    (output / "gui_error_guidance_smoke.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.hold:
        finish = time.monotonic() + args.hold
        while time.monotonic() < finish:
            root.update()
            time.sleep(0.02)
    app.close()


if __name__ == "__main__":
    main()
