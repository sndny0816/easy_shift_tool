"""実Tkでサンプル・計算・変更・保存を通す検証。画面画像の目視確認は別途必要。"""
from __future__ import annotations
import argparse
from datetime import date
import json
from pathlib import Path
import platform
import sys
import time
import tkinter as tk

# このファイルをリポジトリ直下から直接実行できるようにする。
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sinta_shift.gui import ShiftApp
from sinta_shift.runtime import prepare_tk_runtime


def wait_for(root, predicate, timeout=45):
    start = time.monotonic()
    while time.monotonic() - start < timeout:
        root.update()
        if predicate():
            return
        time.sleep(0.02)
    raise AssertionError("GUI処理が制限時間内に完了しませんでした。")


def widget_bounds(root, widget):
    """rootのクライアント領域に対する表示位置と寸法を記録する。"""
    root.update_idletasks()
    x = widget.winfo_rootx() - root.winfo_rootx()
    y = widget.winfo_rooty() - root.winfo_rooty()
    width, height = widget.winfo_width(), widget.winfo_height()
    mapped = bool(widget.winfo_ismapped())
    inside = (mapped and width > 0 and height > 0 and x >= 0 and y >= 0
              and x + width <= root.winfo_width()
              and y + height <= root.winfo_height())
    return {"mapped": mapped, "bounds": [x, y, width, height], "inside_root": inside}


def row_bbox(tree, iid):
    bbox = tree.bbox(iid)
    return list(bbox) if bbox else []


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/gui_verification"))
    parser.add_argument("--hold", type=float, default=0,
                        help="最終結果画面を指定秒数表示（画像の目視確認用）")
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    prepare_tk_runtime()
    root = tk.Tk()
    app = ShiftApp(root)
    root.update()
    assert root.winfo_ismapped()
    assert app.notebook.winfo_width() > 0
    app.load_sample()
    wait_for(root, lambda: app.data is not None and not app.busy)
    assert app.preview_tree.get_children()
    assert len(app.data.employees) == 3
    app.notebook.select(app.import_tab)
    root.update()
    preview_geometry = widget_bounds(root, app.preview_tree)
    assert preview_geometry["mapped"], "読み込みプレビューが表示されていません。"
    assert preview_geometry["inside_root"], f"読み込みプレビューがウィンドウ外です: {preview_geometry}"
    app.start_solve()
    wait_for(root, lambda: app.schedule is not None and not app.busy)
    # 作成直後は希望休カレンダー。従来の勤務表検証は表へ切り替える。
    app.show_result_view("schedule")
    root.update()
    assert app.evaluation.valid
    assert app.evaluation.adopted_count == 3
    assert app.schedule[("A", date(2026, 11, 7))] is False
    assert app.schedule[("B", date(2026, 11, 8))] is False
    assert app.schedule[("C", date(2026, 11, 9))] is False
    assert app.result_tree.winfo_ismapped()
    assert len(app.request_tree.get_children()) == 3
    app.save_result(args.output_dir.resolve())
    wait_for(root, lambda: app.last_saved is not None and not app.busy)
    saved = app.last_saved
    assert saved.is_file()
    before_files = set(args.output_dir.glob("*.xlsx"))
    app.toggle_cell("A", date(2026, 11, 7))
    wait_for(root, lambda: not app.busy)
    assert not app.evaluation.valid
    assert str(app.save_button.cget("state")) == "disabled"
    assert any("⚠" in str(app.result_tree.item(iid, "values")) for iid in app.result_tree.get_children())
    assert app.error_tree.get_children()
    app.save_result(args.output_dir.resolve())
    assert set(args.output_dir.glob("*.xlsx")) == before_files
    violations = [issue.message for issue in app.evaluation.issues]
    app.toggle_cell("A", date(2026, 11, 7))
    wait_for(root, lambda: not app.busy)
    assert app.evaluation.valid
    assert app.evaluation.adopted_count == 3
    assert app.evaluation.priority_sum == 4
    assert all(count == 2 for count in app.evaluation.daily_counts.values())
    assert all(count == 2 for count in app.evaluation.employee_counts.values())
    assert "必須条件：すべて適合" in app.result_summary.get()
    assert "希望休採用 3/3申請" in app.result_summary.get()
    assert "順位合計4" in app.result_summary.get()
    assert app.edited
    assert "手修正" in app.result_summary.get()
    assert str(app.save_button.cget("state")) == "normal"
    app.notebook.select(app.result_tab)
    root.update()
    save_button_geometry = widget_bounds(root, app.save_button)
    assert save_button_geometry["mapped"], "結果保存ボタンが表示されていません。"
    assert save_button_geometry["inside_root"], f"結果保存ボタンがウィンドウ外です: {save_button_geometry}"
    r3_iid = next(iid for iid in app.request_tree.get_children()
                  if app.request_tree.item(iid, "values")[0] == "R3")
    report = {
        "python": sys.version,
        "os": platform.platform(),
        "tk": root.tk.call("info", "patchlevel"),
        "mapped": bool(root.winfo_ismapped()),
        "window": [root.winfo_width(), root.winfo_height()],
        "result_table": [app.result_tree.winfo_width(), app.result_tree.winfo_height()],
        "preview_geometry": preview_geometry,
        "save_button_geometry": save_button_geometry,
        "daily_footer_bbox": row_bbox(app.result_tree, "daily_actual"),
        "request_r3_bbox": row_bbox(app.request_tree, r3_iid),
        "result_summary": app.result_summary.get(),
        "restored_counts_valid": True,
        "sample_input_read": True,
        "sample_computed": True,
        "saved_excel": str(saved),
        "manual_edit_invalidated": True,
        "invalid_save_blocked": True,
        "manual_edit_restored": True,
        "violations": violations,
        "visual_review": "画像の目視確認はこのスクリプトの対象外",
    }
    report_path = args.output_dir / "gui_smoke.json"
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
