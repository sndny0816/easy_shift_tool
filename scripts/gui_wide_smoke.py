"""実Tkで3日分の独立希望と、明示した管理者有給確認を検証する。"""
from __future__ import annotations

import argparse
from datetime import timedelta
import json
from pathlib import Path
import platform
import sys
import time
import tkinter as tk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from sinta_shift.gui import ShiftApp, REQUEST_FORMAT_NAMES
from sinta_shift.models import PUBLIC_REQUEST, PAID_REQUEST
from sinta_shift.excel_io import WIDE_REQUEST_FIELDS
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


def geometry(root, widget):
    root.update_idletasks()
    x, y = widget.winfo_rootx() - root.winfo_rootx(), widget.winfo_rooty() - root.winfo_rooty()
    width, height = widget.winfo_width(), widget.winfo_height()
    mapped = bool(widget.winfo_ismapped())
    inside = mapped and width > 0 and height > 0 and x >= 0 and y >= 0 and x + width <= root.winfo_width() and y + height <= root.winfo_height()
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
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/gui_wide_verification"))
    parser.add_argument("--hold", type=float, default=0)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    inputs = args.output_dir / "inputs"
    inputs.mkdir(exist_ok=True)
    prepare_tk_runtime()
    root = tk.Tk()
    app = ShiftApp(root)
    root.update()
    app.load_independent_sample()
    wait_for(root, lambda: app.base_data is not None and not app.busy)
    assert app.base_data.request_format == "independent"
    assert len(app.base_data.independent_requests) == 9
    assert sum(r.leave_type == PUBLIC_REQUEST for r in app.base_data.independent_requests) == 6
    assert sum(r.leave_type == PAID_REQUEST for r in app.base_data.independent_requests) == 3
    assert app.request_format_var.get() == REQUEST_FORMAT_NAMES["independent"]
    assert app.data is None and app.schedule is None
    assert len(app.paid_manager_tree.get_children()) == 3
    assert all(app.paid_manager_tree.item(iid, "values")[-1] == "確認待ち" for iid in app.paid_manager_tree.get_children())
    assert str(app.solve_button.cget("state")) == "disabled"
    assert str(app.save_button.cget("state")) == "disabled"
    assert str(app.confirm_paid_button.cget("state")) == "normal"
    app.start_solve()
    assert app.schedule is None
    app.notebook.select(app.import_tab)
    app.role_notebook.select(app.role_notebook.tabs()[app._roles.index("希望休")])
    wait_for(root, lambda: app._preview_role == "希望休" and not app.busy)
    assert len(app.preview_tree.get_children()) == 3
    preview_geometry = geometry(root, app.preview_tree)
    assert preview_geometry["mapped"] and preview_geometry["inside_root"]
    mapping_geometry = {field: geometry(root, app.mapping_boxes["希望休"]["columns"][field]) for field in WIDE_REQUEST_FIELDS}
    assert all(info["mapped"] and info["inside_root"] for info in mapping_geometry.values())
    # 別名の横形式を明示対応し、形式も含めてJSONへ保存・復元する。
    aliases = dict(zip(WIDE_REQUEST_FIELDS, ("職員コード", "申請日A", "休暇種別A", "申請日B", "休暇種別B", "申請日C", "休暇種別C")))
    alias_path = (inputs / "wide_alias.xlsx").resolve()
    workbook = load_workbook(app.input_path)
    source_headers = [cell.value for cell in workbook["希望休"][1]]
    for column, field in enumerate(source_headers, 1):
        if field in aliases:
            workbook["希望休"].cell(1, column, aliases[field])
    workbook.save(alias_path)
    app.load_file(alias_path)
    idle(root, app)
    app.request_format_var.set(REQUEST_FORMAT_NAMES["independent"])
    app._request_format_changed()
    for field, actual in aliases.items():
        app.mapping_vars["希望休"]["columns"][field].set(actual)
    app.read_conditions()
    idle(root, app)
    assert app.base_data is not None and len(app.base_data.independent_requests) == 9
    mapping_path = args.output_dir / "wide_mapping.json"
    app.save_mapping(mapping_path)
    idle(root, app)
    payload = json.loads(mapping_path.read_text(encoding="utf-8"))
    assert payload["mappings"]["希望休"]["format"] == "independent"
    assert payload["mappings"]["希望休"]["columns"] == aliases
    app.request_format_var.set(REQUEST_FORMAT_NAMES["alternatives"])
    app._request_format_changed()
    app.load_mapping(mapping_path)
    idle(root, app)
    assert app._get_mappings()["希望休"]["format"] == "independent"
    assert app._get_mappings()["希望休"]["columns"] == aliases
    app.read_conditions()
    idle(root, app)
    settings_path = args.output_dir / "wide_settings.json"
    app.save_settings(settings_path)
    idle(root, app)
    before_settings = app._get_settings()
    app.load_settings(settings_path)
    idle(root, app)
    assert app._get_settings() == before_settings
    app.read_conditions()
    idle(root, app)
    paid_ids = [r.request_id for r in app._paid_requests()]
    assert len(paid_ids) == 3
    # 1件の明示確認では他2件を自動確定しない。
    app.set_paid_request_confirmation([paid_ids[0]], confirmed=True)
    idle(root, app)
    assert app.confirmed_paid_request_ids == {paid_ids[0]}
    assert app.data is None
    assert sum(app.paid_manager_tree.item(iid, "values")[-1] == "有給確定" for iid in app.paid_manager_tree.get_children()) == 1
    assert len(app.base_data.paid_leaves) == 0
    app.select_all_paid_requests()
    app.set_paid_request_confirmation(confirmed=True)
    idle(root, app)
    assert app.confirmed_paid_request_ids == set(paid_ids)
    assert app.data is not None, app.status_var.get()
    assert len(app.data.paid_leaves) == 3
    assert len(app.base_data.paid_leaves) == 0
    # 休日区分を無効にしても既存の有給確認は解除でき、再有効化で自動復活しない。
    app.leave_enabled_var.set(False)
    app.set_paid_request_confirmation([paid_ids[0]], confirmed=False)
    idle(root, app)
    assert paid_ids[0] not in app.confirmed_paid_request_ids
    app.leave_enabled_var.set(True)
    app.read_conditions()
    idle(root, app)
    assert paid_ids[0] not in app.confirmed_paid_request_ids
    assert app.data is None and app.schedule is None
    assert str(app.solve_button.cget("state")) == "disabled"
    assert str(app.save_button.cget("state")) == "disabled"
    assert any(app.paid_manager_tree.item(iid, "values")[-1] == "確認待ち" for iid in app.paid_manager_tree.get_children())
    app.set_paid_request_confirmation([paid_ids[0]], confirmed=True)
    idle(root, app)
    assert app.confirmed_paid_request_ids == set(paid_ids)
    assert app.data is not None, app.status_var.get()
    solve(root, app)
    assert len(app.evaluation.request_results) == 6
    assert app.evaluation.adopted_count == 6
    assert len(app.evaluation.paid_request_results) == 3
    assert all(r.status == "有給確定" for r in app.evaluation.paid_request_results)
    assert len(app.request_tree.get_children()) == 9
    assert "日数" in app.request_tree.cget("columns")
    assert not any("順位" in column for column in app.request_tree.cget("columns"))
    rows = [app.request_tree.item(iid, "values") for iid in app.request_tree.get_children()]
    assert all(str(row[-2]) == "1" for row in rows)
    assert all(row[-1] == "採用" for row in rows if row[2] == PUBLIC_REQUEST)
    assert all(row[-1] == "有給確定" for row in rows if row[2] == PAID_REQUEST)
    assert "公休希望採用 6/6" in app.result_summary.get()
    assert "有給確定 3件" in app.result_summary.get()
    assert "順位" not in app.result_summary.get()
    assert all(count == 9 for count in app.evaluation.employee_public_counts.values())
    assert all(count == 1 for count in app.evaluation.employee_paid_counts.values())
    results = (args.output_dir / "results").resolve()
    results.mkdir(exist_ok=True)
    app.save_result(results)
    wait_for(root, lambda: app.last_saved is not None and not app.busy)
    saved = app.last_saved
    assert saved.is_file()
    count_before = set(results.glob("*.xlsx"))
    # 確認解除で完成結果を破棄し、保存も停止する。
    app.set_paid_request_confirmation([paid_ids[0]], confirmed=False)
    idle(root, app)
    assert app.data is None and app.schedule is None
    assert paid_ids[0] not in app.confirmed_paid_request_ids
    assert str(app.save_button.cget("state")) == "disabled"
    app.save_result(results)
    assert set(results.glob("*.xlsx")) == count_before
    # 同じファイルの設定再確認では、残りの確認IDは維持する。
    retained_ids = set(app.confirmed_paid_request_ids)
    app.max_days_var.set(app.max_days_var.get())
    app.read_conditions()
    idle(root, app)
    assert app.confirmed_paid_request_ids == retained_ids
    # 同じファイルの申請日が変わった場合、その古いIDの確認を引き継がない。
    app.set_paid_request_confirmation(paid_ids, confirmed=True)
    idle(root, app)
    old_request = next(r for r in app._paid_requests() if r.request_id == paid_ids[0])
    workbook = load_workbook(alias_path)
    sheet = workbook["希望休"]
    headers = [cell.value for cell in sheet[1]]
    changed = False
    for row in sheet.iter_rows(min_row=2):
        if str(row[headers.index(aliases["職員ID"])] .value) != old_request.employee_id:
            continue
        for slot in (1, 2, 3):
            if row[headers.index(aliases[f"区分{slot}"])].value == PAID_REQUEST:
                old_value = row[headers.index(aliases[f"希望日{slot}"])].value
                row[headers.index(aliases[f"希望日{slot}"])].value = old_value + timedelta(days=1)
                changed = True
    assert changed
    workbook.save(alias_path)
    app.read_conditions()
    idle(root, app)
    assert paid_ids[0] not in app.confirmed_paid_request_ids
    assert len(app.confirmed_paid_request_ids) == 2
    assert app.data is None
    # 新ファイルの明示読込では確認を全解除する。最後は再確認した正常な案を表示する。
    app.load_independent_sample()
    wait_for(root, lambda: app.base_data is not None and not app.busy)
    assert not app.confirmed_paid_request_ids
    app.select_all_paid_requests()
    app.set_paid_request_confirmation(confirmed=True)
    idle(root, app)
    solve(root, app)
    app.notebook.select(app.result_tab)
    root.update()
    save_geometry = geometry(root, app.save_button)
    assert save_geometry["mapped"] and save_geometry["inside_root"]
    report = {"python": sys.version, "os": platform.platform(), "tk": root.tk.call("info", "patchlevel"),
              "window": [root.winfo_width(), root.winfo_height()], "saved_excel": str(saved),
              "independent_requests": 9, "public_requests": 6, "paid_requests": 3,
              "unconfirmed_processing_blocked": True, "single_confirmation_not_automatic_all": True,
              "all_confirmed_and_solved": True, "confirmation_revocation_blocks_save": True,
              "revocation_while_leave_disabled_retained_after_reenable": True,
              "changed_date_does_not_inherit_confirmation": True, "new_file_clears_confirmation": True,
              "mapping_json_round_trip": True, "settings_json_round_trip": True,
              "preview_rows": 3, "preview_geometry": preview_geometry, "mapping_geometry": mapping_geometry,
              "save_button_geometry": save_geometry, "result_summary": app.result_summary.get(),
              "visual_review": "実画面の目視確認は別途必要"}
    (args.output_dir / "gui_wide_smoke.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.hold:
        deadline = time.monotonic() + args.hold
        while time.monotonic() < deadline:
            root.update()
            time.sleep(0.04)
    app.close()


if __name__ == "__main__":
    main()
