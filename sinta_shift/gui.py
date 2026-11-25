"""日本語のローカル勤務表GUI。処理はworker、Tkの変更はmain threadのみ。"""
from __future__ import annotations

from datetime import date, timedelta
from dataclasses import replace
from decimal import Decimal, InvalidOperation
import json
from pathlib import Path
import queue
import tempfile
import threading
import tkinter as tk
import unicodedata
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from . import __version__
from .models import (InputData, InputError, Issue, SolveResult, SolverStatus, ShiftSettings,
                     SHIFT_MODES, DAY_ONLY, TWO_SHIFT, THREE_SHIFT, OFF, PUBLIC_OFF, PAID_LEAVE, RECOVERY,
                     ShiftTime, DIRECT_DAYS, PUBLIC_DAYS, PUBLIC_REQUEST, PAID_REQUEST, assignment_label, assignment_labels)
from .planning import (configure_input, public_holiday_target, confirm_paid_requests, set_leave_choices,
                       plan_paid_mix, attendance_ratio_for_day, attendance_source_for_day, week_start)
from .evaluation import evaluate_schedule
from .validation import validate_input
from .solver import solve_schedule
from .runtime import prepare_tk_runtime
from .visual_ui import Collapsible, AttendanceCalendar, MixBar, ScrollPanel
from .request_heatmap_ui import RequestHeatmap
from .request_heatmap import build_request_heatmap
from .error_guidance import guidance_for_issue, issues_for_solve_result, issues_for_exception
from .excel_io import (REQUIRED_FIELDS, inspect_workbook, preview_sheet,
                       read_input, export_template, export_result, settings_to_payload, settings_from_payload,
                       OPTIONAL_SHEET_FIELDS, WIDE_REQUEST_FIELDS, export_leave_sample, export_independent_sample)

NAVY = "#17324D"
GREEN = "#E5F2E8"
INK = "#213549"
BG = "#F3F6F8"
DANGER = "#A02C32"
WEEKDAYS = "月火水木金土日"
UNUSED_COLUMN = "（使用しない）"
OPTIONAL_FIELDS = {
    "職員": ("前期間の連続勤務日数", "前日の勤務区分", "対象期間の公休日数"),
    "必要人数": ("日勤人数", "夜勤人数", "準夜勤人数", "深夜勤人数"),
}
INPUT_FIELDS = {**REQUIRED_FIELDS, **OPTIONAL_SHEET_FIELDS}
WORK_DAYS_FIELD = "対象期間の勤務日数"
REQUEST_FORMAT_LABELS = {"3日分（各日独立）": "independent", "従来（代替候補）": "alternatives", "自動判定": "auto"}
REQUEST_FORMAT_NAMES = {value: label for label, value in REQUEST_FORMAT_LABELS.items()}


def display_day(day: date) -> str:
    return f"{day:%m/%d}({WEEKDAYS[day.weekday()]})"


class ShiftApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Sinta 勤務表作成 {__version__}")
        self.root.geometry("1220x850")
        self.root.minsize(1120, 800)
        self.root.configure(bg=BG)
        self.input_path: Path | None = None
        self.data: InputData | None = None
        self.base_data: InputData | None = None
        self.confirmed_paid_request_ids = set()
        self.converted_public_request_ids = set()
        self.weekly_ratio_overrides = {}
        self.date_ratio_overrides = {}
        self._syncing_visual = False
        self._active_step = 1
        self.result_week_index = 0
        self.current_result_view = "schedule"
        self._heatmap_report = ()
        self._heatmap_data = None
        self._heatmap_after = None
        self._heatmap_has_computed = False
        self._paid_request_by_iid = {}
        self._settings_trace_ready = False
        self._loading_settings = False
        self._settings_from_file_pending = False
        self.result: SolveResult | None = None
        self.schedule = None
        self.evaluation = None
        self.edited = False
        self.busy = False
        self.last_saved: Path | None = None
        self.last_time_limit = 30.0
        self._temp = tempfile.TemporaryDirectory(prefix="sinta-shift-")
        self._queue = queue.Queue()
        self._controls: list[tuple[tk.Widget, str]] = []
        self._sheet_names: list[str] = []
        self._columns: dict[str, list[str]] = {}
        self.mapping_vars = {}
        self.mapping_boxes = {}
        self._issue_by_iid: dict[str, Issue] = {}
        self._issue_operation = None
        self.selected_error_guidance = None
        self._closed = False
        self._sample_pending = False
        self._setup_style()
        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(70, self._poll)

    def _setup_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        self.root.option_add("*Font", ("Helvetica", 12))
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=INK, font=("Helvetica", 12))
        style.configure("TButton", padding=(12, 7), font=("Helvetica", 12))
        style.configure("Accent.TButton", background=NAVY, foreground="white", padding=(16, 7))
        style.map("Accent.TButton", background=[("active", "#244E6C"), ("disabled", "#8999A8")])
        style.configure("Green.TLabel", background=GREEN, foreground=NAVY, padding=6)
        style.configure("TLabelframe", background=BG)
        style.configure("TLabelframe.Label", background=BG, foreground=NAVY,
                        font=("Helvetica", 12, "bold"))
        style.configure("Treeview", rowheight=32, font=("Helvetica", 12),
                        background="white", fieldbackground="white", foreground=INK)
        style.configure("Treeview.Heading", background=NAVY, foreground="white",
                        font=("Helvetica", 11, "bold"), padding=7)
        style.map("Treeview", background=[("selected", "#C8DDED")],
                  foreground=[("selected", NAVY)])
        style.configure("TNotebook.Tab", padding=(13, 6), font=("Helvetica", 12))
        style.configure("TCombobox", padding=4)
        style.layout("Workflow.TNotebook.Tab", [])
        style.configure("Workflow.TNotebook", tabmargins=0, borderwidth=0)
        style.configure("Card.TFrame", background="white")
        style.configure("Card.TLabel", background="white", foreground=INK)
        style.configure("Hero.TButton", padding=(18, 12), font=("Helvetica", 13, "bold"),
                        background=NAVY, foreground="white")
        style.map("Hero.TButton", background=[("active", "#244E6C"), ("disabled", "#97A4B0")])
        style.configure("Step.TButton", padding=(10, 14), font=("Helvetica", 13, "bold"),
                        background=NAVY, foreground="white", anchor="w", relief="flat", borderwidth=0)
        style.map("Step.TButton", background=[("disabled", NAVY), ("active", "#31556D")],
                  foreground=[("disabled", "#D4E2EB"), ("active", "white")])
        style.configure("CurrentStep.TButton", padding=(10, 14), font=("Helvetica", 13, "bold"),
                        background="#E1F0E6", foreground=NAVY, anchor="w", relief="flat", borderwidth=0)
        style.map("CurrentStep.TButton", background=[("disabled", "#E1F0E6"), ("active", "#CFE7D8")],
                  foreground=[("disabled", "#476453"), ("active", NAVY)])

    def _control(self, widget, state="normal"):
        self._controls.append((widget, state))
        return widget

    def _button(self, parent, text, command, accent=False):
        return self._control(ttk.Button(parent, text=text, command=command,
                                       style="Accent.TButton" if accent else "TButton"))

    def _build(self):
        head = tk.Frame(self.root, bg=NAVY, padx=22, pady=11)
        head.pack(fill="x")
        tk.Label(head, text="Sinta 勤務表作成", bg=NAVY, fg="white",
                 font=("Helvetica", 21, "bold")).pack(side="left")
        tk.Label(head, text="担当者確認用のシフト案  |  データはこの端末内で処理", bg=NAVY,
                 fg="#DDECE4", font=("Helvetica", 11)).pack(side="right")
        shell = ttk.Frame(self.root, padding=(14, 12))
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)
        sidebar = tk.Frame(shell, bg=NAVY, width=182, padx=10, pady=14)
        sidebar.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        sidebar.pack_propagate(False)
        self.step_buttons = {}
        labels = ("1  Excelを\n    読み込む", "2  人数と連勤\n    を決める",
                  "3  公休・有給\n    を決める", "4  シフト確認\n    Excelに保存")
        for step, text in enumerate(labels, 1):
            button = ttk.Button(sidebar, text=text, style="Step.TButton", command=lambda number=step: self.go_step(number))
            button.pack(fill="x", pady=(0, 8))
            self.step_buttons[step] = button
        self.completed_var = tk.StringVar(value="完了 0 / 4")
        tk.Label(sidebar, textvariable=self.completed_var, bg=NAVY, fg="#DDECE4", font=("Helvetica", 12)).pack(anchor="w", pady=(18, 8))
        tk.Label(sidebar, text="順番に進めるだけで\n条件確認からExcel保存まで", bg=NAVY, fg="#DDECE4",
                 justify="left", font=("Helvetica", 10)).pack(anchor="w")
        content = ttk.Frame(shell)
        content.grid(row=0, column=1, sticky="nsew")
        self.headline_var = tk.StringVar(value="1  Excelを読み込む")
        ttk.Label(content, textvariable=self.headline_var, font=("Helvetica", 19, "bold")).pack(anchor="w", pady=(0, 5))
        self.next_action_var = tk.StringVar(value="次にすること：ファイル、またはサンプルを選んでください。")
        ttk.Label(content, textvariable=self.next_action_var, style="Green.TLabel", wraplength=940).pack(fill="x", pady=(0, 8))
        self.path_var = tk.StringVar(value="入力：未選択")
        self.notebook = ttk.Notebook(content, style="Workflow.TNotebook")
        self.import_tab = ttk.Frame(self.notebook, padding=6)
        self.settings_tab = ttk.Frame(self.notebook, padding=6)
        self.manager_tab = ttk.Frame(self.notebook, padding=6)
        self.conditions_tab = ttk.Frame(self.notebook, padding=8)
        self.result_tab = ttk.Frame(self.notebook, padding=6)
        self.errors_tab = ttk.Frame(self.notebook, padding=6)
        for frame, text in ((self.import_tab, "読み込み"), (self.settings_tab, "人数と連勤"),
                            (self.manager_tab, "公休・有給"), (self.conditions_tab, "条件確認"),
                            (self.result_tab, "シフト確認"), (self.errors_tab, "エラー一覧")):
            self.notebook.add(frame, text=text)
        self._build_import()
        self._build_settings()
        self._build_manager()
        self._build_conditions()
        self._build_results()
        self._build_errors()
        bottom = ttk.Frame(content)
        bottom.pack(side="bottom", fill="x", pady=(8, 0))
        self.workflow_button = self._control(ttk.Button(bottom, text="Excelを選ぶ", command=self._workflow_action, style="Hero.TButton"))
        self.workflow_button.pack(side="right")
        self.error_alert_button = self._button(bottom, "エラーの直し方", self.show_error_guidance)
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=90)
        self.progress.pack(side="left", padx=(0, 10))
        self.status_var = tk.StringVar(value="ファイルまたはサンプルを選んでください。")
        self.status_label = ttk.Label(bottom, textvariable=self.status_var, wraplength=610, justify="left")
        self.status_label.pack(side="left", fill="x", expand=True)
        self.status_label.bind("<Configure>", self._wrap_status)
        self.saved_var = tk.StringVar(value="")
        ttk.Label(content, textvariable=self.saved_var, wraplength=930).pack(side="bottom", anchor="w", pady=(5, 0))
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._workflow_changed)
        self._settings_trace_ready = True
        self._update_buttons()

    def _wrap_status(self, event):
        width = max(1, event.width - 6)
        if int(self.status_label.cget("wraplength")) != width:
            self.status_label.configure(wraplength=width)

    def go_step(self, number):
        if self.busy:
            return
        self._active_step = number
        self.notebook.select({1: self.import_tab, 2: self.settings_tab, 3: self.manager_tab, 4: self.result_tab}[number])
        self._sync_workflow()

    def _workflow_changed(self, _event=None):
        selected = self.notebook.select()
        pages = {str(self.import_tab): 1, str(self.settings_tab): 2, str(self.conditions_tab): 2,
                 str(self.manager_tab): 3, str(self.result_tab): 4, str(self.advanced_settings_tab): 2}
        if selected in pages:
            self._active_step = pages[selected]
        self._sync_workflow()

    def _sync_workflow(self):
        if not hasattr(self, "workflow_button"):
            return
        pending = len([r for r in self._paid_requests() if r.request_id not in self.confirmed_paid_request_ids
                       and r.request_id not in self.converted_public_request_ids])
        valid_result = self.evaluation is not None and self.evaluation.valid and self.schedule is not None
        completed = int(self.input_path is not None) + int(self.data is not None) + int(self.base_data is not None and not pending) + int(valid_result and self.last_saved is not None)
        self.completed_var.set(f"完了 {completed} / 4")
        titles = {1: "Excelを読み込む", 2: "人数と連勤を決める", 3: "公休・有給を決める", 4: "シフトを確認して保存する"}
        self.headline_var.set(f"{self._active_step}  {titles[self._active_step]}")
        for step, button in self.step_buttons.items():
            button.configure(style="CurrentStep.TButton" if step == self._active_step else "Step.TButton",
                             state="disabled" if self.busy else "normal")
        if self.input_path is None:
            self.workflow_button.configure(text="Excelを選ぶ", state="disabled" if self.busy else "normal")
            self.next_action_var.set("次にすること：先に『1 Excelを読み込む』でファイルかサンプルを選んでください。")
            return
        if self._active_step == 1:
            label = "Excelを選ぶ" if self.input_path is None else "内容を確認して次へ"
            hint = "Excelまたはサンプルを選び、先頭20行と列名を確認してください。"
        elif self._active_step == 2:
            label = "条件を確認して公休・有給へ"
            hint = "カレンダーの日付・週を押して出勤率を設定し、最大連勤を決めてください。"
        elif self._active_step == 3:
            label = "この配分を確定" if pending else "シフト案を作る"
            hint = f"未決定の有給希望が{pending}件あります。配分案を見てから確定してください。" if pending else "公休・有給の確認が終わりました。シフト案を作成できます。"
        else:
            invalid_edit = self.schedule is not None and self.evaluation is not None and not self.evaluation.valid
            label = "Excelに保存" if valid_result else "違反の理由を確認" if invalid_edit else "シフト案を作る" if self.data else "条件をもう一度確認"
            hint = ("週を切り替えて日ごとの区分と人数を確認し、Excelへ保存してください。" if valid_result else
                    "変更後の必須条件に違反があります。理由を確認してセルを直すと保存できます。" if invalid_edit else
                    "条件確認を実行すると、計算または修正が必要な理由を確認できます。")
        self.workflow_button.configure(text=label, state="disabled" if self.busy else "normal")
        self.next_action_var.set("次にすること：" + hint)

    def _workflow_action(self):
        if self.busy:
            return
        if self.input_path is None:
            self.go_step(1)
            self.choose_file()
            return
        if self._active_step == 1:
            self.choose_file() if self.input_path is None else self.read_conditions(show_settings=True)
        elif self._active_step == 2:
            self.read_conditions(show_manager=True)
        elif self._active_step == 3:
            if any(r.request_id not in self.confirmed_paid_request_ids and r.request_id not in self.converted_public_request_ids for r in self._paid_requests()):
                self.apply_paid_mix()
            elif self.data is not None:
                self.start_solve()
            else:
                self.read_conditions(show_manager=True)
        elif self.schedule is not None and self.evaluation is not None and self.evaluation.valid:
            self.save_result()
        elif self.schedule is not None and self.evaluation is not None and not self.evaluation.valid:
            self.notebook.select(self.errors_tab)
        elif self.data is not None:
            self.start_solve()
        else:
            self.read_conditions()

    def _build_import(self):
        actions = ttk.Frame(self.import_tab)
        actions.pack(fill="x", pady=(0, 8))
        self._button(actions, "Excelを選ぶ", self.choose_file, True).pack(side="left")
        self._button(actions, "サンプルで試す", self.load_sample).pack(side="left", padx=7)
        self._button(actions, "3日分の希望・有給サンプル", self.load_independent_sample).pack(side="left")
        self._button(actions, "入力ひな形を保存", self.save_template).pack(side="left", padx=7)
        ttk.Label(self.import_tab, textvariable=self.path_var, wraplength=920).pack(anchor="w", pady=(0, 6))
        self.mapping_summary_var = tk.StringVar(value="列名を自動で対応付けます。違う列名の場合は詳しい設定を開いてください。")
        ttk.Label(self.import_tab, textvariable=self.mapping_summary_var, style="Green.TLabel", wraplength=920).pack(fill="x", pady=(0, 6))
        self.import_details = Collapsible(self.import_tab, "詳しい設定：シート・列名の対応")
        self.import_details.pack(fill="x", pady=(0, 7))
        self.role_notebook = ttk.Notebook(self.import_details.body)
        self.role_notebook.pack(fill="x")
        self._roles = list(INPUT_FIELDS)
        self._preview_role = None
        for role, fields in INPUT_FIELDS.items():
            role_tab = ttk.Frame(self.role_notebook, padding=(8, 6))
            self.role_notebook.add(role_tab, text=f"{role}（任意）" if role in OPTIONAL_SHEET_FIELDS else role)
            group = ttk.LabelFrame(role_tab, text=f"{role}の列対応", padding=(8, 4))
            group.pack(fill="x")
            values = {"sheet": tk.StringVar(value=role), "header_row": tk.IntVar(value=1), "columns": {}}
            boxes = {}
            ttk.Label(group, text="シート").grid(row=0, column=0, sticky="w")
            sheet_box = self._control(ttk.Combobox(group, textvariable=values["sheet"], width=13, state="readonly"), "readonly")
            sheet_box.grid(row=1, column=0, padx=(0, 8), sticky="ew")
            sheet_box.bind("<<ComboboxSelected>>", lambda _e, r=role: self.refresh_role(r))
            ttk.Label(group, text="見出し行").grid(row=0, column=1, sticky="w")
            spin = self._control(ttk.Spinbox(group, from_=1, to=1000, width=5,
                                           textvariable=values["header_row"]))
            spin.grid(row=1, column=1, padx=(0, 8))
            spin.bind("<KeyRelease>", lambda _e: self._invalidate_mapping())
            spin.bind("<<Increment>>", lambda _e: self._invalidate_mapping())
            spin.bind("<<Decrement>>", lambda _e: self._invalidate_mapping())
            for col, field in enumerate(() if role == "希望休" else fields, 2):
                values["columns"][field] = tk.StringVar(value=field)
                ttk.Label(group, text=field).grid(row=0, column=col, sticky="w")
                box = self._control(ttk.Combobox(group, textvariable=values["columns"][field],
                                               width=13, state="readonly"), "readonly")
                box.grid(row=1, column=col, padx=(0, 8), sticky="ew")
                box.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_mapping())
                boxes[field] = box
                group.columnconfigure(col, weight=1)
            self._button(group, "見出しを反映", lambda r=role: self.refresh_role(r)).grid(row=1, column=6, sticky="e")
            group.columnconfigure(0, weight=1)
            if role == "希望休":
                ttk.Label(group, text="入力形式").grid(row=0, column=2, sticky="w")
                self.request_format_var = tk.StringVar(value="自動判定")
                self._detected_request_format = "independent"
                self.request_format_box = self._control(ttk.Combobox(group, textvariable=self.request_format_var,
                    values=tuple(REQUEST_FORMAT_LABELS), state="readonly", width=25), "readonly")
                self.request_format_box.grid(row=1, column=2, columnspan=3, sticky="ew", padx=(0, 10))
                self.request_format_box.bind("<<ComboboxSelected>>", self._request_format_changed)
                request_fields_frame = ttk.Frame(role_tab)
                request_fields_frame.pack(fill="x", pady=(6, 0))
                self.request_field_frames = {}
                for field in dict.fromkeys((*REQUIRED_FIELDS[role], *WIDE_REQUEST_FIELDS)):
                    values["columns"][field] = tk.StringVar(value=field)
                    item = ttk.Frame(request_fields_frame)
                    ttk.Label(item, text=field).pack(anchor="w")
                    box = self._control(ttk.Combobox(item, textvariable=values["columns"][field],
                        width=18, state="readonly"), "readonly")
                    box.pack(fill="x")
                    box.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_mapping())
                    boxes[field] = box
                    self.request_field_frames[field] = item
                for column in range(4):
                    request_fields_frame.columnconfigure(column, weight=1)
                self._layout_request_fields()
            if role in OPTIONAL_FIELDS:
                optional_group = ttk.LabelFrame(role_tab, text="任意列（履歴・区分人数）", padding=(8, 4))
                optional_group.pack(fill="x", pady=(5, 0))
                for col, field in enumerate(OPTIONAL_FIELDS[role]):
                    values["columns"][field] = tk.StringVar(value=UNUSED_COLUMN)
                    ttk.Label(optional_group, text=field).grid(row=0, column=col, sticky="w")
                    box = self._control(ttk.Combobox(optional_group, textvariable=values["columns"][field],
                                                   width=16, state="readonly"), "readonly")
                    box.grid(row=1, column=col, padx=(0, 10), sticky="ew")
                    box.bind("<<ComboboxSelected>>", lambda _e: self._invalidate_mapping())
                    optional_group.columnconfigure(col, weight=1)
                    boxes[field] = box
            self.mapping_vars[role] = values
            self.mapping_boxes[role] = {"sheet": sheet_box, "header": spin, "columns": boxes}
        self.role_notebook.bind("<<NotebookTabChanged>>", self._on_role_preview_change)
        commands = ttk.Frame(self.import_details.body)
        commands.pack(fill="x", pady=8)
        self._button(commands, "対応設定を保存", self.save_mapping).pack(side="left")
        self._button(commands, "対応設定を読込", self.load_mapping).pack(side="left", padx=8)
        self.check_button = self._button(commands, "条件を確認", self.read_conditions, True)
        self.check_button.pack(side="right")
        self.preview_label = ttk.Label(self.import_tab, text="先頭20行のプレビュー：未読込")
        self.preview_label.pack(anchor="w", pady=(0, 4))
        self.preview_tree = self._tree(self.import_tab, (), height=4)

    def _build_settings(self):
        self.advanced_settings_tab = ttk.Frame(self.notebook, padding=6)
        self.notebook.add(self.advanced_settings_tab, text="詳しい勤務設定")
        ttk.Button(self.advanced_settings_tab, text="‹ 人数と連勤に戻る", command=lambda: self.go_step(2)).pack(anchor="w", pady=(0, 8))
        scroll = ScrollPanel(self.advanced_settings_tab)
        scroll.pack(fill="both", expand=True)
        self._build_detailed_settings(scroll.body)
        top = ttk.Frame(self.settings_tab)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="日付か週を押すと、その範囲の出勤率を変えられます。", font=("Helvetica", 12)).pack(side="left")
        self._button(top, "詳しい設定", lambda: self.notebook.select(self.advanced_settings_tab)).pack(side="right")
        pane = ttk.Frame(self.settings_tab)
        pane.pack(fill="both", expand=True)
        pane.columnconfigure(0, weight=3)
        pane.columnconfigure(1, weight=1, minsize=210)
        pane.rowconfigure(0, weight=1)
        self.attendance_calendar = AttendanceCalendar(pane,
            lambda day: self._attendance_dialog(day, "day"), lambda monday: self._attendance_dialog(monday, "week"))
        self.attendance_calendar.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.settings_controls_panel = ScrollPanel(pane, background="white", body_style="Card.TFrame")
        self.settings_controls_panel.grid(row=0, column=1, sticky="nsew")
        controls = ttk.Frame(self.settings_controls_panel.body, padding=10, style="Card.TFrame")
        controls.pack(fill="x", expand=True)
        ttk.Label(controls, text="基本の出勤率", font=("Helvetica", 14, "bold"), style="Card.TLabel").pack(anchor="w")
        self.visual_source_box = self._control(ttk.Combobox(controls, textvariable=self.people_source_var,
            values=("Excelの配置人数", "曜日ごとの割合"), state="readonly", width=18), "readonly")
        self.visual_source_box.pack(fill="x", pady=(8, 5))
        days_frame = ttk.Frame(controls, style="Card.TFrame")
        days_frame.pack(fill="x")
        self.visual_ratio_boxes = []
        for index, day in enumerate(WEEKDAYS):
            ttk.Label(days_frame, text=day, style="Card.TLabel").grid(row=index, column=0, padx=(0, 12), pady=2)
            box = self._control(ttk.Spinbox(days_frame, from_=0, to=100, increment=1,
                textvariable=self.ratio_vars[index], width=6))
            box.grid(row=index, column=1, sticky="w", pady=2)
            ttk.Label(days_frame, text="％", style="Card.TLabel").grid(row=index, column=2, padx=4)
            self.visual_ratio_boxes.append(box)
        ttk.Separator(controls).pack(fill="x", pady=9)
        self._control(ttk.Checkbutton(controls, text="連勤の上限を使う", variable=self.max_enabled_var)).pack(anchor="w")
        self.visual_cap_var = tk.StringVar(value="最大5日まで連続勤務")
        self.visual_cap_label = ttk.Label(controls, textvariable=self.visual_cap_var,
            font=("Helvetica", 12, "bold"), style="Card.TLabel", wraplength=180)
        self.visual_cap_label.pack(anchor="w", pady=(5, 3))
        self.cap_scale = self._control(ttk.Scale(controls, from_=1, to=31, command=self._cap_scale_changed))
        self.cap_scale.pack(fill="x")
        self.visual_cap_spin = self._control(ttk.Spinbox(controls, from_=1, to=366, width=6, textvariable=self.max_days_var))
        self.visual_cap_spin.pack(anchor="w", pady=(4, 3))
        self.visual_cap_explanation = ttk.Label(controls, text="明けも実勤務として数えます。",
            style="Card.TLabel", wraplength=180)
        self.visual_cap_explanation.pack(anchor="w")
        ttk.Separator(controls).pack(fill="x", pady=9)
        ttk.Label(controls, text="勤務モード", style="Card.TLabel").pack(anchor="w")
        self._control(ttk.Combobox(controls, textvariable=self.mode_var,
            values=(DAY_ONLY, TWO_SHIFT, THREE_SHIFT), state="readonly", width=18), "readonly").pack(fill="x", pady=(4, 5))
        self.clear_attendance_button = self._button(controls, "週・日付の指定を\nすべて解除", self.clear_attendance_overrides)
        self.clear_attendance_button.pack(fill="x", pady=(5, 0))
        def fit_control_labels(event):
            width = max(100, event.width - 20)
            self.visual_cap_label.configure(wraplength=width)
            self.visual_cap_explanation.configure(wraplength=width)
        controls.bind("<Configure>", fit_control_labels)
        self.visual_settings_note = tk.StringVar(value="人数は職員数×出勤率を切り上げます。")
        ttk.Label(self.settings_tab, textvariable=self.visual_settings_note, style="Green.TLabel", wraplength=930).pack(fill="x", pady=(7, 0))

    def _cap_scale_changed(self, value):
        if self._syncing_visual or self.busy:
            return
        self.max_days_var.set(str(round(float(value))))

    def _refresh_visual_settings(self):
        if not hasattr(self, "attendance_calendar") or self._syncing_visual:
            return
        self._syncing_visual = True
        try:
            enabled = not self.busy and self.max_enabled_var.get()
            self.cap_scale.configure(state="normal" if enabled else "disabled")
            self.visual_cap_spin.configure(state="normal" if enabled else "disabled")
            try:
                count = int(self.max_days_var.get())
                self.cap_scale.set(min(31, max(1, count)))
                self.visual_cap_var.set(f"最大{count}日まで連続勤務" if self.max_enabled_var.get() else "連勤上限は設定しません")
            except ValueError:
                self.visual_cap_var.set("連勤上限は整数で入力してください")
            for box in self.visual_ratio_boxes:
                box.configure(state="normal" if not self.busy and self.people_source_var.get() == "曜日ごとの割合" else "disabled")
            try:
                settings = self._get_settings()
            except ValueError as exc:
                self.visual_settings_note.set(str(exc))
                return
            data = self.base_data or self.data
            self.attendance_calendar.set_data(data, settings)
            if data:
                self.visual_settings_note.set(f"職員{len(data.employees)}人。日付指定→週指定→曜日指定→Excel人数の順で適用。変更後は条件確認で人数と勤務日数の整合性を検証します。")
        finally:
            self._syncing_visual = False

    def set_attendance_ratio(self, day, percentage, *, scope="day"):
        if self.busy:
            return
        key = week_start(day) if scope == "week" else day
        collection = self.weekly_ratio_overrides if scope == "week" else self.date_ratio_overrides
        if percentage is None:
            collection.pop(key, None)
        else:
            value = percentage if isinstance(percentage, Decimal) else Decimal(str(percentage))
            if not value.is_finite() or not Decimal(0) <= value <= Decimal(100):
                raise ValueError("出勤率は0〜100％で指定してください。")
            collection[key] = value
        self._settings_changed()
        self._refresh_visual_settings()

    def clear_attendance_overrides(self):
        if not self.busy:
            self.weekly_ratio_overrides.clear()
            self.date_ratio_overrides.clear()
            self._settings_changed()
            self._refresh_visual_settings()

    def _attendance_dialog(self, day, scope):
        if self.busy:
            return
        try:
            settings = self._get_settings()
            if scope == "week":
                current = self.weekly_ratio_overrides.get(week_start(day))
                if current is None and settings.weekday_ratios is not None:
                    current = settings.weekday_ratios[0]
            else:
                current = attendance_ratio_for_day(day, settings)
        except ValueError as exc:
            self.status_var.set(f"先に勤務条件の入力を直してください：{exc}")
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("週の出勤率" if scope == "week" else "この日の出勤率")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=20)
        body.pack(fill="both", expand=True)
        text = f"{day:%Y/%m/%d}からの週全体" if scope == "week" else f"{day:%Y/%m/%d} のみ"
        ttk.Label(body, text=text, font=("Helvetica", 15, "bold")).pack(anchor="w", pady=(0, 8))
        value_var = tk.StringVar(value=str(current if current is not None else 50))
        scale = ttk.Scale(body, from_=0, to=100)
        scale.set(float(current) if current is not None else 50)
        scale.configure(command=lambda value: value_var.set(str(round(float(value)))))
        scale.pack(fill="x", pady=8)
        line = ttk.Frame(body)
        line.pack(fill="x")
        ttk.Entry(line, textvariable=value_var, width=9).pack(side="left")
        ttk.Label(line, text="％（職員数に対する勤務開始人数）").pack(side="left", padx=8)
        error = tk.StringVar(value="週の指定より、日付の指定を優先します。")
        ttk.Label(body, textvariable=error, wraplength=420).pack(anchor="w", pady=8)
        buttons = ttk.Frame(body)
        buttons.pack(fill="x")
        def apply(clear=False):
            try:
                self.set_attendance_ratio(day, None if clear else Decimal(value_var.get()), scope=scope)
            except (ValueError, InvalidOperation) as exc:
                error.set(str(exc))
                return
            dialog.destroy()
        ttk.Button(buttons, text="設定を解除", command=lambda: apply(True)).pack(side="left")
        ttk.Button(buttons, text="この率を設定", command=apply, style="Accent.TButton").pack(side="right")
        dialog.grab_set()

    def _build_detailed_settings(self, parent):
        self.settings_sections = ttk.Notebook(parent)
        self.settings_sections.pack(fill="x", pady=(0, 6))
        basic = ttk.Frame(self.settings_sections, padding=8)
        sequence = ttk.Frame(self.settings_sections, padding=8)
        leave = ttk.Frame(self.settings_sections, padding=8)
        timing = ttk.Frame(self.settings_sections, padding=8)
        self.settings_sections.add(basic, text="人数・連勤")
        self.settings_sections.add(sequence, text="夜勤後・勤務順序")
        self.settings_sections.add(leave, text="公休・有給")
        self.settings_sections.add(timing, text="勤務時刻・明け")
        options = ttk.Frame(basic)
        options.pack(fill="x", pady=(0, 8))
        ttk.Label(options, text="勤務モード").pack(side="left")
        self.mode_var = tk.StringVar(value=DAY_ONLY)
        self._control(ttk.Combobox(options, textvariable=self.mode_var,
                                  values=(DAY_ONLY, TWO_SHIFT, THREE_SHIFT), width=15,
                                  state="readonly"), "readonly").pack(side="left", padx=(8, 22))
        self.max_enabled_var = tk.BooleanVar(value=True)
        self._control(ttk.Checkbutton(options, text="連続勤務の上限を設定", variable=self.max_enabled_var)).pack(side="left")
        self.max_days_var = tk.StringVar(value="5")
        self.max_days_spin = self._control(ttk.Spinbox(options, from_=1, to=366,
                                                       width=5, textvariable=self.max_days_var))
        self.max_days_spin.pack(side="left", padx=(8, 4))
        ttk.Label(options, text="日").pack(side="left")
        source_line = ttk.Frame(basic)
        source_line.pack(fill="x", pady=(0, 6))
        ttk.Label(source_line, text="配置人数の決め方").pack(side="left")
        self.people_source_var = tk.StringVar(value="Excelの配置人数")
        self._control(ttk.Combobox(source_line, textvariable=self.people_source_var,
                                  values=("Excelの配置人数", "曜日ごとの割合"), width=20,
                                  state="readonly"), "readonly").pack(side="left", padx=10)
        ttk.Label(source_line, text="割合は職員数に対する勤務者の割合です。", wraplength=580).pack(side="left")
        ratios = ttk.LabelFrame(basic, text="曜日ごとの勤務割合（月〜日）", padding=(8, 5))
        ratios.pack(fill="x", pady=(0, 6))
        self.ratio_vars = [tk.StringVar(value="50") for _ in range(7)]
        self.ratio_boxes = []
        for index, weekday in enumerate(WEEKDAYS):
            item = ttk.Frame(ratios)
            item.grid(row=0, column=index, sticky="w", padx=(0, 12))
            ttk.Label(item, text=weekday).pack(side="left", padx=(0, 4))
            spin = self._control(ttk.Spinbox(item, from_=0, to=100, increment=1,
                                            width=5, textvariable=self.ratio_vars[index]))
            spin.pack(side="left")
            ttk.Label(item, text="％").pack(side="left")
            self.ratio_boxes.append(spin)
        weights = ttk.LabelFrame(basic, text="区分ごとの配分（相対比・未入力は1）", padding=(8, 5))
        weights.pack(fill="x", pady=(0, 6))
        self.weight_vars = {shift: tk.StringVar(value="1") for shift in ("日勤", "夜勤", "準夜勤", "深夜勤")}
        self.weight_boxes = {}
        self.weight_frames = {}
        for index, (shift, var) in enumerate(self.weight_vars.items()):
            item = ttk.Frame(weights)
            item.grid(row=0, column=index, sticky="w", padx=(0, 24))
            self.weight_frames[shift] = item
            ttk.Label(item, text=shift).pack(side="left", padx=(0, 6))
            spin = self._control(ttk.Spinbox(item, from_=0, to=1000, width=6, textvariable=var))
            spin.pack(side="left")
            self.weight_boxes[shift] = spin
        ttk.Label(basic,
                  text="割合からの総人数は小数点以下を切り上げ。区分別人数は配分比で割り振り、端数の大きい順に残りを配ります。\n"
                       "各人の勤務日数は自動変更しません。前期間の連勤は未入力なら0日と仮定。Excelの区分内訳列は選択モードと一致させます。",
                  wraplength=880).pack(anchor="w", pady=(0, 6))
        self.leave_enabled_var = tk.BooleanVar(value=False)
        self._control(ttk.Checkbutton(leave, text="公休・有給・明けを区別する", variable=self.leave_enabled_var)).pack(anchor="w", pady=(0, 12))
        count_row = ttk.Frame(leave)
        count_row.pack(fill="x", pady=(0, 12))
        ttk.Label(count_row, text="日数の指定方法").pack(side="left")
        self.day_count_mode_var = tk.StringVar(value=DIRECT_DAYS)
        self.day_count_box = self._control(ttk.Combobox(count_row, textvariable=self.day_count_mode_var,
            values=(DIRECT_DAYS, PUBLIC_DAYS), width=21, state="readonly"), "readonly")
        self.day_count_box.pack(side="left", padx=10)
        ttk.Label(count_row, text="対象期間の共通公休日数").pack(side="left", padx=(12, 0))
        self.public_holidays_var = tk.StringVar(value="8")
        self.public_holidays_spin = self._control(ttk.Spinbox(count_row, from_=0, to=366,
            width=6, textvariable=self.public_holidays_var))
        self.public_holidays_spin.pack(side="left", padx=8)
        ttk.Label(count_row, text="日（空欄なら個別指定）").pack(side="left")
        ttk.Label(leave,
            text="公休は対象期間の休日数です。『月間』ではありません。職員シートの個別公休日数が共通値より優先します。\n"
                 "勤務日数方式では勤務開始回数を入力し、公休日数は目標にしません。公休から算出する方式では勤務日数を独立入力しません。\n"
                 "有給は任意の『有給』シートへ確定日を入力。公休とは別集計で、公休希望の採用にも数えません。\n"
                 "日付をまたぐ勤務の翌日は明けです。明けは公休・有給に含めず、実勤務日として連勤に数えます。",
            wraplength=880, justify="left").pack(anchor="w", pady=(4, 10))
        ttk.Label(leave, text="確定休は公休として固定します。有給との重複や、明けを公休へ変更した案は保存できません。",
                  style="Green.TLabel", wraplength=880).pack(fill="x")
        self.time_include_vars = {}
        self.time_start_vars = {}
        self.time_end_vars = {}
        self.time_next_vars = {}
        self.time_widgets = {}
        times_grid = ttk.Frame(timing)
        times_grid.pack(fill="x")
        for col, text in enumerate(("使用する区分", "開始（時:分）", "終了（時:分）", "終了日")):
            ttk.Label(times_grid, text=text).grid(row=0, column=col, sticky="w", padx=(0, 16), pady=(0, 5))
        presets = {"日勤": ("08:00", "17:00", "当日"), "夜勤": ("16:00", "09:00", "翌日"),
                   "準夜勤": ("16:00", "24:00", "当日"), "深夜勤": ("00:00", "08:00", "当日")}
        for row, (label, (start_text, end_text, next_text)) in enumerate(presets.items(), 1):
            self.time_include_vars[label] = tk.BooleanVar(value=label == "日勤")
            self.time_start_vars[label] = tk.StringVar(value=start_text)
            self.time_end_vars[label] = tk.StringVar(value=end_text)
            self.time_next_vars[label] = tk.StringVar(value=next_text)
            include = self._control(ttk.Checkbutton(times_grid, text=label, variable=self.time_include_vars[label]))
            include.grid(row=row, column=0, sticky="w", pady=3)
            start_entry = self._control(ttk.Entry(times_grid, textvariable=self.time_start_vars[label], width=10))
            end_entry = self._control(ttk.Entry(times_grid, textvariable=self.time_end_vars[label], width=10))
            next_box = self._control(ttk.Combobox(times_grid, textvariable=self.time_next_vars[label],
                values=("当日", "翌日"), width=8, state="readonly"), "readonly")
            for col, widget in enumerate((start_entry, end_entry, next_box), 1):
                widget.grid(row=row, column=col, sticky="w", padx=(0, 16), pady=3)
            self.time_widgets[label] = (include, start_entry, end_entry, next_box)
        self.times_confirmed_var = tk.BooleanVar(value=False)
        self.times_confirmed_check = self._control(ttk.Checkbutton(timing,
            text="勤務時刻を実際の職場の条件として確認しました（記入例のまま確定しないでください）",
            variable=self.times_confirmed_var))
        self.times_confirmed_check.pack(anchor="w", pady=(8, 4))
        ttk.Label(timing,
            text="交代制で公休・有給・明けを区別する場合、開始・終了時刻の確認が必須です。翌09:00は終了09:00＋翌日。\n"
                 "24:00終了は当日24:00。前日区分がモード外の場合、その区分の時刻も使用してください。\n"
                 "跨ぎ勤務の翌日は明けとし、その日に新たな勤務を割り当てません。夜勤後休み設定を無効にしても明けは必要です。\n"
                 "時間数・休息時間・法令適合は自動判定しません。最終日の跨ぎ勤務は次期間へ明けとして引き継ぎます。",
            wraplength=880).pack(anchor="w", pady=(2, 0))
        self._last_time_mode = DAY_ONLY
        self._times_dirty = False
        self._last_leave_enabled = False
        self.night_rest_var = tk.BooleanVar(value=True)
        self.night_rest_check = self._control(ttk.Checkbutton(sequence,
            text="夜勤・深夜勤の翌日は休み（連続夜勤不可）", variable=self.night_rest_var))
        self.night_rest_check.pack(anchor="w", pady=(0, 8))
        transition_line = ttk.Frame(sequence)
        transition_line.pack(fill="x", pady=(0, 6))
        ttk.Label(transition_line, text="禁止する勤務順序：前日").pack(side="left")
        self.transition_from_var = tk.StringVar(value="日勤")
        self.transition_to_var = tk.StringVar(value=OFF)
        self.transition_from_box = self._control(ttk.Combobox(transition_line, textvariable=self.transition_from_var,
                                       state="readonly", width=10), "readonly")
        self.transition_from_box.pack(side="left", padx=6)
        ttk.Label(transition_line, text="→ 翌日").pack(side="left")
        self.transition_to_box = self._control(ttk.Combobox(transition_line, textvariable=self.transition_to_var,
                                       state="readonly", width=10), "readonly")
        self.transition_to_box.pack(side="left", padx=6)
        self._button(transition_line, "禁止を追加", self._add_transition).pack(side="left", padx=8)
        self._button(transition_line, "選択を削除", self._remove_transition).pack(side="left")
        self.forbidden_transitions = []
        self._last_mode = DAY_ONLY
        self.transition_tree = self._tree(sequence, ("前日の区分", "翌日の区分（禁止）"), height=2)
        ttk.Label(sequence,
                  text="登録した前日→翌日の組合せを禁止します。３交代へ切替時は「準夜勤→日勤」を追加。休みも指定できます。\n"
                       "高度モードの明けは勤務時刻から判定します。休息時間・法令適合の自動判定は行いません。期間末の勤務は次期間へ引き継ぎます。",
                  wraplength=880).pack(anchor="w", pady=(6, 0))
        commands = ttk.Frame(parent)
        commands.pack(fill="x", pady=(0, 6))
        self._button(commands, "勤務条件を保存", self.save_settings).pack(side="left")
        self._button(commands, "勤務条件を読込", self.load_settings).pack(side="left", padx=8)
        self._button(commands, "人数をプレビュー・条件確認", lambda: self.read_conditions(show_settings=True), True).pack(side="right")
        self.settings_preview_var = tk.StringVar(value="入力後、条件確認で日ごとの必要人数を表示します。")
        ttk.Label(parent, textvariable=self.settings_preview_var, wraplength=880).pack(anchor="w", pady=(0, 4))
        self.requirements_tree = self._tree(parent, ("日付", "割合", "配置人数"), height=3)
        for variable in (self.mode_var, self.max_enabled_var, self.max_days_var, self.people_source_var, self.night_rest_var,
                         self.leave_enabled_var, self.day_count_mode_var, self.public_holidays_var, self.times_confirmed_var,
                         *self.ratio_vars, *self.weight_vars.values()):
            variable.trace_add("write", self._settings_changed)
        for variable in (*self.time_include_vars.values(), *self.time_start_vars.values(),
                         *self.time_end_vars.values(), *self.time_next_vars.values()):
            variable.trace_add("write", self._times_changed)

    def _update_setting_controls(self):
        if not hasattr(self, "mode_var"):
            return
        ratio_enabled = not self.busy and self.people_source_var.get() == "曜日ごとの割合"
        for spin in self.ratio_boxes:
            spin.configure(state="normal" if ratio_enabled else "disabled")
        self.max_days_spin.configure(state="normal" if not self.busy and self.max_enabled_var.get() else "disabled")
        active = SHIFT_MODES.get(self.mode_var.get(), ())
        self.night_rest_check.configure(state="normal" if not self.busy else "disabled")
        advanced = self.leave_enabled_var.get()
        self.night_rest_check.configure(text="夜勤・深夜勤の翌日の勤務開始を禁止（跨ぎ勤務の翌日は設定に関係なく明け）"
                                        if advanced else "夜勤・深夜勤の翌日は休み（連続夜勤不可）")
        self.day_count_box.configure(state="readonly" if advanced and not self.busy else "disabled")
        self.public_holidays_spin.configure(state="normal" if advanced and not self.busy and self.day_count_mode_var.get() == PUBLIC_DAYS else "disabled")
        self.times_confirmed_check.configure(state="normal" if advanced and not self.busy else "disabled")
        for label, (include, start_entry, end_entry, next_box) in self.time_widgets.items():
            include.configure(state="normal" if advanced and not self.busy and label not in active else "disabled")
            enabled = advanced and not self.busy and self.time_include_vars[label].get()
            start_entry.configure(state="normal" if enabled else "disabled")
            end_entry.configure(state="normal" if enabled else "disabled")
            next_box.configure(state="readonly" if enabled else "disabled")
        labels = assignment_labels(self.mode_var.get(), advanced)
        for box, var in ((self.transition_from_box, self.transition_from_var),
                         (self.transition_to_box, self.transition_to_var)):
            box.configure(values=labels)
            if var.get() not in labels:
                var.set(labels[0])
        for shift, spin in self.weight_boxes.items():
            spin.configure(state="normal" if not self.busy and shift in active else "disabled")
            if shift in active:
                self.weight_frames[shift].grid()
            else:
                self.weight_frames[shift].grid_remove()

    def _settings_changed(self, *_args):
        if self._loading_settings or not self._settings_trace_ready or self.busy:
            return
        self._settings_from_file_pending = False
        mode = self.mode_var.get()
        advanced = self.leave_enabled_var.get()
        if advanced != self._last_leave_enabled:
            translated = []
            for a, b in self.forbidden_transitions:
                pair = tuple(PUBLIC_OFF if advanced and x == OFF else OFF if not advanced and x == PUBLIC_OFF else x for x in (a, b))
                if all(x in assignment_labels(mode, advanced) for x in pair):
                    translated.append(pair)
            self.forbidden_transitions = translated
            self._last_leave_enabled = advanced
            self._render_transitions()
        if mode != self._last_mode:
            allowed = assignment_labels(mode, advanced)
            self.forbidden_transitions = [pair for pair in self.forbidden_transitions if all(s in allowed for s in pair)]
            if mode == THREE_SHIFT and ("準夜勤", "日勤") not in self.forbidden_transitions:
                self.forbidden_transitions.append(("準夜勤", "日勤"))
            self._last_mode = mode
            self._render_transitions()
        if mode != self._last_time_mode:
            self._set_time_presets(mode)
        self._invalidate()
        self.requirements_tree.delete(*self.requirements_tree.get_children())
        self.settings_preview_var.set("勤務条件が変わりました。条件を再確認して算出人数を確認してください。")
        self.status_var.set("勤務条件を変更しました。「人数をプレビュー・条件確認」を押してください。")
        self._update_setting_controls()
        self._refresh_visual_settings()
        self._sync_workflow()

    def _set_time_presets(self, mode):
        self._loading_settings = True
        try:
            self._last_time_mode = mode
            self.time_start_vars["日勤"].set("08:00")
            self.time_end_vars["日勤"].set("16:00" if mode == THREE_SHIFT else "17:00")
            self.time_next_vars["日勤"].set("当日")
            for label in SHIFT_MODES[mode]:
                self.time_include_vars[label].set(True)
            self.times_confirmed_var.set(False)
        finally:
            self._loading_settings = False

    def _times_changed(self, *_args):
        if self._loading_settings or not self._settings_trace_ready or self.busy:
            return
        self._times_dirty = True
        self._loading_settings = True
        try:
            self.times_confirmed_var.set(False)
        finally:
            self._loading_settings = False
        self._settings_changed()

    @staticmethod
    def _parse_clock(text, *, allow_24=False):
        parts = text.strip().split(":")
        try:
            if len(parts) != 2:
                raise ValueError()
            hour, minute = map(int, parts)
            if not 0 <= minute <= 59 or not 0 <= hour <= (24 if allow_24 else 23) or (hour == 24 and minute != 0):
                raise ValueError()
        except ValueError:
            raise ValueError("勤務時刻は08:00の形式で入力してください。終了の24:00だけ使用できます。") from None
        return hour * 60 + minute

    def _render_transitions(self):
        self.transition_tree.delete(*self.transition_tree.get_children())
        for index, pair in enumerate(self.forbidden_transitions):
            self.transition_tree.insert("", "end", iid=str(index), values=pair)

    def _add_transition(self):
        if self.busy:
            return
        pair = (self.transition_from_var.get(), self.transition_to_var.get())
        if pair not in self.forbidden_transitions:
            self.forbidden_transitions.append(pair)
            self._render_transitions()
            self._settings_changed()

    def _remove_transition(self):
        if self.busy:
            return
        selected = self.transition_tree.selection()
        if selected:
            index = int(selected[0])
            del self.forbidden_transitions[index]
            self._render_transitions()
            self._settings_changed()

    def _get_settings(self):
        mode = self.mode_var.get()
        if mode not in SHIFT_MODES:
            raise ValueError("勤務モードを選択してください。")
        cap = None
        if self.max_enabled_var.get():
            try:
                cap = int(self.max_days_var.get())
            except ValueError:
                raise ValueError("連続勤務上限は1以上の整数で指定してください。") from None
            if not 1 <= cap <= 366:
                raise ValueError("連続勤務上限は1〜366日の整数で指定してください。")
        ratios = None
        if self.people_source_var.get() == "曜日ごとの割合":
            values = []
            for index, var in enumerate(self.ratio_vars):
                try:
                    value = Decimal(var.get().strip())
                except (InvalidOperation, ValueError):
                    raise ValueError(f"{WEEKDAYS[index]}曜の割合を0〜100の数値で指定してください。") from None
                if not value.is_finite() or not 0 <= value <= 100:
                    raise ValueError(f"{WEEKDAYS[index]}曜の割合を0〜100の数値で指定してください。")
                values.append(value)
            ratios = tuple(values)
        weights = []
        for shift in SHIFT_MODES[mode]:
            text = self.weight_vars[shift].get().strip()
            try:
                value = int(text) if text else 1
            except ValueError:
                raise ValueError(f"{shift}の配分を0以上の整数で指定してください。") from None
            if not 0 <= value <= 1_000_000:
                raise ValueError(f"{shift}の配分を0〜1000000の整数で指定してください。")
            weights.append(value)
        if not any(weights):
            raise ValueError("区分の配分をすべて0にはできません。少なくとも1区分を1以上にしてください。")
        advanced = self.leave_enabled_var.get()
        public_days = None
        if advanced and self.day_count_mode_var.get() == PUBLIC_DAYS and self.public_holidays_var.get().strip():
            try:
                public_days = int(self.public_holidays_var.get())
            except ValueError:
                raise ValueError("共通公休日数は対象期間に対する0以上の整数、または空欄にしてください。") from None
            if not 0 <= public_days <= 366:
                raise ValueError("共通公休日数は0〜366の整数にしてください。")
        times = []
        use_times = advanced and (mode != DAY_ONLY or self.times_confirmed_var.get() or self._times_dirty
                                 or any(v.get() for label, v in self.time_include_vars.items() if label != "日勤"))
        if use_times:
            for label, included in self.time_include_vars.items():
                if included.get():
                    start = self._parse_clock(self.time_start_vars[label].get())
                    end = self._parse_clock(self.time_end_vars[label].get(), allow_24=True)
                    if self.time_next_vars[label].get() == "翌日":
                        end += 1440
                    if not start < end <= start + 1440:
                        raise ValueError(f"{label}：終了は開始より後、勤務時間は24時間以内で指定してください。")
                    times.append(ShiftTime(label, start, end))
        return ShiftSettings(mode=mode, max_consecutive_days=cap, weekday_ratios=ratios,
            shift_weights=tuple(weights), night_rest=self.night_rest_var.get(),
            forbidden_transitions=tuple(self.forbidden_transitions), leave_enabled=advanced,
            day_count_mode=self.day_count_mode_var.get() if advanced else DIRECT_DAYS,
            default_public_holidays=public_days, shift_times=tuple(times),
            shift_times_confirmed=self.times_confirmed_var.get() if advanced else False,
            weekly_ratios=tuple(sorted(self.weekly_ratio_overrides.items())),
            date_ratios=tuple(sorted(self.date_ratio_overrides.items())))

    def _apply_settings(self, settings, *, invalidate=False):
        self._loading_settings = True
        try:
            self.weekly_ratio_overrides = dict(settings.weekly_ratios)
            self.date_ratio_overrides = dict(settings.date_ratios)
            self.mode_var.set(settings.mode)
            self._last_mode = settings.mode
            self._last_time_mode = settings.mode
            self._times_dirty = bool(settings.shift_times)
            self._last_leave_enabled = settings.leave_enabled
            self.leave_enabled_var.set(settings.leave_enabled)
            self.day_count_mode_var.set(settings.day_count_mode)
            self.public_holidays_var.set("" if settings.default_public_holidays is None else str(settings.default_public_holidays))
            # 読込のたびに全行を記入例へ戻す。以前のファイルの時刻を持ち越さない。
            presets = {"日勤": ("08:00", "16:00" if settings.mode == THREE_SHIFT else "17:00", "当日"),
                       "夜勤": ("16:00", "09:00", "翌日"), "準夜勤": ("16:00", "24:00", "当日"),
                       "深夜勤": ("00:00", "08:00", "当日")}
            for label, var in self.time_include_vars.items():
                var.set(label in SHIFT_MODES[settings.mode])
                start_text, end_text, following_text = presets[label]
                self.time_start_vars[label].set(start_text)
                self.time_end_vars[label].set(end_text)
                self.time_next_vars[label].set(following_text)
            for timing in settings.shift_times:
                self.time_include_vars[timing.label].set(True)
                self.time_start_vars[timing.label].set(f"{timing.start_minute // 60:02}:{timing.start_minute % 60:02}")
                following = timing.end_minute > 1440
                end = timing.end_minute - 1440 if following else timing.end_minute
                self.time_end_vars[timing.label].set(f"{end // 60:02}:{end % 60:02}")
                self.time_next_vars[timing.label].set("翌日" if following else "当日")
            supplied_labels = {timing.label for timing in settings.shift_times}
            active_times_complete = set(SHIFT_MODES[settings.mode]) <= supplied_labels
            # 未入力行に補った記入例へ、元データの確認済みフラグを流用しない。
            self.times_confirmed_var.set(settings.shift_times_confirmed and active_times_complete)
            self.night_rest_var.set(settings.night_rest)
            self.forbidden_transitions = list(settings.forbidden_transitions)
            self._render_transitions()
            self.max_enabled_var.set(settings.max_consecutive_days is not None)
            self.max_days_var.set(str(settings.max_consecutive_days or 5))
            self.people_source_var.set("曜日ごとの割合" if settings.weekday_ratios is not None else "Excelの配置人数")
            for var, value in zip(self.ratio_vars, settings.weekday_ratios or (Decimal(50),) * 7):
                var.set(str(value))
            for var in self.weight_vars.values():
                var.set("1")
            for shift, value in zip(SHIFT_MODES[settings.mode], settings.shift_weights):
                self.weight_vars[shift].set(str(value))
        finally:
            self._loading_settings = False
        self._update_setting_controls()
        if invalidate:
            self._settings_changed()

    def set_settings(self, settings):
        """スモーク検証・外部起動からも使える勤務条件設定。"""
        if not self.busy:
            self._apply_settings(settings, invalidate=True)

    def _render_requirement_preview(self, data):
        active = SHIFT_MODES[data.settings.mode]
        columns = ("日付", "割合", "配置人数", *active)
        self.requirements_tree.delete(*self.requirements_tree.get_children())
        self.requirements_tree.configure(columns=columns)
        for column in columns:
            self.requirements_tree.heading(column, text=column)
            self.requirements_tree.column(column, width=165 if column == "日付" else 120, minwidth=100)
        for requirement in sorted(data.requirements, key=lambda r: r.day):
            ratio = (f"{data.settings.weekday_ratios[requirement.day.weekday()]}％"
                     if data.settings.weekday_ratios is not None else "Excel")
            counts = dict(requirement.shift_counts)
            self.requirements_tree.insert("", "end", values=[display_day(requirement.day), ratio,
                requirement.required, *[counts.get(shift, requirement.required if len(active) == 1 else 0) for shift in active]])
        self.settings_preview_var.set(f"{data.settings.mode} ／ {len(data.employees)}人・{len(data.days)}日 ／ 配置人数合計 {sum(r.required for r in data.requirements)}人日")

    def save_settings(self, path=None):
        if self.busy:
            return
        try:
            settings = self._get_settings()
        except ValueError as exc:
            self._show_exception(exc, preserve_result=True, operation="settings")
            return
        if path is None:
            path = filedialog.asksaveasfilename(parent=self.root, title="勤務条件を保存", defaultextension=".json",
                                               initialfile="sinta_shift_勤務条件.json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        payload = {"version": 1, "settings": settings_to_payload(settings)}
        def task():
            Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            return Path(path)
        self._run_worker(task, lambda p: self.status_var.set(f"勤務条件を保存しました：{p}"), "勤務条件を保存しています…",
                         preserve_result_on_error=True, operation="settings_save")

    def load_settings(self, path=None):
        if self.busy:
            return
        if path is None:
            path = filedialog.askopenfilename(parent=self.root, title="勤務条件を選択", filetypes=[("JSON", "*.json")])
        if not path:
            return
        def task():
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
            if not isinstance(payload, dict) or payload.get("version") != 1:
                raise ValueError("勤務条件設定の形式・バージョンが未対応です。")
            return settings_from_payload(payload["settings"])
        def done(settings):
            self.set_settings(settings)
            self.notebook.select(self.settings_tab)
            self.status_var.set("勤務条件を復元しました。入力の条件確認を再実行してください。")
        self._run_worker(task, done, "保存された勤務条件を読み込んでいます…", preserve_result_on_error=True, operation="settings_load")

    def _build_manager(self):
        ttk.Label(self.manager_tab, text="有給希望を見て、公休・有給の配分を決めます。",
                  font=("Helvetica", 15, "bold")).pack(anchor="w", pady=(0, 5))
        ttk.Label(self.manager_tab,
            text="スライダーは案の表示だけです。有給残高・承認ルールを外部の台帳等で確認してから、配分を確定してください。",
            wraplength=930).pack(anchor="w", pady=(0, 7))
        self.manager_summary_var = tk.StringVar(value="先にExcelを読み込んでください。")
        ttk.Label(self.manager_tab, textvariable=self.manager_summary_var, style="Green.TLabel", wraplength=930).pack(fill="x", pady=(0, 6))
        mix = ttk.Frame(self.manager_tab, padding=10, style="Card.TFrame")
        mix.pack(fill="x", pady=(0, 8))
        line = ttk.Frame(mix, style="Card.TFrame")
        line.pack(fill="x")
        ttk.Label(line, text="未決定の有給希望から、有給にする件数", style="Card.TLabel").pack(side="left")
        self.mix_paid_count_var = tk.StringVar(value="0")
        self.mix_count_entry = self._control(ttk.Spinbox(line, from_=0, to=999, width=6, textvariable=self.mix_paid_count_var))
        self.mix_count_entry.pack(side="left", padx=10)
        self.mix_apply_button = self._button(line, "この配分を確定", self.apply_paid_mix, True)
        self.mix_apply_button.pack(side="right")
        self.mix_scale = self._control(ttk.Scale(mix, from_=0, to=1, command=self._mix_scale_changed))
        self.mix_scale.pack(fill="x", pady=(4, 0))
        self.mix_bar = MixBar(mix)
        self.mix_bar.pack(fill="x")
        self.mix_hint_var = tk.StringVar(value="案を表示しても確定状態は変わりません。")
        ttk.Label(mix, textvariable=self.mix_hint_var, style="Card.TLabel", wraplength=880).pack(anchor="w")
        self._syncing_mix = False
        self._last_mix_pending = None
        self._planned_paid_ids, self._planned_public_ids = set(), set()
        self.mix_paid_count_var.trace_add("write", lambda *_args: self._refresh_mix_preview())
        commands = ttk.Frame(self.manager_tab)
        commands.pack(fill="x", pady=(0, 7))
        self.confirm_paid_button = self._button(commands, "選択を有給に", lambda: self.set_paid_request_confirmation(confirmed=True), True)
        self.confirm_paid_button.pack(side="left")
        self.convert_public_button = self._button(commands, "選択を公休希望に", lambda: self.set_leave_choice(choice="public"))
        self.convert_public_button.pack(side="left", padx=6)
        self.unconfirm_paid_button = self._button(commands, "選択を保留に戻す", lambda: self.set_paid_request_confirmation(confirmed=False))
        self.unconfirm_paid_button.pack(side="left")
        self.select_paid_button = self._button(commands, "すべて選択", self.select_all_paid_requests)
        self.select_paid_button.pack(side="left", padx=6)
        self.paid_manager_tree = self._tree(self.manager_tab,
            ("申請ID", "職員ID", "氏名", "申請日", "今回の配分案", "管理者の確認状態"), height=6)
        for column in self.paid_manager_tree.cget("columns"):
            self.paid_manager_tree.column(column, width=160, minwidth=105)
        self.paid_manager_tree.configure(displaycolumns=("職員ID", "氏名", "申請日", "今回の配分案", "管理者の確認状態"))

    def _paid_requests(self):
        return [request for request in self.base_data.independent_requests
                if (request.original_leave_type or request.leave_type) == PAID_REQUEST] if self.base_data else []

    def _refresh_mix_preview(self):
        if not hasattr(self, "mix_scale") or self._syncing_mix:
            return
        requests = self._paid_requests()
        undecided = [request for request in requests if request.request_id not in self.confirmed_paid_request_ids
                     and request.request_id not in self.converted_public_request_ids]
        pending = len(undecided)
        self._syncing_mix = True
        try:
            if pending != self._last_mix_pending:
                previous = self._last_mix_pending
                text = self.mix_paid_count_var.get()
                count = int(text) if text.isdigit() else pending
                count = pending if previous is None or count == previous else min(count, pending)
                self.mix_paid_count_var.set(str(count))
                self._last_mix_pending = pending
            try:
                count = int(self.mix_paid_count_var.get())
                if not 0 <= count <= pending:
                    raise ValueError()
            except ValueError:
                self.mix_hint_var.set(f"有給にする件数を0〜{pending}の整数で入力してください。")
                self.mix_apply_button.configure(state="disabled")
                return
            self.mix_scale.configure(to=max(1, pending), state="normal" if pending and not self.busy else "disabled")
            self.mix_count_entry.configure(to=pending, state="normal" if pending and not self.busy else "disabled")
            self.mix_scale.set(count)
            if self.base_data:
                prepared = set_leave_choices(self.base_data, self.confirmed_paid_request_ids, self.converted_public_request_ids)
                paid_ids, public_ids = plan_paid_mix(prepared, count)
                self._planned_paid_ids, self._planned_public_ids = set(paid_ids), set(public_ids)
            else:
                self._planned_paid_ids, self._planned_public_ids = set(), set()
            self.mix_bar.set_counts(len(self._planned_public_ids), len(self._planned_paid_ids))
            self.mix_hint_var.set(f"今回の案：公休希望 {pending - count}件・有給 {count}件。確定済みの配分と元の有給シートは変更しません。")
            self.mix_apply_button.configure(state="normal" if pending and not self.busy else "disabled")
            self._render_paid_rows()
        except InputError as exc:
            self.mix_hint_var.set(str(exc))
            self.mix_apply_button.configure(state="disabled")
        finally:
            self._syncing_mix = False

    def _mix_scale_changed(self, value):
        if not self._syncing_mix and not self.busy:
            self.mix_paid_count_var.set(str(round(float(value))))

    def _render_paid_rows(self):
        self.paid_manager_tree.delete(*self.paid_manager_tree.get_children())
        self._paid_request_by_iid = {}
        names = {employee.id: employee.name for employee in self.base_data.employees} if self.base_data else {}
        for index, request in enumerate(self._paid_requests()):
            iid = f"paid{index}"
            self._paid_request_by_iid[iid] = request.request_id
            state = "有給確定" if request.request_id in self.confirmed_paid_request_ids else "公休希望へ変更" if request.request_id in self.converted_public_request_ids else "確認待ち"
            proposal = "—" if state != "確認待ち" else "公休希望（案）" if request.request_id in self._planned_public_ids else "有給（案）"
            self.paid_manager_tree.insert("", "end", iid=iid,
                values=[request.request_id, request.employee_id, names.get(request.employee_id, ""), request.day.isoformat(), proposal, state])

    def _render_paid_manager(self):
        requests = self._paid_requests()
        paid_count = sum(request.request_id in self.confirmed_paid_request_ids for request in requests)
        public_count = sum(request.request_id in self.converted_public_request_ids for request in requests)
        original_count = sum(not paid.request_id for paid in self.base_data.paid_leaves) if self.base_data else 0
        self.manager_summary_var.set(f"元の有給希望 {len(requests)}件 ／ 有給確定 {paid_count}件 ／ 公休へ変更 {public_count}件 ／ 保留 {len(requests) - paid_count - public_count}件。元の確定有給 {original_count}件はそのままです。")
        self._refresh_mix_preview()
        self._render_paid_rows()

    def select_all_paid_requests(self):
        if not self.busy:
            self.paid_manager_tree.selection_set(self.paid_manager_tree.get_children())

    def apply_paid_mix(self):
        if self.busy or self.base_data is None:
            return
        if not self.leave_enabled_var.get():
            self.status_var.set("詳しい設定の『公休・有給・明けを区別する』を有効にしてください。")
            self.notebook.select(self.advanced_settings_tab)
            return
        self._refresh_mix_preview()
        if str(self.mix_apply_button.cget("state")) == "disabled":
            return
        self.confirmed_paid_request_ids = set(self._planned_paid_ids)
        self.converted_public_request_ids = set(self._planned_public_ids)
        self._invalidate()
        self._render_paid_manager()
        self.read_conditions(show_manager=True)

    def set_leave_choice(self, request_ids=None, *, choice="paid"):
        if self.busy or self.base_data is None:
            return
        if choice != "pending" and not self.leave_enabled_var.get():
            self.status_var.set("配分を決める前に、詳しい設定の公休・有給区分を有効にしてください。")
            self.notebook.select(self.advanced_settings_tab)
            return
        if request_ids is None:
            request_ids = [self._paid_request_by_iid[iid] for iid in self.paid_manager_tree.selection()]
        if choice not in ("paid", "public", "pending"):
            raise ValueError("配分は有給・公休・保留から選んでください。")
        selected = set(request_ids) & {request.request_id for request in self._paid_requests()}
        if not selected:
            self.status_var.set("配分を変える希望を一覧から選択してください。")
            return
        self.confirmed_paid_request_ids.difference_update(selected)
        self.converted_public_request_ids.difference_update(selected)
        if choice == "paid":
            self.confirmed_paid_request_ids.update(selected)
        elif choice == "public":
            self.converted_public_request_ids.update(selected)
        elif choice != "pending":
            raise ValueError("配分は有給・公休・保留から選んでください。")
        self._invalidate()
        self._render_paid_manager()
        self.read_conditions(show_manager=True)

    def set_paid_request_confirmation(self, request_ids=None, *, confirmed=True):
        self.set_leave_choice(request_ids, choice="paid" if confirmed else "pending")

    def _build_conditions(self):
        self.conditions_var = tk.StringVar(value="入力を読み込み、「条件を確認」を押してください。")
        ttk.Label(self.conditions_tab, textvariable=self.conditions_var,
                  font=("Helvetica", 14), wraplength=1060, justify="left").pack(anchor="w", pady=(5, 20))
        ttk.Label(self.conditions_tab,
                  text="必須条件：日別人数、選択した勤務日数・公休日数の条件、確定休・有給を守ります。\n"
                       "希望休：まず採用申請数を最大化し、その数を保って優先順位合計を最小化します。\n"
                       "同一申請の候補日は代替候補です。同じ職員・同じ休みを二重計上しません。",
                  wraplength=1060, justify="left").pack(anchor="w", pady=(0, 22))
        options = ttk.Frame(self.conditions_tab)
        options.pack(fill="x")
        ttk.Label(options, text="全体の制限時間（秒）").pack(side="left")
        self.time_var = tk.StringVar(value="30")
        self._control(ttk.Spinbox(options, from_=0.1, to=300, increment=1, textvariable=self.time_var,
                                 width=8)).pack(side="left", padx=10)
        self.solve_button = self._button(options, "シフト案を作る", self.start_solve, True)
        self.solve_button.pack(side="left", padx=10)
        ttk.Label(self.conditions_tab,
                  text="区分別人数・連勤上限・夜勤後休み・禁止した勤務順序を検証します。前期間の情報も入力できます。\n"
                       "高度モードは公休・有給・明けを区別。明けも実勤務として連勤に数えます。法令適合・休息時間は自動判定しません。",
                  style="Green.TLabel", wraplength=1060, justify="left").pack(fill="x", pady=(28, 0))

    def _build_results(self):
        self.result_summary = tk.StringVar(value="担当者確認用のシフト案：未作成")
        self.result_summary_label = ttk.Label(self.result_tab, textvariable=self.result_summary, wraplength=930,
                  font=("Helvetica", 13, "bold"))
        self.result_summary_label.pack(anchor="w", pady=(0, 4))
        switch = ttk.Frame(self.result_tab)
        self.result_switch_frame = switch
        switch.pack(fill="x", pady=(3, 5))
        self.schedule_view_button = self._button(switch, "勤務表", lambda: self.show_result_view("schedule"), True)
        self.schedule_view_button.pack(side="left", padx=(0, 8))
        self.requests_view_button = self._button(switch, "希望休カレンダー", lambda: self.show_result_view("requests"))
        self.requests_view_button.pack(side="left")
        self.heatmap_teaser_var = tk.StringVar(value="案が未作成でも、読込済みの希望が多い日を確認できます。")
        self.heatmap_teaser_label = ttk.Label(self.result_tab, textvariable=self.heatmap_teaser_var, wraplength=930)
        self.heatmap_teaser_label.pack(anchor="w", pady=(0, 5))
        self.result_views = ttk.Notebook(self.result_tab, style="Workflow.TNotebook")
        self.result_views.pack(fill="both", expand=True)
        self.schedule_view = ttk.Frame(self.result_views)
        self.requests_view = ttk.Frame(self.result_views)
        self.result_views.add(self.schedule_view, text="勤務表")
        self.result_views.add(self.requests_view, text="希望休カレンダー")
        self.request_heatmap = RequestHeatmap(self.requests_view)
        self.request_heatmap.pack(fill="both", expand=True)
        ttk.Label(self.schedule_view,
                  text="日付セルをダブルクリックして区分を選択。Spaceで順に切替。⚠は必須条件違反のセルです。",
                  wraplength=1100).pack(anchor="w", pady=(0, 4))
        week_line = ttk.Frame(self.schedule_view)
        week_line.pack(fill="x", pady=(2, 7))
        self._button(week_line, "‹ 前の週", lambda: self.show_result_week(self.result_week_index - 1)).pack(side="left")
        self.result_week_var = tk.StringVar(value="シフト案を作ると、1週ずつ表示します")
        ttk.Label(week_line, textvariable=self.result_week_var, font=("Helvetica", 12, "bold")).pack(side="left", padx=12)
        self._button(week_line, "次の週 ›", lambda: self.show_result_week(self.result_week_index + 1)).pack(side="left")
        self.show_all_days_var = tk.BooleanVar(value=False)
        self._control(ttk.Checkbutton(week_line, text="全期間を表示", variable=self.show_all_days_var,
                                     command=self._update_result_week)).pack(side="right")
        # 保存導線を先に確保し、表は残りの高さでスクロール表示する。
        save_line = ttk.Frame(self.schedule_view)
        save_line.pack(side="bottom", fill="x", pady=(10, 0))
        self.save_button = self._button(save_line, "結果を保存", self.save_result, True)
        self.save_button.pack(side="right")
        self.save_hint = tk.StringVar(value="必須条件をすべて満たす案のみExcelに保存できます。")
        ttk.Label(save_line, textvariable=self.save_hint, wraplength=760).pack(side="left")
        tables = ttk.Panedwindow(self.schedule_view, orient="vertical")
        tables.pack(fill="both", expand=True)
        schedule_frame = ttk.Frame(tables)
        requests_frame = ttk.Frame(tables)
        tables.add(schedule_frame, weight=3)
        tables.add(requests_frame, weight=2)
        self.result_tree = self._tree(schedule_frame, (), height=5)
        self.result_tree.bind("<Double-1>", self._on_result_double_click)
        self.result_tree.bind("<Button-1>", self._remember_cell)
        self.result_tree.bind("<Button-3>", self._show_shift_menu)
        self.result_tree.bind("<Control-Button-1>", self._show_shift_menu)
        self.result_tree.bind("<space>", self._on_space)
        self._selected_cell = None
        ttk.Label(requests_frame, text="希望の結果（公休の採否・有給確定・従来形式の第2希望も表示）").pack(anchor="w", pady=(7, 4))
        self.request_tree = self._tree(requests_frame,
                                      ("申請ID", "職員ID", "採用日", "優先順位", "結果"), height=3)

    def show_result_view(self, view):
        if view not in ("schedule", "requests"):
            raise ValueError("表示はscheduleまたはrequestsを指定してください。")
        if self.busy:
            return
        self.current_result_view = view
        if view == "requests":
            self.result_summary_label.pack_forget()
            self.heatmap_teaser_label.pack_forget()
        else:
            self.result_summary_label.pack(anchor="w", pady=(0, 4), before=self.result_switch_frame)
            self.heatmap_teaser_label.pack(anchor="w", pady=(0, 5), before=self.result_views)
        self.result_views.select(self.requests_view if view == "requests" else self.schedule_view)
        self.schedule_view_button.configure(style="Accent.TButton" if view == "schedule" else "TButton")
        self.requests_view_button.configure(style="Accent.TButton" if view == "requests" else "TButton")
        self.notebook.select(self.result_tab)
        if view == "requests" and not self._heatmap_report:
            if self.base_data is not None:
                self._refresh_input_heatmap()
            elif self.input_path is not None:
                self.read_conditions(show_requests=True)

    def _set_request_heatmap(self, report, *, context, reset_selection=False):
        self._heatmap_report = tuple(report)
        if reset_selection:
            self.request_heatmap.month = None
            self.request_heatmap.selected_day = None
        self.request_heatmap.set_summaries(self._heatmap_report, context=context)
        period = [item for item in self._heatmap_report if item.in_period]
        if not period:
            self.heatmap_teaser_var.set("希望休カレンダー：集計できる対象期間はありません。")
            return
        peak = max(period, key=lambda item: item.request_employee_count)
        missing = {entry.request_id for item in period for entry in item.entries if entry.status == "未採用"}
        result_note = f"未採用 {len(missing)}申請" if any(item.adopted_employee_count is not None for item in period) else "採否は未計算"
        reference = "（参考）" if any(item.is_reference for item in period) and any(item.adopted_employee_count is not None for item in period) else ""
        self.heatmap_teaser_var.set(f"希望最多 {peak.day:%m/%d}：{peak.request_employee_count}人 ｜ {result_note}{reference}。希望休カレンダーで申請者を確認できます。")

    def _cancel_heatmap_refresh(self):
        if self._heatmap_after is not None:
            self.root.after_cancel(self._heatmap_after)
            self._heatmap_after = None

    def _schedule_heatmap_refresh(self):
        self._cancel_heatmap_refresh()
        if self.base_data is not None:
            self._heatmap_after = self.root.after(250, self._refresh_input_heatmap)

    def _refresh_input_heatmap(self):
        # 手動切替から即時更新した場合も、既存debounceを残さない。
        self._cancel_heatmap_refresh()
        if self._closed or self.base_data is None or self.schedule is not None:
            return
        if self.busy:
            self._heatmap_after = self.root.after(200, self._refresh_input_heatmap)
            return
        try:
            settings = self._get_settings()
        except ValueError as exc:
            self.request_heatmap.clear(f"シフトは未作成。勤務条件の入力を直してください：{exc}")
            return
        raw = self.base_data
        paid_ids, public_ids = frozenset(self.confirmed_paid_request_ids), frozenset(self.converted_public_request_ids)
        def task():
            prepared = raw
            fallback = False
            try:
                prepared = set_leave_choices(raw, paid_ids, public_ids)
                configured = configure_input(prepared, settings)
            except InputError:
                configured = replace(prepared, settings=settings)
                fallback = True
            return configured, build_request_heatmap(configured, raw_fallback=fallback)
        def done(payload):
            configured, report = payload
            self._heatmap_data = configured
            self._set_request_heatmap(report, context="シフトは未作成。希望人数の入力集計です（採否は未計算）。")
            self.status_var.set("希望休カレンダーの入力集計を更新しました。案の作成は条件確認から進めてください。")
        self._run_worker(task, done, "希望休カレンダーの入力を更新しています…", preserve_result_on_error=True)

    def _build_errors(self):
        self.error_summary_var = tk.StringVar(value="確認が必要な項目はありません。")
        ttk.Label(self.errors_tab, textvariable=self.error_summary_var,
                  style="Green.TLabel", wraplength=930).pack(fill="x", pady=(0, 7))
        panes = ttk.Panedwindow(self.errors_tab, orient="vertical")
        panes.pack(fill="both", expand=True)
        table_frame = ttk.Frame(panes)
        detail_frame = ttk.Frame(panes, padding=(0, 7, 0, 0))
        panes.add(table_frame, weight=2)
        panes.add(detail_frame, weight=4)
        self.error_tree = self._tree(table_frame,
                                    ("ファイル名", "シート名", "元行", "職員ID", "日付", "理由"), height=3)
        self.error_tree.column("理由", width=430, minwidth=260)
        self.error_tree.column("元行", width=65, minwidth=55, stretch=False)
        self.error_tree.column("日付", width=115, minwidth=100, stretch=False)
        self.error_tree.bind("<<TreeviewSelect>>", self._focus_issue)
        self.error_title_var = tk.StringVar(value="一覧の行を選ぶと、直し方を表示します。")
        self.error_location_var = tk.StringVar(value="")
        self.error_reason_var = tk.StringVar(value="")
        self.error_steps_var = tk.StringVar(value="")
        self.error_next_var = tk.StringVar(value="")
        ttk.Label(detail_frame, textvariable=self.error_title_var,
                  font=("Helvetica", 13, "bold"), wraplength=930).pack(anchor="w", pady=(0, 5))
        actions = ttk.Frame(detail_frame)
        actions.pack(fill="x", pady=(0, 6))
        self.error_action_button = self._button(actions, "修正する場所を開く", self.error_action, True)
        self.error_action_button.pack(side="left")
        self.error_retry_button = self._button(actions, "修正した入力を再読込", self.retry_after_error)
        self.error_retry_button.pack(side="left", padx=8)
        raw = ttk.Frame(detail_frame)
        raw.pack(side="bottom", fill="x", pady=(6, 0))
        ttk.Label(raw, text="エラーの全文（技術情報・選択してコピーできます）", font=("Helvetica", 10)).pack(anchor="w", pady=(0, 3))
        self.error_raw_text = self._readonly_error_text(raw, height=2, font_size=11)
        friendly = ttk.Frame(detail_frame)
        friendly.pack(fill="both", expand=True)
        self.error_detail_text = self._readonly_error_text(friendly, height=8, font_size=12)
        self.error_detail_text.tag_configure("heading", foreground=NAVY, font=("Helvetica", 12, "bold"))
        self._set_error_text(self.error_detail_text, "一覧の行を選んでください。選ぶだけでは別の画面へ移動しません。")

    def _readonly_error_text(self, parent, *, height, font_size):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, height=height, wrap="word", font=("Helvetica", font_size),
                       bg="white", fg=INK, relief="flat", padx=9, pady=7, state="disabled")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return text

    @staticmethod
    def _set_error_text(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")
        widget.yview_moveto(0)

    def _tree(self, parent, columns, height):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=height)
        y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, minwidth=100, anchor="w")
        return tree

    def _run_worker(self, task: Callable, callback: Callable, message: str, *, preserve_result_on_error=False, operation=None):
        if self.busy:
            return
        self.busy = True
        self.status_var.set(message)
        self.progress.start(12)
        self._update_buttons()
        def work():
            try:
                value = task()
                self._queue.put((True, value, callback, preserve_result_on_error, operation))
            except Exception as exc:
                self._queue.put((False, exc, callback, preserve_result_on_error, operation))
        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        if self._closed:
            return
        try:
            ok, value, callback, preserve_result, operation = self._queue.get_nowait()
        except queue.Empty:
            self.root.after(70, self._poll)
            return
        self.busy = False
        self.progress.stop()
        try:
            if ok:
                callback(value)
            else:
                self._show_exception(value, preserve_result=preserve_result, operation=operation)
        except Exception as exc:
            self._show_exception(exc, operation=operation)
        self._update_buttons()
        self.root.after(70, self._poll)

    def _update_buttons(self):
        for widget, normal_state in self._controls:
            widget.configure(state="disabled" if self.busy else normal_state)
        if not self.busy:
            self.check_button.configure(state="normal" if self.input_path else "disabled")
            self.solve_button.configure(state="normal" if self.data else "disabled")
            valid = (self.schedule is not None and self.evaluation is not None and self.evaluation.valid)
            self.save_button.configure(state="normal" if valid else "disabled")
        manager_state = "normal" if not self.busy and self._paid_requests() else "disabled"
        for button in (self.confirm_paid_button, self.convert_public_button, self.unconfirm_paid_button, self.select_paid_button):
            button.configure(state=manager_state)
        self._update_setting_controls()
        self._sync_workflow()
        self._refresh_visual_settings()
        self._refresh_mix_preview()
        self.request_heatmap.set_busy(self.busy)
        self._update_error_buttons()
        for role in OPTIONAL_SHEET_FIELDS:
            unused = self.mapping_vars[role]["sheet"].get() == UNUSED_COLUMN
            for widget in (self.mapping_boxes[role]["header"], *self.mapping_boxes[role]["columns"].values()):
                if self.busy or unused:
                    widget.configure(state="disabled")
        selected_role_tab = self.role_notebook.select()
        for tab in self.role_notebook.tabs():
            self.role_notebook.tab(tab, state="disabled" if self.busy and tab != selected_role_tab else "normal")

    def _invalidate_mapping(self):
        if self.busy:
            return
        # 列・シート・見出しを変えた後、前の対応で読んだrawを再集計しない。
        self.base_data = None
        self._heatmap_data = None
        self._render_paid_manager()
        self._invalidate()
        self.mapping_summary_var.set("列対応を変更しました。条件確認でファイルを読み直すまで入力集計は更新しません。")

    def _invalidate(self):
        if self.busy:
            return
        self.data = None
        self._clear_result()
        self.conditions_var.set("入力条件が変更されました。「条件を確認」を実行してください。")
        self.status_var.set("列対応を確認し、「条件を確認」を押してください。")
        self._update_buttons()
        self._schedule_heatmap_refresh()

    def _clear_result(self):
        self._cancel_heatmap_refresh()
        self._heatmap_report = ()
        self.request_heatmap.clear("シフト案は未作成です。入力の希望を再確認・更新してください。")
        self.heatmap_teaser_var.set("シフト案は未作成。希望休カレンダーで読込済みの希望を確認できます。")
        self.result = None
        self.schedule = None
        self.evaluation = None
        self.edited = False
        self.result_week_index = 0
        self._selected_cell = None
        self.result_summary.set("担当者確認用のシフト案：未作成")
        self.result_tree.delete(*self.result_tree.get_children())
        self.request_tree.delete(*self.request_tree.get_children())
        self.saved_var.set("")
        self.last_saved = None
        self.save_hint.set("必須条件をすべて満たす案のみExcelに保存できます。")

    def choose_file(self):
        if self.busy:
            return
        options = {"parent": self.root, "title": "Excel入力（.xlsx）を選択"}
        # Aqua/Tkのnative絞り込みで正常な.xlsxを選択できない環境に対応する。
        # ファイル選択後はload_fileが拡張子を検証し、対象外を読み込まない。
        if self.root.tk.call("tk", "windowingsystem") != "aqua":
            options["filetypes"] = [("Excel（.xlsx）", "*.xlsx")]
        path = filedialog.askopenfilename(**options)
        if path:
            self.load_file(path)

    def load_file(self, path, *, sample=False):
        if self.busy:
            return
        self._sample_pending = sample
        path = Path(path).resolve()
        if path.suffix.lower() != ".xlsx":
            self._show_exception(ValueError("入力は.xlsx形式のみです。.xls・マクロ・CSV・PDFは対象外です。"), operation="read")
            return
        self.input_path = path
        self.data = None
        self.base_data = None
        self._heatmap_data = None
        self._heatmap_has_computed = False
        self.confirmed_paid_request_ids.clear()
        self.converted_public_request_ids.clear()
        self._last_mix_pending = None
        self._render_paid_manager()
        self._settings_from_file_pending = True
        self._clear_result()
        self.path_var.set(f"入力：{path}")
        self._show_issues([])
        def task():
            names = inspect_workbook(path)
            previews = {}
            for role in INPUT_FIELDS:
                if role in OPTIONAL_SHEET_FIELDS and role not in names:
                    previews[role] = (UNUSED_COLUMN, None)
                    continue
                sheet = role if role in names else names[0]
                previews[role] = (sheet, preview_sheet(path, sheet, 1))
            return names, previews
        self._run_worker(task, self._on_workbook_loaded, "入力Excelのシートと先頭20行を読み込んでいます…", operation="read")

    def _on_workbook_loaded(self, payload):
        names, previews = payload
        self._sheet_names = names
        for role, values in self.mapping_vars.items():
            sheet, frame = previews[role]
            values["sheet"].set(sheet)
            values["header_row"].set(1)
            self.mapping_boxes[role]["sheet"].configure(values=[UNUSED_COLUMN, *names] if role in OPTIONAL_SHEET_FIELDS else names)
            self._apply_columns(role, frame.columns if frame is not None else (), auto_format=role == "希望休")
        role = self._roles[self.role_notebook.index("current")]
        self._show_preview(role, previews[role][1])
        self.conditions_var.set("先頭20行を確認してから、人数と連勤へ進んでください。")
        needed = matched = 0
        for role, fields in INPUT_FIELDS.items():
            if role in OPTIONAL_SHEET_FIELDS and self.mapping_vars[role]["sheet"].get() == UNUSED_COLUMN:
                continue
            selected = self._request_fields() if role == "希望休" else fields
            for field in selected:
                if role == "職員" and field == WORK_DAYS_FIELD:
                    continue
                needed += 1
                matched += bool(self.mapping_vars[role]["columns"][field].get())
        self.mapping_summary_var.set(f"列の自動対応 {matched}/{needed}項目。" +
            ("先頭20行を確認して次へ進んでください。" if matched == needed else "未対応があります。詳しい設定を開いて元の列を選んでください。"))
        if matched != needed:
            self.import_details.show()
        self.status_var.set("読み込み完了。先頭20行・列対応と「勤務条件」を確認してください。")
        self.notebook.select(self.import_tab)
        if getattr(self, "_sample_pending", False):
            self._sample_pending = False
            self.root.after(20, self.read_conditions)

    def _request_format(self):
        return REQUEST_FORMAT_LABELS[self.request_format_var.get()]

    def _request_fields(self, format_name=None):
        chosen = self._request_format() if format_name is None else format_name
        if chosen == "auto":
            chosen = self._detected_request_format
        return WIDE_REQUEST_FIELDS if chosen == "independent" else REQUIRED_FIELDS["希望休"]

    def _layout_request_fields(self):
        for item in self.request_field_frames.values():
            item.grid_remove()
        for index, field in enumerate(self._request_fields()):
            self.request_field_frames[field].grid(row=index // 4, column=index % 4,
                sticky="ew", padx=(0, 10), pady=(0, 5))

    def _request_format_changed(self, _event=None):
        if self.busy:
            return
        self._layout_request_fields()
        self._apply_columns("希望休", self._columns.get("希望休", ()), preserve=True)
        self._invalidate_mapping()

    def _apply_columns(self, role, columns, preserve=False, auto_format=False):
        headers = [str(c) for c in columns]
        normalized = {unicodedata.normalize("NFKC", header).strip(): header for header in headers}
        self._columns[role] = headers
        if role == "希望休":
            wide = set(WIDE_REQUEST_FIELDS) <= set(normalized)
            alternatives = set(REQUIRED_FIELDS[role]) <= set(normalized)
            detected = "independent" if wide and not alternatives else "alternatives" if alternatives and not wide else None
            if detected:
                self._detected_request_format = detected
                if auto_format or self._request_format() == "auto":
                    self.request_format_var.set(REQUEST_FORMAT_NAMES[detected])
            elif auto_format:
                self.request_format_var.set(REQUEST_FORMAT_NAMES["auto"])
            self._layout_request_fields()
        for field, box in self.mapping_boxes[role]["columns"].items():
            optional = field in OPTIONAL_FIELDS.get(role, ()) or (role == "職員" and field == WORK_DAYS_FIELD)
            box.configure(values=[UNUSED_COLUMN, *headers] if optional else headers)
            var = self.mapping_vars[role]["columns"][field]
            if preserve and (var.get() in headers or (optional and var.get() == UNUSED_COLUMN)):
                continue
            var.set(normalized.get(field, UNUSED_COLUMN if optional else ""))

    def _show_preview(self, role, frame):
        self._preview_role = role
        tree = self.preview_tree
        tree.delete(*tree.get_children())
        if frame is None:
            tree.configure(columns=())
            self.preview_label.configure(text=f"{role}：使用しません（任意の入力表です）")
            return
        # Unique internal column IDs also display duplicate source headings clearly.
        cols = [f"c{i}" for i in range(len(frame.columns))]
        tree.configure(columns=["source_row", *cols])
        tree.heading("source_row", text="元行")
        tree.column("source_row", width=70, minwidth=60, stretch=False, anchor="center")
        for col, heading in zip(cols, frame.columns):
            tree.heading(col, text=str(heading))
            tree.column(col, width=180, minwidth=120)
        for source_row, values in zip(frame.head(20).index, frame.head(20).itertuples(index=False, name=None)):
            tree.insert("", "end", values=[source_row, *["" if v is None else str(v) for v in values]])
        values = self.mapping_vars[role]
        self.preview_label.configure(text=f"先頭20行のプレビュー：{values['sheet'].get()} ／ 見出し {values['header_row'].get()}行目")

    def _on_role_preview_change(self, _event=None):
        if self.input_path and self._roles[self.role_notebook.index("current")] != self._preview_role:
            self.import_details.show()
        # タブ閲覧は列対応の変更ではないため、検証済みデータ・案を保持する。
        if self.busy or not self.input_path:
            return
        role = self._roles[self.role_notebook.index("current")]
        if role == self._preview_role:
            return
        try:
            header = self.mapping_vars[role]["header_row"].get()
            sheet = self.mapping_vars[role]["sheet"].get()
            if not 1 <= header <= 1000:
                raise ValueError("見出し行は1〜1000の整数で指定してください。")
        except (ValueError, tk.TclError) as exc:
            self._show_exception(exc, preserve_result=True, operation="mapping")
            return
        if sheet == UNUSED_COLUMN and role in OPTIONAL_SHEET_FIELDS:
            self._show_preview(role, None)
            return
        path = self.input_path
        def done(frame):
            self._show_preview(role, frame)
            self.status_var.set(f"{sheet} の先頭20行を表示しました。条件・シフト案は保持しています。")
        self._run_worker(lambda: preview_sheet(path, sheet, header), done,
                         "選択した役割の先頭20行を読み込んでいます…", preserve_result_on_error=True, operation="read")

    def refresh_role(self, role):
        if self.busy or not self.input_path:
            return
        self._invalidate_mapping()
        try:
            header = self.mapping_vars[role]["header_row"].get()
            if not 1 <= header <= 1000:
                raise ValueError("見出し行は1〜1000の整数で指定してください。")
            sheet = self.mapping_vars[role]["sheet"].get()
        except (ValueError, tk.TclError) as exc:
            self._show_exception(exc, operation="mapping")
            return
        if sheet == UNUSED_COLUMN and role in OPTIONAL_SHEET_FIELDS:
            self._show_preview(role, None)
            self._update_buttons()
            return
        def done(frame):
            self._apply_columns(role, frame.columns, preserve=True)
            self._show_preview(role, frame)
            self.status_var.set(f"{sheet} の見出しと先頭20行を反映しました。")
        self._run_worker(lambda: preview_sheet(self.input_path, sheet, header), done, "シートの見出しを確認しています…", operation="read")

    def _get_mappings(self):
        result = {}
        for role, values in self.mapping_vars.items():
            try:
                header = values["header_row"].get()
            except tk.TclError:
                raise ValueError(f"{role}：見出し行は整数で指定してください。") from None
            if not 1 <= header <= 1000:
                raise ValueError(f"{role}：見出し行は1〜1000の整数で指定してください。")
            if role in OPTIONAL_SHEET_FIELDS and values["sheet"].get() == UNUSED_COLUMN:
                result[role] = {"sheet": ""}
                continue
            active_fields = self._request_fields() if role == "希望休" else INPUT_FIELDS[role]
            used_fields = (*active_fields, *OPTIONAL_FIELDS.get(role, ()))
            columns = {field: ("" if values["columns"][field].get() == UNUSED_COLUMN else values["columns"][field].get())
                       for field in used_fields}
            if any(not columns.get(field) for field in active_fields if not (role == "職員" and field == WORK_DAYS_FIELD)):
                raise ValueError(f"{role}：必要項目に対応する元の列をすべて選択してください。")
            result[role] = {"sheet": values["sheet"].get(), "header_row": header, "columns": columns}
            if role == "希望休":
                result[role]["format"] = self._request_format()
        return result

    def read_conditions(self, *, show_settings=False, show_manager=False, show_requests=False):
        if self.busy or not self.input_path:
            return
        self.data = None
        self._clear_result()
        use_file_settings = self._settings_from_file_pending and "勤務条件" in self._sheet_names
        try:
            mappings = self._get_mappings()
        except ValueError as exc:
            self._show_exception(exc, preserve_result=True, operation="mapping")
            self._update_buttons()
            return
        try:
            chosen_settings = None if use_file_settings else self._get_settings()
        except ValueError as exc:
            self._show_exception(exc, preserve_result=True, operation="settings")
            self._update_buttons()
            return
        path = self.input_path
        captured_ids = frozenset(self.confirmed_paid_request_ids)
        captured_public_ids = frozenset(self.converted_public_request_ids)
        def task():
            raw = read_input(path, mappings, settings_override=None if use_file_settings else chosen_settings)
            settings = raw.settings if use_file_settings else chosen_settings
            eligible_ids = {request.request_id for request in raw.independent_requests if (request.original_leave_type or request.leave_type) == PAID_REQUEST}
            valid_ids = captured_ids & eligible_ids
            valid_public_ids = captured_public_ids & eligible_ids
            prepared = raw
            try:
                prepared = set_leave_choices(raw, valid_ids, valid_public_ids)
                data = configure_input(prepared, settings)
                report = build_request_heatmap(data)
                return raw, data, validate_input(data), valid_ids, valid_public_ids, data, report
            except InputError as exc:
                reference = replace(prepared, settings=settings)
                report = build_request_heatmap(reference, raw_fallback=True)
                return raw, None, exc.issues, valid_ids, valid_public_ids, reference, report
        def done(payload):
            raw, data, issues, valid_ids, valid_public_ids, heatmap_data, heatmap_report = payload
            self.base_data = raw
            self.confirmed_paid_request_ids = set(valid_ids)
            self.converted_public_request_ids = set(valid_public_ids)
            self._settings_from_file_pending = False
            self._heatmap_data = heatmap_data
            self._set_request_heatmap(heatmap_report, context="シフトは未作成。希望人数の入力集計です（採否は未計算）。")
            if use_file_settings:
                self._apply_settings(raw.settings)
            self._render_paid_manager()
            pending_paid = [request for request in self._paid_requests() if request.request_id not in self.confirmed_paid_request_ids and request.request_id not in self.converted_public_request_ids]
            if data is not None:
                self._render_requirement_preview(data)
            self._show_issues(issues, operation="read")
            if issues:
                self.data = None
                prefix = "条件不成立" if any(i.kind == "条件不成立" for i in issues) else "入力不備"
                self.conditions_var.set(f"{prefix}：{len(issues)}件。勤務条件と日数・人数を見直してください。")
                self.status_var.set(f"{prefix}：{len(issues)}件。条件の変更後に再確認してください。")
                if show_requests:
                    self.show_result_view("requests")
                elif pending_paid:
                    self.status_var.set(f"有給希望 {len(pending_paid)}件の配分が未決定です。人数を確認した後、公休・有給を決めてください。")
                    self.notebook.select(self.settings_tab if show_settings else self.manager_tab)
                else:
                    self.notebook.select(self.errors_tab)
                return
            self.data = data
            total = sum(e.work_days for e in data.employees if e.work_days is not None)
            work_days_text = (f"勤務日数合計：{total}人日" if all(e.work_days is not None for e in data.employees)
                              else "勤務日数：公休・有給・明けの条件から決定")
            reqs = len({r.request_id for r in data.requests})
            cap = f"{data.settings.max_consecutive_days}日" if data.settings.max_consecutive_days is not None else "制限なし"
            source = "曜日割合から算出" if data.settings.weekday_ratios is not None else "Excelの配置人数"
            history = "指定した連勤日数・前日区分を使用" if any(e.previous_work_days for e in data.employees) else "全員0日（未入力は前日休みと仮定）"
            night = (f"跨ぎ翌日は明け／非跨ぎ夜勤の翌日勤務禁止：{'有効' if data.settings.night_rest else '無効'}"
                     if data.settings.leave_enabled else "夜勤後は休み" if data.settings.night_rest else "夜勤後休み指定なし")
            transitions = "・".join(f"{a}→{b}" for a, b in data.settings.forbidden_transitions) or "指定なし"
            self.conditions_var.set(
                f"入力チェック完了：{len(data.employees)}人 ／ {len(data.days)}日 ／ {data.settings.mode}\n"
                f"対象期間：{data.days[0]:%Y-%m-%d} 〜 {data.days[-1]:%Y-%m-%d}\n"
                f"{work_days_text} ／ 配置人数合計：{sum(r.required for r in data.requirements)}人日\n"
                f"確定休：{len(data.fixed_off)}件 ／ 確定有給：{len(data.paid_leaves)}件 ／ 公休希望：{reqs}申請（候補{len(data.requests)}行）\n"
                f"人数：{source} ／ 連続勤務上限：{cap} ／ {night} ／ 禁止順序：{transitions}\n"
                f"前期間：{history}。最終日の跨ぎ勤務は次期間へ引き継ぎます。" +
                (f"\n休日区分：有効 ／ {data.settings.day_count_mode} ／ 公休は対象期間の値、有給・明けは別集計。" if data.settings.leave_enabled else ""))
            self.status_var.set("条件確認完了。算出人数と連勤条件を確認してシフト案を作成してください。")
            if show_requests:
                self.show_result_view("requests")
            else:
                self.notebook.select(self.manager_tab if show_manager else self.settings_tab)
        self._run_worker(task, done, "勤務モード・曜日割合・連勤条件と入力表を検証しています…", operation="read")

    def load_sample(self):
        if self.busy:
            return
        try:
            settings = self._get_settings()
        except ValueError as exc:
            self._show_exception(exc, preserve_result=True, operation="settings")
            return
        if settings.leave_enabled:
            self.set_settings(replace(settings, leave_enabled=False, day_count_mode=DIRECT_DAYS,
                                      default_public_holidays=None, shift_times_confirmed=False))
        def done(path):
            self.load_file(path, sample=True)
        self._run_worker(lambda: export_template(Path(self._temp.name)), done,
                         "記入例入りサンプルを作成し、選択した勤務条件を適用します…")

    def load_leave_sample(self):
        if self.busy:
            return
        def done(path):
            self.load_file(path, sample=True)
        self._run_worker(lambda: export_leave_sample(Path(self._temp.name)), done,
                         "公休・有給を分けた記入例入りサンプルを作成しています…")

    def load_independent_sample(self):
        if self.busy:
            return
        def done(path):
            self.load_file(path, sample=True)
        self._run_worker(lambda: export_independent_sample(Path(self._temp.name)), done,
                         "3日分を独立した希望として記入したサンプルを作成しています…")

    def save_template(self):
        if self.busy:
            return
        try:
            settings = self._get_settings()
        except ValueError as exc:
            self._show_exception(exc, preserve_result=True, operation="settings")
            return
        folder = filedialog.askdirectory(parent=self.root, title="記入例入り入力ひな形の保存先フォルダ")
        if not folder:
            return
        def done(path):
            self.status_var.set(f"入力ひな形を保存しました（3日分を独立した希望として入力・{settings.mode}）。")
            self.saved_var.set(f"保存先：{path}")
        self._run_worker(lambda: export_template(Path(folder), settings=settings, request_format="independent"), done,
                         "勤務条件を含む入力ひな形を保存しています…", preserve_result_on_error=True, operation="template_save")

    def save_mapping(self, path=None):
        if self.busy or not self.input_path:
            return
        try:
            mappings = self._get_mappings()
        except ValueError as exc:
            self._show_exception(exc, preserve_result=True, operation="mapping")
            return
        if path is None:
            path = filedialog.asksaveasfilename(parent=self.root, title="列対応設定を保存", defaultextension=".json",
                                               initialfile="sinta_shift_列対応.json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        def task():
            Path(path).write_text(json.dumps({"version": 1, "mappings": mappings}, ensure_ascii=False, indent=2), encoding="utf-8")
            return path
        self._run_worker(task, lambda p: self.status_var.set(f"列対応設定を保存しました：{p}"), "列対応設定を保存しています…",
                         preserve_result_on_error=True, operation="mapping_save")

    def load_mapping(self, path=None):
        if self.busy or not self.input_path:
            return
        if path is None:
            path = filedialog.askopenfilename(parent=self.root, title="列対応設定を選択", filetypes=[("JSON", "*.json")])
        if not path:
            return
        def task():
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
            if payload.get("version") != 1:
                raise ValueError("対応設定のバージョンが未対応です。")
            mappings = payload["mappings"]
            if not isinstance(mappings, dict) or not set(REQUIRED_FIELDS) <= set(mappings) or set(mappings) - set(INPUT_FIELDS):
                raise ValueError("列対応の表名が不足しているか、未対応の表名が含まれています。")
            previews = {}
            for role, fields in INPUT_FIELDS.items():
                if role not in mappings:
                    continue
                entry = mappings[role]
                if not isinstance(entry, dict):
                    raise ValueError(f"{role}：列対応の形式が不正です。")
                if role in OPTIONAL_SHEET_FIELDS and entry.get("sheet") == "":
                    previews[role] = None
                    continue
                header = entry["header_row"]
                if type(header) is not int or not 1 <= header <= 1000:
                    raise ValueError(f"{role}：見出し行が不正です。")
                columns = entry.get("columns")
                if role == "希望休":
                    format_name = entry.get("format", "alternatives")
                    if format_name not in REQUEST_FORMAT_NAMES:
                        raise ValueError("希望休：入力形式が不正です。")
                    fields = WIDE_REQUEST_FIELDS if format_name == "independent" or (format_name == "auto" and isinstance(columns, dict) and set(WIDE_REQUEST_FIELDS) <= set(columns)) else REQUIRED_FIELDS[role]
                allowed = set(fields) | set(OPTIONAL_FIELDS.get(role, ()))
                if not isinstance(columns, dict) or not set(fields) <= set(columns):
                    raise ValueError(f"{role}：必要項目が不足しています。")
                if set(columns) - allowed:
                    raise ValueError(f"{role}：未対応の必要項目が含まれています。")
                if entry["sheet"] not in self._sheet_names:
                    raise ValueError(f"{role}：シート {entry['sheet']} が入力ファイルにありません。")
                frame = preview_sheet(self.input_path, entry["sheet"], header)
                if not all(isinstance(v, str) and ((not v and (field not in fields or (role == "職員" and field == WORK_DAYS_FIELD))) or v in frame.columns)
                           for field, v in columns.items()):
                    raise ValueError(f"{role}：保存された列が入力ファイルにありません。")
                previews[role] = frame
            return mappings, previews
        def done(payload):
            self._invalidate_mapping()
            mappings, previews = payload
            for role, entry in mappings.items():
                values = self.mapping_vars[role]
                if role in OPTIONAL_SHEET_FIELDS and entry.get("sheet") == "":
                    values["sheet"].set(UNUSED_COLUMN)
                    self._apply_columns(role, ())
                    continue
                values["sheet"].set(entry["sheet"])
                values["header_row"].set(entry["header_row"])
                if role == "希望休":
                    self.request_format_var.set(REQUEST_FORMAT_NAMES[entry.get("format", "alternatives")])
                    self._detected_request_format = "independent" if set(WIDE_REQUEST_FIELDS) <= set(entry["columns"]) else "alternatives"
                self._apply_columns(role, previews[role].columns)
                for field, source in entry["columns"].items():
                    values["columns"][field].set(UNUSED_COLUMN if (field in OPTIONAL_FIELDS.get(role, ()) or (role == "職員" and field == WORK_DAYS_FIELD)) and source == "" else source)
            role = self._roles[self.role_notebook.index("current")]
            if role in previews:
                self._show_preview(role, previews[role])
            self.status_var.set("保存した列対応を復元しました。「条件を確認」を押してください。")
        self._run_worker(task, done, "保存した列対応を検証しています…", preserve_result_on_error=True, operation="mapping_load")

    def start_solve(self):
        if self.busy or self.data is None:
            return
        try:
            limit = float(self.time_var.get())
            if not 0 < limit <= 300:
                raise ValueError()
        except ValueError:
            self._show_exception(ValueError("制限時間は0より大きく300以下の秒数で指定してください。"), preserve_result=True, operation="solve")
            return
        self.last_time_limit = limit
        self._clear_result()
        data = self.data
        def task():
            result = solve_schedule(data, time_limit=limit)
            evaluation = evaluate_schedule(data, result.schedule) if result.schedule is not None else None
            report = build_request_heatmap(data, result.schedule, evaluation)
            return result, evaluation, report
        def done(payload):
            result, evaluation, report = payload
            self.result = result
            self.schedule = dict(result.schedule) if result.schedule is not None else None
            self.evaluation = evaluation
            self.edited = False
            context = "計算済みの案から採否を表示しています。" if evaluation is not None and evaluation.valid else f"{result.status.value}。採否は未計算で、入力の希望を表示しています。"
            self._set_request_heatmap(report, context=context, reset_selection=not self._heatmap_has_computed)
            self._heatmap_has_computed = evaluation is not None
            self._show_issues(issues_for_solve_result(result) if evaluation is None else evaluation.issues, operation="solve")
            if self.schedule is not None:
                self._render_result()
                self.show_result_view("requests")
                self.status_var.set(f"{result.status.value}：{len(data.employees)}人×{len(data.days)}日を計算しました。")
            else:
                self.result_summary.set(f"担当者確認用のシフト案：{result.status.value}\n{result.message}")
                self.status_var.set(f"{result.status.value}：{result.message}")
                self.show_result_view("requests")
        self._run_worker(task, done, f"シフト案を計算しています（全体の制限時間 {limit:g}秒）…", operation="solve")

    def _render_result(self):
        if self.data is None or self.schedule is None or self.evaluation is None:
            return
        data, ev = self.data, self.evaluation
        days = data.days
        totals_columns = ["勤務開始数", "公休", "有給", "明け", "実勤務日数"] if data.settings.leave_enabled else ["勤務日数"]
        columns = ["職員ID", "氏名"] + [d.isoformat() for d in days] + totals_columns
        tree = self.result_tree
        tree.delete(*tree.get_children())
        # 別期間の読込時に、前の週の日付列IDを新しい列へ持ち越さない。
        tree.configure(displaycolumns="#all")
        tree.configure(columns=columns)
        for col in columns:
            tree.heading(col, text=display_day(date.fromisoformat(col)) if col in [d.isoformat() for d in days] else col)
            tree.column(col, width=118 if col not in ("氏名",) else 150, minwidth=105, stretch=False, anchor="center")
        invalid_cells = set()
        for issue in ev.issues:
            if issue.employee_id and issue.day:
                invalid_cells.add((issue.employee_id, issue.day))
            elif issue.employee_id:
                invalid_cells.update((issue.employee_id, day) for day in days)
            elif issue.day:
                invalid_cells.update((e.id, issue.day) for e in data.employees)
        for index, employee in enumerate(data.employees):
            cells = []
            for day in days:
                value = assignment_label(self.schedule.get((employee.id, day)), data.settings) or "不正区分"
                cells.append(("⚠ " if (employee.id, day) in invalid_cells else "") + value)
            count = ev.employee_counts.get(employee.id, 0)
            target = employee.work_days
            if data.settings.leave_enabled:
                public_target = public_holiday_target(employee, data.settings)
                public_count = ev.employee_public_counts.get(employee.id, 0)
                totals = [f"{count} / {target}" if target is not None else count,
                    f"{public_count} / {public_target}" if public_target is not None else public_count,
                    ev.employee_paid_counts.get(employee.id, 0), ev.employee_recovery_counts.get(employee.id, 0),
                    ev.employee_actual_work_days.get(employee.id, 0)]
            else:
                totals = [f"{count} / {target}"]
            tree.insert("", "end", iid=f"emp{index}", values=[employee.id, employee.name, *cells, *totals])
        required = {r.day: r.required for r in data.requirements}
        tree.insert("", "end", iid="daily_actual", values=["配置人数", "実際 / 必要", *[
            ("⚠ " if ev.daily_counts.get(d, 0) != required[d] else "") + f"{ev.daily_counts.get(d, 0)} / {required[d]}" for d in days], *[""] * len(totals_columns)])
        tree.tag_configure("footer", background=GREEN, foreground=NAVY)
        tree.item("daily_actual", tags=("footer",))
        if data.settings.mode != DAY_ONLY:
            required_shifts = {r.day: dict(r.shift_counts) for r in data.requirements}
            for shift_index, shift in enumerate(SHIFT_MODES[data.settings.mode]):
                values = []
                for day in days:
                    actual = ev.daily_shift_counts.get(day, {}).get(shift, 0)
                    target = required_shifts.get(day, {}).get(shift, 0)
                    values.append(("⚠ " if actual != target else "") + f"{actual} / {target}")
                tree.insert("", "end", iid=f"shift_actual_{shift_index}",
                            values=[shift, "実際 / 必要", *values, *[""] * len(totals_columns)], tags=("footer",))
        self.request_tree.delete(*self.request_tree.get_children())
        independent = {request.request_id: request for request in data.independent_requests}
        wide_format = data.request_format == "independent"
        if wide_format:
            has_converted = any((request.original_leave_type or request.leave_type) != request.leave_type for request in data.independent_requests)
            request_columns = ("申請ID", "職員ID", "区分", "希望日・採用日", "日数", "結果", *(("元の希望",) if has_converted else ()))
            self.request_tree.configure(columns=request_columns)
            for column in request_columns:
                self.request_tree.heading(column, text=column)
                self.request_tree.column(column, width=180 if column in ("申請ID", "結果") else 135, minwidth=100)
            for result in (*ev.request_results, *ev.paid_request_results):
                request = independent.get(result.request_id)
                requested_day = request.day if request else result.adopted_day
                leave_type = request.leave_type if request else PUBLIC_REQUEST
                display_status = ("採用" if result.adopted_day is not None else "未採用") if leave_type == PUBLIC_REQUEST else result.status
                self.request_tree.insert("", "end", values=[result.request_id, result.employee_id, leave_type,
                    (result.adopted_day or requested_day).isoformat() if result.adopted_day or requested_day else "—",
                    1, display_status, *([(request.original_leave_type or request.leave_type) if request else leave_type] if has_converted else [])])
        else:
            request_columns = ("申請ID", "職員ID", "採用日", "優先順位", "結果")
            self.request_tree.configure(columns=request_columns)
            for column in request_columns:
                self.request_tree.heading(column, text=column)
                self.request_tree.column(column, width=150, minwidth=100)
            for result in ev.request_results:
                self.request_tree.insert("", "end", values=[result.request_id, result.employee_id,
                    result.adopted_day.isoformat() if result.adopted_day else "—",
                    f"第{result.priority}希望" if result.priority is not None else "—", result.status])
        unsatisfied = len(ev.request_results) - ev.adopted_count
        state = "手修正済み（最適性は未保証）" if self.edited else self.result.status.value
        check = "必須条件：すべて適合" if ev.valid else f"必須条件：違反{len(ev.issues)}件・完成版保存不可"
        detail = ("変更後の必須条件と希望休の採用状況を再計算しました。最適性は未保証です。"
                  if self.edited else self.result.message)
        if wide_format and not self.edited:
            detail = ("公休希望は各日独立です。有給は担当者が確定した日に固定しています。" +
                      ("公休の採用件数は最大と確認済みです。" if self.result.count_optimal
                       else "公休の採用件数の最適性は未確認です。"))
        request_summary = (f"公休希望採用 {ev.adopted_count}/{len(ev.request_results)}件・未採用{unsatisfied}件 ／ "
                           f"有給確定 {sum(result.status == '有給確定' for result in ev.paid_request_results)}件"
                           if wide_format else f"希望休採用 {ev.adopted_count}/{len(ev.request_results)}申請・未採用{unsatisfied}件・順位合計{ev.priority_sum}")
        self.result_summary.set(f"担当者確認用のシフト案 ｜ {data.settings.mode} ｜ {'公休・有給・明けを区別 ｜ ' if data.settings.leave_enabled else ''}{state}\n"
                                f"{check} ｜ {request_summary}\n"
                                f"{detail}")
        self.save_hint.set("必須条件に適合しています。担当者確認用の案として保存できます。" if ev.valid
                           else "保存できません。⚠セルと「エラー・違反一覧」の理由を確認し、修正してください。")
        self._show_issues(ev.issues)
        self._update_result_week()

    def show_result_week(self, index):
        if self.data is None or self.schedule is None:
            return
        weeks = list(dict.fromkeys(week_start(day) for day in self.data.days))
        self.result_week_index = max(0, min(int(index), len(weeks) - 1))
        self.show_all_days_var.set(False)
        self._update_result_week()

    def _update_result_week(self):
        if self.data is None or self.schedule is None:
            return
        days = self.data.days
        weeks = list(dict.fromkeys(week_start(day) for day in days))
        self.result_week_index = max(0, min(self.result_week_index, len(weeks) - 1))
        if self.show_all_days_var.get():
            self.result_tree.configure(displaycolumns="#all")
            self.result_week_var.set(f"全期間 {days[0]:%m/%d}〜{days[-1]:%m/%d}")
            return
        monday = weeks[self.result_week_index]
        visible_days = [day for day in days if week_start(day) == monday]
        display = ("職員ID", "氏名", *(day.isoformat() for day in visible_days))
        self.result_tree.configure(displaycolumns=display)
        if self._selected_cell and self._selected_cell[1] not in display:
            self._selected_cell = None
        available = max(550, self.root.winfo_width() - 270)
        date_width = max(66, min(122, int((available - 190) / max(1, len(visible_days)))))
        self.result_tree.column("職員ID", width=74, minwidth=64, stretch=False)
        self.result_tree.column("氏名", width=112, minwidth=90, stretch=False)
        for day in visible_days:
            self.result_tree.column(day.isoformat(), width=date_width, minwidth=64, stretch=False)
        self.result_week_var.set(f"第{self.result_week_index + 1}/{len(weeks)}週  {visible_days[0]:%m/%d}〜{visible_days[-1]:%m/%d}")
        self.result_tree.xview_moveto(0)

    def _remember_cell(self, event):
        row = self.result_tree.identify_row(event.y)
        identifier = self.result_tree.identify_column(event.x)
        display = self.result_tree.cget("displaycolumns")
        display = self.result_tree.cget("columns") if display in ("#all", ("#all",)) else display
        index = int(identifier[1:]) - 1 if identifier.startswith("#") else -1
        actual_column = display[index] if 0 <= index < len(display) else None
        self._selected_cell = (row, actual_column) if row and actual_column else None

    def _on_result_double_click(self, event):
        return self._show_shift_menu(event)

    def _show_shift_menu(self, event):
        self._remember_cell(event)
        cell = self._selected_employee_day()
        if self.busy or cell is None:
            return "break"
        employee_id, day = cell
        menu = tk.Menu(self.root, tearoff=False)
        current = assignment_label(self.schedule[(employee_id, day)], self.data.settings)
        for label in assignment_labels(self.data.settings):
            menu.add_command(label=f"{'✓ ' if label == current else '   '}{label}",
                             command=lambda value=label: self.toggle_cell(employee_id, day, value))
        menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def _on_space(self, _event):
        self._toggle_selected()
        return "break"

    def _selected_employee_day(self):
        if self.busy or not self._selected_cell or self.data is None or self.schedule is None:
            return None
        row, col = self._selected_cell
        if not row.startswith("emp"):
            return None
        index = int(row[3:])
        day_by_column = {day.isoformat(): day for day in self.data.days}
        display = self.result_tree.cget("displaycolumns")
        display = self.result_tree.cget("columns") if display in ("#all", ("#all",)) else display
        if not 0 <= index < len(self.data.employees) or col not in day_by_column or col not in display:
            return None
        return self.data.employees[index].id, day_by_column[col]

    def _toggle_selected(self):
        cell = self._selected_employee_day()
        if cell is not None:
            self.toggle_cell(*cell)

    def toggle_cell(self, employee_id, day, value=None):
        if self.busy or self.data is None or self.schedule is None:
            return
        key = (employee_id, day)
        if key not in self.schedule:
            return
        data = self.data
        labels = assignment_labels(data.settings)
        if value is None:
            current = assignment_label(self.schedule[key], data.settings)
            value = labels[(labels.index(current) + 1) % len(labels)] if current in labels else OFF
        if value not in labels:
            self._show_exception(ValueError("選択した勤務モードの区分を選んでください。"), preserve_result=True)
            return
        schedule = dict(self.schedule)
        # 日勤のみのbool形式は既存ファイル・検証スクリプトとの互換性を維持する。
        schedule[key] = value == "日勤" if data.settings.mode == DAY_ONLY and not data.settings.leave_enabled else value
        self._cancel_heatmap_refresh()
        def task():
            evaluation = evaluate_schedule(data, schedule)
            return evaluation, build_request_heatmap(data, schedule, evaluation)
        def done(payload):
            ev, report = payload
            self.schedule = schedule
            self.evaluation = ev
            self.edited = True
            self.saved_var.set("")
            self.last_saved = None
            self._render_result()
            self._set_request_heatmap(report, context="手修正後の案です。採否は参考表示（必須条件に違反あり）。" if not ev.valid else "手修正後の案から採否を再計算しました。")
            self.status_var.set("手修正を再検証しました。" + ("全必須条件を満たします。" if ev.valid else f"違反{len(ev.issues)}件：完成版保存は停止しています。"))
        self._run_worker(task, done, "変更後の人数・連勤・希望休を再計算しています…", operation="solve")

    def save_result(self, folder=None):
        if self.busy or self.data is None or self.schedule is None or self.result is None:
            return
        if self.evaluation is None or not self.evaluation.valid:
            self.status_var.set("必須条件違反があるため、完成版は保存できません。")
            self.notebook.select(self.errors_tab)
            return
        if folder is None:
            folder = filedialog.askdirectory(parent=self.root, title="結果の保存先（原本とは別フォルダ）")
        if not folder:
            return
        data, schedule, result = self.data, dict(self.schedule), self.result
        def done(path):
            self.last_saved = Path(path)
            self.saved_var.set(f"保存先：{path}")
            self.status_var.set(f"保存完了：{len(data.employees)}人×{len(data.days)}日、公休希望{len(self.evaluation.request_results)}件・有給申請{len(self.evaluation.paid_request_results)}件の結果を保存しました。")
        self._run_worker(lambda: export_result(data, schedule, result, Path(folder),
                         time_limit=self.last_time_limit, edited=self.edited, input_path=self.input_path),
                         done, "保存前に必須条件を再検証し、Excelを作成しています…", preserve_result_on_error=True, operation="save")

    def _show_issues(self, issues, *, operation=None, raw_detail=None):
        self.error_tree.delete(*self.error_tree.get_children())
        self._issue_by_iid = {}
        self._issue_operation = operation
        self._error_raw_detail = raw_detail
        for n, issue in enumerate(issues):
            iid = f"issue{n}"
            self._issue_by_iid[iid] = issue
            source = issue.source
            day_text = issue.day.isoformat() if hasattr(issue.day, "isoformat") else str(issue.day) if issue.day is not None else "—"
            guide = guidance_for_issue(issue, data=self.data or self._heatmap_data or self.base_data,
                has_schedule=self.schedule is not None, input_path=self.input_path, operation=operation)
            is_schedule = guide.target == "schedule"
            self.error_tree.insert("", "end", iid=iid, values=["表示中の案" if is_schedule else Path(source.file).name if source.file else "—",
                "勤務表（画面）" if is_schedule else source.sheet or "—", "—" if is_schedule else source.row or "—",
                issue.employee_id or "—", day_text, issue.message])
        count = len(self._issue_by_iid)
        self.notebook.tab(self.errors_tab, text=f"エラー・違反一覧（{count}件）")
        if count:
            retained = self.schedule is not None and self.evaluation is not None and self.evaluation.valid
            note = "案は保持しています。今回の保存は未完了です。" if retained and operation == "save" else "案は保持しています。直前の操作は未完了です。" if retained else "案作成・完成結果の保存を停止しています。"
            self.error_summary_var.set(f"確認が必要な項目 {count}件。{note} 行を選ぶと具体的な直し方を表示します。")
            self.error_alert_button.configure(text=f"エラーの直し方 {count}件")
            self.error_alert_button.pack(side="right", padx=8, after=self.workflow_button)
            first = next(iter(self._issue_by_iid))
            self.error_tree.selection_set(first)
            self.error_tree.focus(first)
            self.error_tree.see(first)
            self._focus_issue()
        else:
            self.error_alert_button.pack_forget()
            self.error_summary_var.set("確認が必要な項目はありません。")
            self.selected_error_guidance = None
            for variable in (self.error_location_var, self.error_reason_var, self.error_steps_var, self.error_next_var):
                variable.set("")
            self.error_title_var.set("確認が必要な項目はありません。")
            self._set_error_text(self.error_detail_text, "入力や勤務表を変更すると、必要な確認をここへ表示します。")
            self._set_error_text(self.error_raw_text, "")
        self._update_error_buttons()

    def show_error_guidance(self):
        self.notebook.select(self.errors_tab)
        if self._issue_by_iid and not self.error_tree.selection():
            self.error_tree.selection_set(next(iter(self._issue_by_iid)))
        self._focus_issue()

    def _focus_issue(self, _event=None):
        selected = self.error_tree.selection()
        issue = self._issue_by_iid.get(selected[0]) if selected else None
        if issue is None:
            self._update_error_buttons()
            return
        guide = guidance_for_issue(issue, data=self.data or self._heatmap_data or self.base_data,
            has_schedule=self.schedule is not None, input_path=self.input_path, operation=self._issue_operation)
        self.selected_error_guidance = guide
        self.error_title_var.set(guide.title)
        self.error_location_var.set(guide.location)
        self.error_reason_var.set(guide.reason)
        self.error_steps_var.set("\n".join(f"{index}. {step}" for index, step in enumerate(guide.steps, 1)))
        self.error_next_var.set(guide.next_action)
        self.error_detail_text.configure(state="normal")
        self.error_detail_text.delete("1.0", "end")
        reason = guide.reason if len(guide.reason) <= 180 else guide.reason[:180] + "…（全文は下の技術情報）"
        for heading, body in (("場所", guide.location), ("理由", reason)):
            self.error_detail_text.insert("end", heading + "：", "heading")
            self.error_detail_text.insert("end", body + "\n")
        self.error_detail_text.insert("end", "\n")
        for heading, body in (("具体的な直し方", self.error_steps_var.get()), ("修正後にすること", guide.next_action)):
            self.error_detail_text.insert("end", heading + "\n", "heading")
            self.error_detail_text.insert("end", body + "\n\n")
        self.error_detail_text.configure(state="disabled")
        self.error_detail_text.yview_moveto(0)
        raw = issue.message
        if self._error_raw_detail:
            raw += "\n\n" + self._error_raw_detail
        self._set_error_text(self.error_raw_text, raw)
        self.error_action_button.configure(text=guide.action_label)
        retry_by_category = {"mapping_load": "列対応JSONを選び直す", "mapping_save": "列対応JSONの保存先を選ぶ",
                             "settings_load": "勤務条件JSONを選び直す", "settings_save": "勤務条件JSONの保存先を選ぶ",
                             "template_save": "入力ひな形の保存先を選ぶ"}
        retry = retry_by_category.get(guide.category) or {
            "schedule": "現在の勤務表を再確認", "save": "保存先を選んで再保存",
            "run": "制限時間を変えて再計算", "settings": "設定を反映して再確認",
            "leave": "配分を反映して再確認", "input": "修正したExcelを再読込",
            "support": "元の理由を確認"}[guide.target]
        self.error_retry_button.configure(text=retry)
        self._update_error_buttons()

    def _update_error_buttons(self):
        if not hasattr(self, "error_action_button"):
            return
        guide = self.selected_error_guidance
        active = guide is not None and not self.busy
        self.error_action_button.configure(state="normal" if active else "disabled")
        json_settings = active and guide.category in ("settings_load", "settings_save", "template_save")
        can_retry = json_settings or active and ((guide.target == "save" and self.schedule is not None and self.evaluation is not None and self.evaluation.valid)
            or (guide.target == "schedule" and self.data is not None and self.schedule is not None)
            or (guide.target == "run" and self.data is not None)
            or (guide.target in ("input", "settings", "leave") and self.input_path is not None))
        self.error_retry_button.configure(state="normal" if can_retry else "disabled")

    def error_action(self):
        if self.busy or self.selected_error_guidance is None:
            return
        guide = self.selected_error_guidance
        if self._retry_error_file_operation(guide.category):
            return
        selected = self.error_tree.selection()
        issue = self._issue_by_iid.get(selected[0]) if selected else None
        if guide.target == "schedule" and issue is not None:
            self._jump_to_schedule_issue(issue)
        elif guide.target == "input":
            self.go_step(1)
            self.import_details.show()
            if issue is not None:
                role = next((role for role, values in self.mapping_vars.items() if values["sheet"].get() == issue.source.sheet), None)
                if role:
                    self.role_notebook.select(self.role_notebook.tabs()[self._roles.index(role)])
        elif guide.target == "settings":
            if guide.category in ("attendance_settings", "consecutive_settings"):
                self.go_step(2)
            else:
                self.notebook.select(self.advanced_settings_tab)
                index = 3 if guide.category == "time_settings" else 1 if guide.category == "sequence_settings" else 2 if guide.category == "leave_counts" else 0
                self.settings_sections.select(index)
        elif guide.target == "leave":
            self.go_step(3)
        elif guide.target == "save":
            self.save_result()
        elif guide.target == "run":
            self.notebook.select(self.conditions_tab)
        else:
            self.show_error_guidance()
            self.error_raw_text.focus_set()
        self.status_var.set(guide.next_action)

    def retry_after_error(self):
        if self.busy or self.selected_error_guidance is None:
            return
        if self._retry_error_file_operation(self.selected_error_guidance.category):
            return
        target = self.selected_error_guidance.target
        if target == "save":
            self.save_result()
        elif target == "schedule" and self.data is not None and self.schedule is not None:
            data, schedule = self.data, dict(self.schedule)
            def task():
                evaluation = evaluate_schedule(data, schedule)
                return evaluation, build_request_heatmap(data, schedule, evaluation)
            def done(payload):
                evaluation, report = payload
                self.evaluation = evaluation
                self._render_result()
                self._set_request_heatmap(report, context="現在の勤務表を再確認しました。" if evaluation.valid else "必須条件に違反があるため、採否は参考表示です。")
                self.show_error_guidance() if evaluation.issues else self.show_result_view("schedule")
                self.status_var.set("現在の勤務表を再確認しました。" + ("必須条件に適合しています。" if evaluation.valid else f"違反{len(evaluation.issues)}件は未解消です。"))
            self._run_worker(task, done, "現在の勤務表を再確認しています…", operation="solve")
        elif target == "run" and self.data is not None:
            self.start_solve()
        elif target == "input" and self.input_path is not None:
            # 外部で直した勤務条件・勤務時刻も含めて原本を読み直す。
            self.load_file(self.input_path, sample=True)
        elif target in ("settings", "leave") and self.input_path is not None:
            self.read_conditions()

    def _retry_error_file_operation(self, category):
        callbacks = {"mapping_load": self.load_mapping, "mapping_save": self.save_mapping,
                     "settings_load": self.load_settings, "settings_save": self.save_settings,
                     "template_save": self.save_template}
        callback = callbacks.get(category)
        if callback is None:
            return False
        callback()
        return True

    def _jump_to_schedule_issue(self, issue):
        if self.data is None or self.schedule is None:
            self.status_var.set("勤務表がまだありません。先に入力と条件を確認してください。")
            return
        self.show_result_view("schedule")
        employee_ids = [e.id for e in self.data.employees]
        if issue.employee_id in employee_ids:
            index = employee_ids.index(issue.employee_id)
            iid = f"emp{index}"
            self.result_tree.selection_set(iid)
            self.result_tree.see(iid)
            if issue.day in self.data.days:
                weeks = list(dict.fromkeys(week_start(day) for day in self.data.days))
                self.show_result_week(weeks.index(week_start(issue.day)))
                self._selected_cell = (iid, issue.day.isoformat())
        elif issue.day in self.data.days:
            weeks = list(dict.fromkeys(week_start(day) for day in self.data.days))
            self.show_result_week(weeks.index(week_start(issue.day)))
            self.result_tree.selection_set("daily_actual")
            self.result_tree.see("daily_actual")

    def _show_exception(self, exc, *, preserve_result=False, operation=None):
        self._sample_pending = False
        raw_detail = f"{type(exc).__name__}: {exc}"
        if isinstance(exc, InputError):
            if not preserve_result:
                self.data = None
                self.base_data = None
                self._heatmap_data = None
                self._clear_result()
                self._render_paid_manager()
            self._show_issues(issues_for_exception(exc, operation=operation), operation=operation, raw_detail=raw_detail)
            prefix = "条件不成立" if any(issue.kind == "条件不成立" for issue in exc.issues) else "操作を完了できません"
            if prefix == "条件不成立":
                self.conditions_var.set("条件不成立：必須条件が両立しません。選択行の直し方を確認してください。")
            self.status_var.set(f"{prefix}：{len(exc.issues)}件。選択行の『具体的な直し方』と『修正後の操作』を確認してください。")
        else:
            self._show_issues(issues_for_exception(exc, operation=operation), operation=operation, raw_detail=raw_detail)
            self.status_var.set("処理を完了できません。選択行の案内を確認してください。原因不明の場合は下のエラー全文も確認できます。")
        self.show_error_guidance()

    def close(self):
        if self.busy:
            self.status_var.set("処理中です。完了後に閉じてください。")
            return
        self._closed = True
        self._cancel_heatmap_refresh()
        self._temp.cleanup()
        self.root.destroy()


def main():
    prepare_tk_runtime()
    root = tk.Tk()
    ShiftApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
