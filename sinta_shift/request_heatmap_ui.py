"""希望職員人数だけを色へ対応付ける月カレンダー。Tk操作はmain thread専用。"""
from __future__ import annotations

import calendar
from datetime import date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk

from .visual_ui import ScrollPanel

NAVY = "#17324D"
INK = "#213549"
PALETTE = ("#FFFFFF", "#E8F4ED", "#C3E2CF", "#82BD9B", "#2F7654")
WEEKDAYS = "月火水木金土日"


class RequestHeatmap(ttk.Frame):
    """summariesは日付順tuple。色の基準は常に全対象期間の最多人数。"""
    def __init__(self, parent):
        super().__init__(parent)
        self.summaries = ()
        self._by_day = {}
        self._period_rows = ()
        self.month = None
        self.selected_day = None
        self.day_buttons = {}
        self.peak_count = 0
        self.context = "入力を読み込むと、希望が多い日を確認できます。"
        self._busy = False
        self.summary_var = tk.StringVar(value=self.context)
        ttk.Label(self, textvariable=self.summary_var, wraplength=930,
                  font=("Helvetica", 12, "bold")).pack(anchor="w", pady=(0, 2))
        self.legend_frame = ttk.Frame(self)
        self.legend_frame.pack(fill="x", pady=(0, 3))
        self.legend_var = tk.StringVar(value="色＝希望している職員人数。採否・確定休では色を変えません。")
        pane = ttk.Frame(self)
        pane.pack(fill="both", expand=True)
        pane.columnconfigure(0, weight=3, minsize=430, uniform="request_heatmap_panes")
        pane.columnconfigure(1, weight=2, minsize=310, uniform="request_heatmap_panes")
        pane.rowconfigure(0, weight=1)
        left = ttk.Frame(pane)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        month_bar = ttk.Frame(left)
        month_bar.pack(fill="x", pady=(0, 1))
        self.prev_button = ttk.Button(month_bar, text="‹ 前月", command=lambda: self.move_month(-1))
        self.prev_button.pack(side="left")
        self.month_var = tk.StringVar(value="対象期間：未読込")
        ttk.Label(month_bar, textvariable=self.month_var, font=("Helvetica", 14, "bold")).pack(side="left", padx=8)
        self.next_button = ttk.Button(month_bar, text="翌月 ›", command=lambda: self.move_month(1))
        self.next_button.pack(side="right")
        self.calendar_panel = ScrollPanel(left, background="white", body_style="Card.TFrame")
        self.calendar_panel.pack(fill="both", expand=True)
        self.grid_frame = ttk.Frame(self.calendar_panel.body, style="Card.TFrame")
        self.grid_frame.pack(fill="both", expand=True)
        right = ttk.Frame(pane)
        right.grid(row=0, column=1, sticky="nsew")
        self.detail_title_var = tk.StringVar(value="日付を選ぶと申請者を表示します")
        self.detail_title = ttk.Label(right, textvariable=self.detail_title_var,
                                     font=("Helvetica", 12, "bold"), wraplength=330)
        self.detail_title.pack(anchor="w", pady=(0, 5))
        self.detail_summary_var = tk.StringVar(value="")
        self.detail_summary = ttk.Label(right, textvariable=self.detail_summary_var, wraplength=330)
        self.detail_summary.pack(anchor="w", pady=(0, 5))
        self.warning_text = tk.Text(right, height=3, wrap="word", font=("Helvetica", 11),
                                    bg="#FFF4D8", fg=INK, relief="flat")
        table = ttk.Frame(right)
        table.pack(fill="both", expand=True)
        columns = ("氏名", "職員ID", "区分", "元の希望", "代替順位", "採否", "採用日", "確認状態", "元ファイル・セル", "申請ID")
        self.detail_tree = ttk.Treeview(table, columns=columns, show="headings", height=7)
        vertical = ttk.Scrollbar(table, orient="vertical", command=self.detail_tree.yview)
        horizontal = ttk.Scrollbar(table, orient="horizontal", command=self.detail_tree.xview)
        self.detail_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.detail_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        for column in columns:
            self.detail_tree.heading(column, text=column)
            self.detail_tree.column(column, width=145 if column not in ("元ファイル・セル", "採否") else 210, minwidth=90)
        self.detail_tree.column("氏名", width=130, minwidth=100)
        self.detail_tree.column("職員ID", width=90, minwidth=70)
        self.detail_tree.column("代替順位", width=100, minwidth=80)
        self.detail_tree.configure(displaycolumns=("氏名", "区分", "採否", "職員ID", "元の希望", "代替順位", "採用日", "確認状態", "元ファイル・セル", "申請ID"))
        right.bind("<Configure>", lambda event: self._fit_detail_labels(event.width))
        ttk.Label(self, text="色に採否・確定休は含めません。同じ職員の同日は1人として数えます。対象外は選べません。",
                  wraplength=930).pack(anchor="w", pady=(2, 0))
        style = ttk.Style(self)
        for level, color in enumerate(PALETTE):
            foreground = "white" if level == 4 else INK
            for selected in (False, True):
                name = self._style_name(level, selected)
                style.configure(name, background=color, foreground=foreground, padding=(1, 0),
                                font=("Helvetica", 11), anchor="center", relief="flat", borderwidth=2 if selected else 1,
                                bordercolor=NAVY if selected else "#D5E0D9")
                style.map(name, background=[("active", color), ("disabled", color)],
                          foreground=[("active", foreground), ("disabled", foreground)])
        self._render_calendar()

    @staticmethod
    def _style_name(level, selected=False):
        return f"Heatmap{'Selected' if selected else ''}{level}.TButton"

    def color_level(self, count):
        return 0 if count <= 0 or self.peak_count <= 0 else min(4, (count * 4 + self.peak_count - 1) // self.peak_count)

    def _fit_detail_labels(self, width):
        self.detail_title.configure(wraplength=max(180, width - 6))
        self.detail_summary.configure(wraplength=max(180, width - 6))
        available = max(260, width - 22)
        name_width = min(115, max(86, int(available * 0.34)))
        kind_width = min(105, max(76, int(available * 0.29)))
        self.detail_tree.column("氏名", width=name_width, minwidth=80, stretch=False)
        self.detail_tree.column("区分", width=kind_width, minwidth=70, stretch=False)
        self.detail_tree.column("採否", width=max(95, available - name_width - kind_width), minwidth=90, stretch=False)

    def set_summaries(self, summaries, *, context=""):
        self.summaries = tuple(summaries)
        self._period_rows = tuple(item for item in self.summaries if item.in_period)
        self._by_day = {item.day: item for item in self._period_rows}
        self.context = context or "希望人数を確認してください。"
        self.peak_count = max((item.request_employee_count for item in self._period_rows), default=0)
        if self._period_rows:
            first = self._period_rows[0].day.replace(day=1)
            last = self._period_rows[-1].day.replace(day=1)
            if self.month is None or not first <= self.month <= last:
                self.month = first
            if self.selected_day not in self._by_day:
                self.selected_day = max(self._period_rows, key=lambda item: item.request_employee_count).day
            peaks = [item.day for item in self._period_rows if item.request_employee_count == self.peak_count]
            peak_text = "、".join(f"{day:%m/%d}" for day in peaks[:3])
            self.summary_var.set(f"{self.context} ｜ 希望最多 {peak_text}：{self.peak_count}人" + (" ほか" if len(peaks) > 3 else ""))
        else:
            self.month = None
            self.selected_day = None
            self.summary_var.set(self.context)
        self._render_legend()
        self._render_calendar()
        if self.selected_day:
            self.select_day(self.selected_day)
        else:
            self.detail_tree.delete(*self.detail_tree.get_children())
            self.detail_title_var.set("日付を選ぶと申請者を表示します")
            self.detail_summary_var.set("")
            self.warning_text.pack_forget()

    def clear(self, message="入力を読み込むと、希望が多い日を確認できます。"):
        self.set_summaries((), context=message)
        self.calendar_panel.canvas.yview_moveto(0)

    def _render_legend(self):
        for child in self.legend_frame.winfo_children():
            child.destroy()
        buckets = {}
        for count in range(self.peak_count + 1):
            buckets.setdefault(self.color_level(count), []).append(count)
        labels = []
        for level, values in sorted(buckets.items()):
            label = f"{values[0]}人" if values[0] == values[-1] else f"{values[0]}〜{values[-1]}人"
            item = ttk.Frame(self.legend_frame)
            item.pack(side="left", padx=(0, 10))
            swatch = tk.Canvas(item, width=21, height=16, bg=PALETTE[level], highlightthickness=1, highlightbackground="#A5B9AC")
            swatch.pack(side="left", padx=(0, 4))
            ttk.Label(item, text=label).pack(side="left")
            labels.append(label)
        ttk.Label(self.legend_frame, text="人数（人）で着色・全期間共通").pack(side="left", padx=(2, 0))
        self.legend_var.set(f"色＝希望職員人数（人）。全対象期間の最多{self.peak_count}人を基準に固定。採否・確定休は色に含めません。")

    def _render_calendar(self):
        for child in self.grid_frame.winfo_children():
            child.destroy()
        self.day_buttons = {}
        for column, text in enumerate(WEEKDAYS):
            ttk.Label(self.grid_frame, text=text, anchor="center", style="Card.TLabel").grid(row=0, column=column, sticky="ew", pady=(1, 1))
            self.grid_frame.columnconfigure(column, weight=1, uniform="heatdays")
        if self.month is None or not self._period_rows:
            ttk.Label(self.grid_frame, text="希望休カレンダーは未読込です。", padding=20, style="Card.TLabel").grid(row=1, column=0, columnspan=7)
            self.prev_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
            self.month_var.set("対象期間：未読込")
            return
        self.month_var.set(f"{self.month.year}年 {self.month.month}月")
        first_month = self._period_rows[0].day.replace(day=1)
        last_month = self._period_rows[-1].day.replace(day=1)
        self.prev_button.configure(state="normal" if not self._busy and self.month > first_month else "disabled")
        self.next_button.configure(state="normal" if not self._busy and self.month < last_month else "disabled")
        for row, week in enumerate(calendar.Calendar(firstweekday=0).monthdatescalendar(self.month.year, self.month.month), 1):
            for column, day in enumerate(week):
                summary = self._by_day.get(day)
                if summary is None:
                    ttk.Label(self.grid_frame, text=f"{day.day}\n対象外", anchor="center", foreground="#88959E", padding=(0, 9), style="Card.TLabel").grid(
                        row=row, column=column, sticky="nsew", padx=2, pady=0)
                    continue
                date_label = str(day.day) if day.month == self.month.month else f"{day:%m/%d}"
                needed = "未設定" if summary.required_count is None else str(summary.required_count)
                text = f"{date_label}\n希望 {summary.request_employee_count}\n必要 {needed}"
                level = self.color_level(summary.request_employee_count)
                button = ttk.Button(self.grid_frame, text=text, style=self._style_name(level, day == self.selected_day),
                                    command=lambda value=day: self.select_day(value))
                button.grid(row=row, column=column, sticky="nsew", padx=2, pady=0)
                button.configure(state="disabled" if self._busy else "normal")
                button.bind("<Left>", lambda _event, value=day: self._keyboard_day(value, -1))
                button.bind("<Right>", lambda _event, value=day: self._keyboard_day(value, 1))
                button.bind("<Up>", lambda _event, value=day: self._keyboard_day(value, -7))
                button.bind("<Down>", lambda _event, value=day: self._keyboard_day(value, 7))
                self.day_buttons[day] = button

    def _keyboard_day(self, day, delta):
        target = day + timedelta(days=delta)
        if target in self._by_day:
            self.select_day(target)
            if target in self.day_buttons:
                self.day_buttons[target].focus_set()
                self._reveal_day(target)
        return "break"

    def _reveal_day(self, day):
        button = self.day_buttons.get(day)
        if button:
            self.update_idletasks()
            canvas = self.calendar_panel.canvas
            height = max(1, self.calendar_panel.body.winfo_height())
            top = button.winfo_y()
            current_top = canvas.canvasy(0)
            if top < current_top or top + button.winfo_height() > current_top + canvas.winfo_height():
                canvas.yview_moveto(max(0, min(1, top / height)))

    def move_month(self, delta):
        if self._busy or self.month is None or not self._period_rows:
            return
        month_number = self.month.year * 12 + self.month.month - 1 + delta
        candidate = date(month_number // 12, month_number % 12 + 1, 1)
        if self._period_rows[0].day.replace(day=1) <= candidate <= self._period_rows[-1].day.replace(day=1):
            self.month = candidate
            self._render_calendar()
            self.calendar_panel.canvas.yview_moveto(0)
            visible_days = [day for day in self._by_day if day.year == candidate.year and day.month == candidate.month]
            if visible_days:
                self.select_day(max(visible_days, key=lambda day: self._by_day[day].request_employee_count))

    def select_day(self, day):
        if day not in self._by_day:
            return
        previous = self.selected_day
        self.selected_day = day
        if self.month != day.replace(day=1):
            self.month = day.replace(day=1)
            self._render_calendar()
        for value in (previous, day):
            if value in self.day_buttons:
                summary = self._by_day[value]
                self.day_buttons[value].configure(style=self._style_name(self.color_level(summary.request_employee_count), value == day))
        summary = self._by_day[day]
        self.detail_title_var.set(f"{day:%Y/%m/%d}（{WEEKDAYS[day.weekday()]}）の希望")
        adopted = "未計算" if summary.adopted_employee_count is None else f"当日採用 {summary.adopted_employee_count}人"
        missed = "未計算" if summary.unadopted_request_count is None else f"未採用 {summary.unadopted_request_count}件"
        other = "" if summary.other_candidate_adopted_count is None else f" ／ 他候補採用 {summary.other_candidate_adopted_count}件"
        self.detail_summary_var.set(f"希望 {summary.request_employee_count}人（公休 {summary.public_request_employee_count}人・有給 {summary.paid_request_employee_count}人）\n"
            f"{adopted} ／ {missed}{other}\n確定休 {summary.fixed_off_employee_count}人・確定有給 {summary.fixed_paid_employee_count}人（色とは別）")
        self.detail_tree.delete(*self.detail_tree.get_children())
        for index, entry in enumerate(summary.entries):
            source = entry.source
            location = f"{Path(source.file).name} / {source.sheet}!{source.column or ('元行' + str(source.row) if source.row else '—')}"
            status = "当日採用" if entry.status == "採用" else entry.status
            self.detail_tree.insert("", "end", iid=f"entry{index}", values=[entry.employee_name, entry.employee_id, entry.kind,
                entry.original_kind or "—", f"第{entry.priority}希望" if entry.is_alternative and entry.priority is not None else "—（独立）", status,
                entry.adopted_day.isoformat() if entry.adopted_day else "—", entry.approval_status or "—", location, entry.request_id])
        if not summary.entries:
            self.detail_summary_var.set(self.detail_summary_var.get() + "\nこの日の希望申請は0件です。")
        self.warning_text.configure(state="normal")
        self.warning_text.delete("1.0", "end")
        if summary.warnings:
            self.warning_text.insert("1.0", "参考表示：" + " / ".join(summary.warnings))
            self.warning_text.pack(fill="x", pady=(0, 5), before=self.detail_tree.master)
        else:
            self.warning_text.pack_forget()
        self.warning_text.configure(state="disabled")

    def set_busy(self, busy):
        if self._busy == busy:
            return
        self._busy = busy
        for button in self.day_buttons.values():
            button.configure(state="disabled" if busy else "normal")
        if self.month and self._period_rows:
            self.prev_button.configure(state="disabled" if busy or self.month <= self._period_rows[0].day.replace(day=1) else "normal")
            self.next_button.configure(state="disabled" if busy or self.month >= self._period_rows[-1].day.replace(day=1) else "normal")
