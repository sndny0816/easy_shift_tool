"""勤務表GUIの表示部品。計算規則はplanningへ委譲する。"""
from __future__ import annotations

import calendar
from datetime import date
from decimal import Decimal
import tkinter as tk
from tkinter import ttk

from .planning import attendance_ratio_for_day, attendance_source_for_day, required_from_ratio, week_start

NAVY = "#17324D"
GREEN = "#DDEFE4"
BLUE = "#E5EFF8"
INK = "#213549"
BG = "#F3F6F8"


class Collapsible(ttk.Frame):
    def __init__(self, parent, text):
        super().__init__(parent)
        self.text = text
        self.expanded = False
        self.button = ttk.Button(self, text=f"＋ {text}", command=self.toggle)
        self.button.pack(anchor="w")
        self.body = ttk.Frame(self, padding=(0, 7, 0, 0))

    def show(self):
        if not self.expanded:
            self.expanded = True
            self.body.pack(fill="both", expand=True)
            self.button.configure(text=f"− {self.text}")

    def hide(self):
        if self.expanded:
            self.expanded = False
            self.body.pack_forget()
            self.button.configure(text=f"＋ {self.text}")

    def toggle(self):
        self.hide() if self.expanded else self.show()


class AttendanceCalendar(ttk.Frame):
    def __init__(self, parent, on_day, on_week):
        super().__init__(parent)
        self.on_day, self.on_week = on_day, on_week
        self.data = None
        self.settings = None
        self.month = None
        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 6))
        ttk.Button(top, text="‹", width=3, command=lambda: self.move_month(-1)).pack(side="left")
        self.title = tk.StringVar(value="Excelを読み込むと対象期間が表示されます")
        ttk.Label(top, textvariable=self.title, font=("Helvetica", 15, "bold")).pack(side="left", padx=8)
        ttk.Button(top, text="›", width=3, command=lambda: self.move_month(1)).pack(side="right")
        self.grid_frame = ttk.Frame(self)
        self.grid_frame.pack(fill="both", expand=True)
        self.day_buttons = {}
        self.week_buttons = {}
        style = ttk.Style(self)
        for name, color in (("CalendarDate.TButton", "white"), ("WeekCalendarDate.TButton", BLUE),
                            ("DayCalendarDate.TButton", GREEN)):
            style.configure(name, background=color, foreground=INK, padding=(2, 7),
                            font=("Helvetica", 11), anchor="center", relief="flat", borderwidth=1)
            style.map(name, background=[("active", "#CFDFD3"), ("disabled", "#E8EDF0")],
                      foreground=[("disabled", "#7B8994"), ("active", INK)])
        self._render()

    def set_data(self, data, settings):
        self.data, self.settings = data, settings
        if data and data.days:
            first = data.days[0].replace(day=1)
            last = data.days[-1].replace(day=1)
            if self.month is None or not first <= self.month <= last:
                self.month = first
        self._render()

    def move_month(self, delta):
        if not self.month:
            return
        total = self.month.year * 12 + self.month.month - 1 + delta
        candidate = date(total // 12, total % 12 + 1, 1)
        if self.data and self.data.days and self.data.days[0].replace(day=1) <= candidate <= self.data.days[-1].replace(day=1):
            self.month = candidate
            self._render()

    def _render(self):
        for child in self.grid_frame.winfo_children():
            child.destroy()
        self.day_buttons, self.week_buttons = {}, {}
        ttk.Label(self.grid_frame, text="週を設定", foreground=NAVY).grid(row=0, column=0, pady=5)
        for index, day in enumerate("月火水木金土日", 1):
            ttk.Label(self.grid_frame, text=day, anchor="center", foreground=NAVY).grid(row=0, column=index, sticky="ew", pady=5)
            self.grid_frame.columnconfigure(index, weight=1, uniform="days", minsize=56)
        if not self.data or not self.data.days or self.month is None:
            ttk.Label(self.grid_frame, text="先にExcelを読み込んでください。", padding=24).grid(row=1, column=0, columnspan=8)
            return
        self.title.set(f"{self.month.year}年 {self.month.month}月")
        valid_days = set(self.data.days)
        required = {item.day: item.required for item in self.data.requirements}
        for row_index, week in enumerate(calendar.Calendar(firstweekday=0).monthdatescalendar(self.month.year, self.month.month), 1):
            visible = [day for day in week if day in valid_days]
            week_button = ttk.Button(self.grid_frame, text=f"{week[0]:%m/%d}\n週全体", width=6,
                                     command=lambda day=week[0]: self.on_week(day))
            week_button.grid(row=row_index, column=0, sticky="nsew", padx=(0, 4), pady=3)
            if not visible:
                week_button.configure(state="disabled")
            self.week_buttons[week[0]] = week_button
            for column, day in enumerate(week, 1):
                if day not in valid_days:
                    ttk.Label(self.grid_frame, text=str(day.day), anchor="center", foreground="#A4ADB6", padding=(0, 20)).grid(
                        row=row_index, column=column, sticky="nsew", padx=2, pady=3)
                    continue
                ratio = attendance_ratio_for_day(day, self.settings)
                source = attendance_source_for_day(day, self.settings)
                count = required_from_ratio(len(self.data.employees), ratio) if ratio is not None else required[day]
                pct = f"{ratio.normalize():f}％" if isinstance(ratio, Decimal) else "Excel"
                button_style = "DayCalendarDate.TButton" if source == "日付指定" else "WeekCalendarDate.TButton" if source == "週指定" else "CalendarDate.TButton"
                text = f"{day.day}\n{pct} / {count}人\n{source}"
                button = ttk.Button(self.grid_frame, text=text, style=button_style,
                    command=lambda selected=day: self.on_day(selected))
                button.grid(row=row_index, column=column, sticky="nsew", padx=2, pady=3)
                self.day_buttons[day] = button
            self.grid_frame.rowconfigure(row_index, weight=1)


class MixBar(tk.Canvas):
    def __init__(self, parent):
        super().__init__(parent, bg="white", height=72, highlightthickness=0)
        self.counts = (0, 0)
        self.bind("<Configure>", lambda _event: self._draw())

    def set_counts(self, public_count, paid_count):
        self.counts = public_count, paid_count
        self._draw()

    def _draw(self):
        self.delete("all")
        public_count, paid_count = self.counts
        total = public_count + paid_count
        width = max(100, self.winfo_width() - 12)
        public_width = width * public_count / total if total else width
        self.create_rectangle(6, 8, 6 + public_width, 34, fill="#7DB797", outline="")
        if total and paid_count:
            self.create_rectangle(6 + public_width, 8, 6 + width, 34, fill=NAVY, outline="")
        public_pct = round(public_count * 100 / total) if total else 0
        paid_pct = 100 - public_pct if total else 0
        self.create_text(6, 52, anchor="w", text=f"公休希望 {public_count}件 ({public_pct}％)", fill=INK, font=("Helvetica", 12))
        self.create_text(6 + width, 52, anchor="e", text=f"有給 {paid_count}件 ({paid_pct}％)", fill=INK, font=("Helvetica", 12))


class ScrollPanel(ttk.Frame):
    def __init__(self, parent, *, background=BG, body_style="TFrame"):
        super().__init__(parent, style=body_style)
        self.canvas = tk.Canvas(self, bg=background, width=1, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.body = ttk.Frame(self.canvas, style=body_style)
        self.window = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda _event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda event: self.canvas.itemconfigure(self.window, width=event.width))
        top = self.winfo_toplevel()
        top.bind("<MouseWheel>", self._wheel, add="+")
        top.bind("<Button-4>", self._wheel, add="+")
        top.bind("<Button-5>", self._wheel, add="+")

    def _wheel(self, event):
        if not self.winfo_ismapped():
            return None
        widget = self.winfo_containing(getattr(event, "x_root", 0), getattr(event, "y_root", 0))
        widget = widget or event.widget
        current = widget
        inside = False
        while current is not None:
            if current is self:
                inside = True
                break
            if current.winfo_class() in {"Treeview", "TCombobox", "TSpinbox", "Spinbox", "Listbox", "Text"}:
                return None
            current = getattr(current, "master", None)
        if not inside:
            return None
        if getattr(event, "num", None) in (4, 5):
            units = -1 if event.num == 4 else 1
        else:
            delta = float(getattr(event, "delta", 0))
            if not delta:
                return None
            units = -round(delta / 120) if abs(delta) >= 120 else -1 if delta > 0 else 1
        self.canvas.yview_scroll(units, "units")
        return "break"
