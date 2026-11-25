"""値の.xlsx専用の入出力。原本を書き換えず、数式を実行・流用しない。"""
from __future__ import annotations

from dataclasses import replace
from datetime import date, datetime, time, timezone, timedelta
from decimal import Decimal, InvalidOperation
import math
import json
import os
from pathlib import Path
import re
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

from .models import (
    DayRequirement, Employee, FixedOff, InputData, InputError, Issue,
    LeaveCandidate, Schedule, SolveResult, SolverStatus, Source,
    ShiftSettings, DAY_ONLY, TWO_SHIFT, THREE_SHIFT, SHIFT_MODES, OFF,
    assignment_label, is_working,
    ShiftTime, PaidLeave, PUBLIC_OFF, PAID_LEAVE, RECOVERY, assignment_labels, is_actual_working,
    IndependentLeaveRequest, PUBLIC_REQUEST, PAID_REQUEST,
)

REQUIRED_FIELDS = {
    "職員": ("職員ID", "氏名", "対象期間の勤務日数"),
    "必要人数": ("日付", "配置人数"),
    "確定休": ("職員ID", "日付"),
    "希望休": ("申請ID", "職員ID", "候補日", "優先順位"),
}
OPTIONAL_SHEET_FIELDS = {"有給": ("職員ID", "日付")}
WIDE_REQUEST_FIELDS = ("職員ID", "希望日1", "区分1", "希望日2", "区分2", "希望日3", "区分3")
REQUEST_FORMATS = ("auto", "independent", "alternatives")
NAVY = "17324D"
GREEN = "E4F2EB"
WHITE = "FFFFFF"
DRAFT = "担当者確認用のシフト案"
SETTINGS_SHEET = "勤務条件"
PREVIOUS_DAYS = "前期間の連続勤務日数"
PREVIOUS_SHIFT = "前日の勤務区分"
PUBLIC_HOLIDAYS = "対象期間の公休日数"
SHIFT_TIMES_SHEET = "勤務時刻"
SHIFT_TIME_FIELDS = ("勤務区分", "開始時刻", "終了時刻", "翌日終了")
ATTENDANCE_SHEET = "出勤率調整"
ATTENDANCE_FIELDS = ("指定単位", "開始日", "出勤率％")
SPECIFY_WORK_DAYS = "勤務日数を指定"
DERIVE_PUBLIC_DAYS = "公休日数から算出"
SHIFT_COLUMNS = {label: f"{label}人数" for label in ("日勤", "夜勤", "準夜勤", "深夜勤")}
WEEKDAY_KEYS = tuple(f"{day}比率％" for day in "月火水木金土日")


def settings_to_payload(settings: ShiftSettings) -> dict:
    """JSONに安全な値だけを返す。比率はDecimalの精度を保つ文字列。"""
    return {
        "mode": settings.mode,
        "max_consecutive_days": settings.max_consecutive_days,
        "weekday_ratios": None if settings.weekday_ratios is None else [str(v) for v in settings.weekday_ratios],
        "shift_weights": list(settings.shift_weights),
        "night_rest": settings.night_rest,
        "forbidden_transitions": [list(pair) for pair in settings.forbidden_transitions],
        "leave_enabled": settings.leave_enabled,
        "day_count_mode": settings.day_count_mode,
        "default_public_holidays": settings.default_public_holidays,
        "shift_times": [{"label": item.label, "start_minute": item.start_minute, "end_minute": item.end_minute} for item in settings.shift_times],
        "shift_times_confirmed": settings.shift_times_confirmed,
        "weekly_ratios": [{"week_start": day.isoformat(), "percentage": str(percentage)} for day, percentage in settings.weekly_ratios],
        "date_ratios": [{"day": day.isoformat(), "percentage": str(percentage)} for day, percentage in settings.date_ratios],
    }


def settings_from_payload(payload: dict, *, source: Source = Source()) -> ShiftSettings:
    """勤務条件JSONの構文・型・範囲を検証する。業務の両立性は別途検証。"""
    if not isinstance(payload, dict):
        raise InputError([Issue("勤務条件", "勤務条件はJSONオブジェクトにしてください。", source)])
    allowed = {"mode", "max_consecutive_days", "weekday_ratios", "shift_weights", "night_rest", "forbidden_transitions",
               "leave_enabled", "day_count_mode", "default_public_holidays", "shift_times", "shift_times_confirmed", "weekly_ratios", "date_ratios"}
    issues = [Issue("勤務条件", f"未知の勤務条件項目です: {key}", source) for key in payload.keys() - allowed]
    mode = payload.get("mode", DAY_ONLY)
    mode = {"2交代制": TWO_SHIFT, "3交代制": THREE_SHIFT}.get(mode, mode) if isinstance(mode, str) else mode
    if not isinstance(mode, str) or mode not in SHIFT_MODES:
        issues.append(Issue("勤務条件", "勤務区分は日勤のみ・２交代制・３交代制から選んでください。", source))
        mode = DAY_ONLY
    maximum = payload.get("max_consecutive_days")
    if maximum is not None and (type(maximum) is not int or maximum < 1):
        issues.append(Issue("勤務条件", "連続勤務上限は1以上の整数、または未設定にしてください。", source))
    ratios = payload.get("weekday_ratios")
    parsed_ratios = None
    if ratios is not None:
        if not isinstance(ratios, (list, tuple)) or len(ratios) != 7:
            issues.append(Issue("勤務条件", "曜日別比率は月曜日から日曜日の7項目を指定してください。", source))
        else:
            parsed_ratios = []
            for index, value in enumerate(ratios):
                try:
                    if isinstance(value, (bool, float)) or _is_blank(value):
                        raise ValueError
                    number = Decimal(str(value).strip())
                    if not number.is_finite() or not Decimal(0) <= number <= Decimal(100):
                        raise ValueError
                    parsed_ratios.append(number)
                except (ValueError, InvalidOperation):
                    issues.append(Issue("勤務条件", f"{WEEKDAY_KEYS[index]}は0〜100の十進数を指定してください。", source))
            parsed_ratios = tuple(parsed_ratios)
    weights = payload.get("shift_weights", ())
    if not isinstance(weights, (list, tuple)) or any(type(v) is not int or v < 0 for v in weights):
        issues.append(Issue("勤務条件", "勤務区分の配分は0以上の整数を指定してください。", source))
        weights = ()
    elif weights and (len(weights) != len(SHIFT_MODES[mode]) or sum(weights) <= 0):
        issues.append(Issue("勤務条件", "配分は有効な勤務区分と同じ項目数にし、合計を1以上にしてください。", source))
    night_rest = payload.get("night_rest", True)
    if type(night_rest) is not bool:
        issues.append(Issue("勤務条件", "夜勤後休みはtrueまたはfalseにしてください。", source))
    leave_enabled = payload.get("leave_enabled", False)
    if type(leave_enabled) is not bool:
        issues.append(Issue("勤務条件", "公休・有給・明けの設定はtrueまたはfalseにしてください。", source))
    day_count_mode = payload.get("day_count_mode", SPECIFY_WORK_DAYS)
    if day_count_mode not in (SPECIFY_WORK_DAYS, DERIVE_PUBLIC_DAYS):
        issues.append(Issue("勤務条件", "日数指定方法は勤務日数を指定・公休日数から算出から選んでください。", source))
    default_public = payload.get("default_public_holidays")
    if default_public is not None and (type(default_public) is not int or default_public < 0):
        issues.append(Issue("勤務条件", "共通公休日数は0以上の整数、または未設定にしてください。", source))
    confirmed = payload.get("shift_times_confirmed", False)
    if type(confirmed) is not bool:
        issues.append(Issue("勤務条件", "勤務時刻確認はtrueまたはfalseにしてください。", source))
    times_payload = payload.get("shift_times", ())
    times = []
    if not isinstance(times_payload, (list, tuple)):
        issues.append(Issue("勤務条件", "勤務時刻は区分・開始分・終了分のオブジェクト配列にしてください。", source))
    else:
        for item in times_payload:
            if (not isinstance(item, dict) or set(item) != {"label", "start_minute", "end_minute"} or
                not isinstance(item["label"], str) or item["label"] not in SHIFT_COLUMNS or
                type(item["start_minute"]) is not int or type(item["end_minute"]) is not int or
                not 0 <= item["start_minute"] < 1440 or
                not item["start_minute"] < item["end_minute"] <= item["start_minute"] + 1440):
                issues.append(Issue("勤務条件", "勤務時刻は既知の区分、開始0〜1439分、開始より後で24時間以内の終了分を指定してください。", source))
                continue
            times.append(ShiftTime(item["label"], item["start_minute"], item["end_minute"]))
        if len(times) != len({item.label for item in times}):
            issues.append(Issue("勤務条件", "勤務時刻の区分が重複しています。", source))
    transitions = payload.get("forbidden_transitions", ())
    extra_labels = (PUBLIC_OFF, PAID_LEAVE, RECOVERY) if leave_enabled is True else (OFF,)
    if not isinstance(transitions, (list, tuple)) or any(
        not isinstance(pair, (list, tuple)) or len(pair) != 2 or
        any(not isinstance(label, str) or label not in (*SHIFT_MODES[mode], *extra_labels) for label in pair)
        for pair in transitions
    ):
        issues.append(Issue("勤務条件", "禁止勤務順序は有効な区分・休みから選んだ2項目の配列にしてください。", source))
        transitions = ()
    elif len(transitions) != len(set(tuple(pair) for pair in transitions)):
        issues.append(Issue("勤務条件", "禁止勤務順序が重複しています。", source))
    weekly_ratios, date_ratios = (), ()
    for key, day_key in (("weekly_ratios", "week_start"), ("date_ratios", "day")):
        try:
            parsed = _override_payload(payload.get(key, ()), day_key, source)
            if key == "weekly_ratios":
                weekly_ratios = parsed
            else:
                date_ratios = parsed
        except InputError as exc:
            issues.extend(exc.issues)
    if issues:
        raise InputError(issues)
    return ShiftSettings(mode, maximum, parsed_ratios, tuple(weights), night_rest, tuple(tuple(pair) for pair in transitions),
                         leave_enabled=leave_enabled, day_count_mode=day_count_mode, default_public_holidays=default_public,
                         shift_times=tuple(times), shift_times_confirmed=confirmed,
                         weekly_ratios=weekly_ratios, date_ratios=date_ratios)


def _percentage(value) -> Decimal:
    if isinstance(value, (bool, float, date, datetime)) or _is_blank(value):
        raise ValueError("0〜100の十進数を指定してください。")
    try:
        number = Decimal(str(value).strip())
        if not number.is_finite() or not Decimal(0) <= number <= Decimal(100):
            raise ValueError("0〜100の有限の十進数を指定してください。")
        return number
    except (InvalidOperation, TypeError) as exc:
        raise ValueError("0〜100の十進数として読み取れません。") from exc


def _is_percentage_format(value):
    code, quoted, index = value or "", False, 0
    while index < len(code):
        character = code[index]
        if character == chr(92):
            index += 2  # Excelの\%は表示用の文字で、倍率を持たない。
            continue
        if character == '"':
            quoted = not quoted
        elif character == "%" and not quoted:
            return True
        index += 1
    return False


def _override_payload(items, day_key, source):
    if not isinstance(items, (list, tuple)):
        raise InputError([Issue("出勤率", "出勤率調整は日付・百分率のオブジェクト配列にしてください。", source)])
    result, issues, seen = [], [], set()
    for item in items:
        try:
            if not isinstance(item, dict) or set(item) != {day_key, "percentage"}:
                raise ValueError(f"各指定には{day_key}とpercentageの2項目が必要です。")
            value = item[day_key]
            if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
                raise ValueError("日付はYYYY-MM-DD形式で指定してください。")
            day = _date(value)
            if day_key == "week_start" and day.weekday() != 0:
                raise ValueError("週の開始日は月曜日にしてください。")
            if day in seen:
                raise ValueError("同じ指定単位の日付が重複しています。")
            seen.add(day)
            result.append((day, _percentage(item["percentage"])))
        except (ValueError, OverflowError) as exc:
            issues.append(Issue("出勤率", f"{day_key}: {exc}", source))
    if issues:
        raise InputError(issues)
    return tuple(result)


def _read_attendance(workbook, path):
    if ATTENDANCE_SHEET not in workbook.sheetnames:
        return (), ()
    ws = workbook[ATTENDANCE_SHEET]
    first = next(ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=3))
    if any(cell.data_type in ("f", "e") for cell in first):
        raise InputError([Issue("数式", "出勤率調整の見出しに数式またはExcelエラーがあります。値を貼り付けてください。", Source(path.name, ATTENDANCE_SHEET, 1))])
    if _headers(ws, 1, path)[:3] != list(ATTENDANCE_FIELDS):
        raise InputError([Issue("出勤率", "出勤率調整の見出しは指定単位・開始日・出勤率％にしてください。", Source(path.name, ATTENDANCE_SHEET, 1))])
    weekly, daily, issues, seen = [], [], [], set()
    for row_no, row in enumerate(ws.iter_rows(min_row=2, min_col=1, max_col=3), 2):
        if all(_is_blank(cell.value) for cell in row):
            continue
        def source(index):
            return Source(path.name, ATTENDANCE_SHEET, row_no, f"{get_column_letter(index + 1)}{row_no}")
        bad = False
        for index, cell in enumerate(row):
            if cell.data_type in ("f", "e"):
                issues.append(Issue("数式", f"{ATTENDANCE_FIELDS[index]}に数式またはExcelエラーがあります。値を貼り付けてください。", source(index)))
                bad = True
        if bad:
            continue
        unit = str(row[0].value).strip() if row[0].value is not None else ""
        if unit not in ("週", "日付"):
            issues.append(Issue("出勤率", "指定単位は週または日付にしてください。", source(0)))
            continue
        try:
            day = _date(row[1].value)
            if unit == "週" and day.weekday() != 0:
                raise ValueError("週の開始日は月曜日にしてください。")
            if (unit, day) in seen:
                raise ValueError("同じ指定単位・開始日が重複しています。")
            seen.add((unit, day))
        except (ValueError, OverflowError) as exc:
            issues.append(Issue("出勤率", str(exc), source(1)))
            continue
        try:
            value = row[2].value
            # Excelが実際の百分率型(70%=0.7)で保存した場合もDecimalで扱う。
            number = Decimal(str(value).strip()) if isinstance(value, float) else value
            if _is_percentage_format(row[2].number_format) and isinstance(number, (int, Decimal)) and not isinstance(number, bool):
                number = Decimal(number) * 100
            percentage = _percentage(number)
        except (ValueError, InvalidOperation, OverflowError) as exc:
            issues.append(Issue("出勤率", f"出勤率％: {exc}", source(2)))
            continue
        (weekly if unit == "週" else daily).append((day, percentage))
    if issues:
        raise InputError(issues)
    return tuple(weekly), tuple(daily)


def _attendance_period_issues(workbook, path, requirements):
    if ATTENDANCE_SHEET not in workbook.sheetnames or not requirements:
        return []
    days = {item.day for item in requirements}
    ordered = sorted(days)
    if len(days) != len(requirements) or any((right - left).days != 1 for left, right in zip(ordered, ordered[1:])):
        return []  # 元の期間自体が不正な場合は、業務側の期間検証を優先する。
    issues = []
    for row_no, row in enumerate(workbook[ATTENDANCE_SHEET].iter_rows(min_row=2, min_col=1, max_col=3), 2):
        if all(_is_blank(cell.value) for cell in row):
            continue
        unit, day = str(row[0].value).strip(), _date(row[1].value)
        if unit == "日付" and day not in days:
            issues.append(Issue("出勤率", f"{day.isoformat()}の日付指定は対象期間外です。", Source(path.name, ATTENDANCE_SHEET, row_no, f"B{row_no}")))
        elif unit == "週" and not any(day <= target <= day + timedelta(days=6) for target in days):
            issues.append(Issue("出勤率", f"{day.isoformat()}からの週は対象期間と重なりません。", Source(path.name, ATTENDANCE_SHEET, row_no, f"B{row_no}")))
    return issues


def _read_settings(workbook, path: Path) -> ShiftSettings:
    weekly_ratios, date_ratios = _read_attendance(workbook, path)
    if SETTINGS_SHEET not in workbook.sheetnames:
        return replace(ShiftSettings(), shift_times=_read_times(workbook, path), weekly_ratios=weekly_ratios, date_ratios=date_ratios)
    ws = workbook[SETTINGS_SHEET]
    first_row = next(ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=2))
    if any(cell.data_type in ("f", "e") for cell in first_row):
        raise InputError([Issue("数式", "勤務条件の見出しに数式またはExcelエラーがあります。値を貼り付けてください。", Source(path.name, SETTINGS_SHEET, 1))])
    if _headers(ws, 1, path)[:2] != ["項目", "値"]:
        raise InputError([Issue("勤務条件", "勤務条件シートの1行目は「項目」「値」にしてください。", Source(path.name, SETTINGS_SHEET, 1))])
    known = {"勤務区分", "連続勤務上限", "必要人数の指定方法", "夜勤後休み", "禁止勤務順序",
             "公休・有給・明け", "日数指定方法", "共通公休日数", "勤務時刻確認済み",
             *WEEKDAY_KEYS, *(f"{label}配分" for label in SHIFT_COLUMNS)}
    values, sources, issues = {}, {}, []
    for row_no, row in enumerate(ws.iter_rows(min_row=2, min_col=1, max_col=2), 2):
        if all(_is_blank(cell.value) for cell in row):
            continue
        source = Source(path.name, SETTINGS_SHEET, row_no, f"B{row_no}")
        if any(cell.data_type in ("f", "e") for cell in row):
            issues.append(Issue("数式", "勤務条件の使用セルに数式またはExcelエラーがあります。値を貼り付けてください。", source))
            continue
        key = str(row[0].value).strip() if row[0].value is not None else ""
        if key not in known:
            issues.append(Issue("勤務条件", f"未知の勤務条件項目です: {key}", source))
        elif key in values:
            issues.append(Issue("勤務条件", f"勤務条件「{key}」が重複しています。", source))
        else:
            values[key], sources[key] = row[1].value, source
    if issues:
        raise InputError(issues)
    def error(key, message):
        issues.append(Issue("勤務条件", message, sources.get(key, Source(path.name, SETTINGS_SHEET, 1))))
    mode = values.get("勤務区分", DAY_ONLY)
    mode = {"2交代制": TWO_SHIFT, "3交代制": THREE_SHIFT}.get(mode, mode) if isinstance(mode, str) else mode
    if not isinstance(mode, str) or mode not in SHIFT_MODES:
        error("勤務区分", "勤務区分は日勤のみ・２交代制・３交代制から選んでください。")
        mode = DAY_ONLY
    maximum = None
    if not _is_blank(values.get("連続勤務上限")):
        try:
            maximum = _integer(values["連続勤務上限"])
            if maximum < 1:
                raise ValueError("1以上の整数を指定してください。")
        except (ValueError, OverflowError) as exc:
            error("連続勤務上限", f"連続勤務上限: {exc}")
    method = values.get("必要人数の指定方法", "Excel人数")
    if method not in ("Excel人数", "曜日別比率"):
        error("必要人数の指定方法", "必要人数の指定方法は「Excel人数」または「曜日別比率」にしてください。")
    ratios = None
    if method == "曜日別比率":
        ratios = []
        for key in WEEKDAY_KEYS:
            try:
                value = values.get(key)
                if _is_blank(value) or isinstance(value, bool):
                    raise ValueError
                number = Decimal(str(value).strip())
                if not number.is_finite() or not Decimal(0) <= number <= Decimal(100):
                    raise ValueError
                ratios.append(number)
            except (ValueError, InvalidOperation):
                error(key, f"{key}は0〜100の十進数を指定してください。空欄を0にはしません。")
        ratios = tuple(ratios)
    active = SHIFT_MODES[mode]
    weights = ()
    if any(not _is_blank(values.get(f"{label}配分")) for label in active):
        weights = []
        for label in active:
            key = f"{label}配分"
            try:
                value = _integer(values.get(key))
                if value < 0:
                    raise ValueError("0以上にしてください。")
                weights.append(value)
            except (ValueError, OverflowError) as exc:
                error(key, f"{key}: {exc}")
        if weights and sum(weights) <= 0:
            error(f"{active[0]}配分", "勤務区分の配分合計は1以上にしてください。")
        weights = tuple(weights)
    for label in set(SHIFT_COLUMNS) - set(active):
        key = f"{label}配分"
        if not _is_blank(values.get(key)):
            try:
                if _integer(values[key]) != 0:
                    raise ValueError("無効な区分の配分は空欄または0にしてください。")
            except (ValueError, OverflowError) as exc:
                error(key, f"{key}: {exc}")
    value = values.get("夜勤後休み", "有効")
    if value in (True, 1, "1", "有効"):
        night_rest = True
    elif value in (False, 0, "0", "無効"):
        night_rest = False
    else:
        error("夜勤後休み", "夜勤後休みは有効・無効、または1・0を指定してください。")
        night_rest = True
    leave_enabled = False
    confirmed = False
    for key in ("公休・有給・明け", "勤務時刻確認済み"):
        try:
            parsed_bool = _boolean(values.get(key, "無効"))
            if key == "公休・有給・明け":
                leave_enabled = parsed_bool
            else:
                confirmed = parsed_bool
        except ValueError as exc:
            error(key, f"{key}: {exc}")
    day_count_mode = values.get("日数指定方法", SPECIFY_WORK_DAYS)
    if day_count_mode not in (SPECIFY_WORK_DAYS, DERIVE_PUBLIC_DAYS):
        error("日数指定方法", "日数指定方法は勤務日数を指定・公休日数から算出から選んでください。")
    default_public = None
    if not _is_blank(values.get("共通公休日数")):
        try:
            default_public = _nonnegative_integer(values["共通公休日数"])
        except (ValueError, OverflowError) as exc:
            error("共通公休日数", f"共通公休日数: {exc}")
    transitions = ()
    if not _is_blank(values.get("禁止勤務順序")):
        try:
            transitions = json.loads(str(values["禁止勤務順序"]))
            parsed = settings_from_payload({"mode": mode, "leave_enabled": leave_enabled, "forbidden_transitions": transitions}, source=sources["禁止勤務順序"])
            transitions = parsed.forbidden_transitions
        except (ValueError, TypeError) as exc:
            if isinstance(exc, InputError):
                issues.extend(exc.issues)
            else:
                error("禁止勤務順序", '禁止勤務順序はJSON配列で指定してください（例: [["準夜勤","日勤"]]）。')
    if issues:
        raise InputError(issues)
    times = _read_times(workbook, path)
    return ShiftSettings(mode, maximum, ratios, weights, night_rest, transitions,
                         leave_enabled=leave_enabled, day_count_mode=day_count_mode,
                         default_public_holidays=default_public, shift_times=times, shift_times_confirmed=confirmed,
                         weekly_ratios=weekly_ratios, date_ratios=date_ratios)


def _open(path: str | Path):
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise InputError([Issue("形式", ".xlsx形式の値の表を選んでください。.xls・マクロ・CSV・PDFは対象外です。", Source(path.name))])
    try:
        return load_workbook(path, data_only=False, read_only=True, keep_links=False)
    except Exception as exc:
        raise InputError([Issue("読込", f"Excelを開けません。破損・パスワード・使用権限を確認してください: {exc}", Source(path.name))]) from exc


def inspect_workbook(path: str | Path) -> list[str]:
    workbook = _open(path)
    try:
        return workbook.sheetnames
    finally:
        workbook.close()


def _headers(sheet, header_row: int, path: Path) -> list[str]:
    if isinstance(header_row, bool) or not isinstance(header_row, int) or header_row < 1:
        raise InputError([Issue("見出し", "見出し行は1以上の整数を指定してください。", Source(path.name, sheet.title))])
    return [str(cell.value).strip() if cell.value is not None else "" for cell in next(sheet.iter_rows(min_row=header_row, max_row=header_row), ())]


def preview_sheet(path: str | Path, sheet: str, header_row: int = 1) -> pd.DataFrame:
    """見出しの次から最大20行。空欄と文字列NAを区別して表示する。"""
    path = Path(path)
    workbook = _open(path)
    try:
        if sheet not in workbook.sheetnames:
            raise InputError([Issue("シート", f"シート「{sheet}」がありません。", Source(path.name, sheet))])
        ws = workbook[sheet]
        headers = _headers(ws, header_row, path)
        # 元行番号はDataFrame indexに保持し、入力データに列を追加しない。
        rows = []
        indices = []
        for row in ws.iter_rows(min_row=header_row + 1, max_row=min(ws.max_row, header_row + 20), values_only=True):
            rows.append(["" if value is None else value for value in row])
            indices.append(header_row + len(rows))
        return pd.DataFrame(rows, columns=headers, index=indices, dtype=object)
    finally:
        workbook.close()


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _text(value: Any) -> str:
    if _is_blank(value) or isinstance(value, bool) or isinstance(value, (date, datetime)):
        raise ValueError("空欄、真偽値、日付は使えません。文字列のID・氏名を記入してください。")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _integer(value: Any) -> int:
    if _is_blank(value) or isinstance(value, (bool, date, datetime)):
        raise ValueError("空欄を0に置き換えません。整数を記入してください。")
    try:
        number = Decimal(str(value).strip())
        if not number.is_finite() or number != number.to_integral_value():
            raise ValueError("小数・無限大は使えません。整数を記入してください。")
        return int(number)
    except (InvalidOperation, TypeError) as exc:
        raise ValueError("整数として読み取れません。") from exc


def _nonnegative_integer(value: Any) -> int:
    number = _integer(value)
    if number < 0:
        raise ValueError("0以上の整数を記入してください。")
    return number


def _previous_days(value: Any) -> int:
    # この任意項目だけ、未記入は「対象期間の前日は休み」の明示された仮定。
    return 0 if _is_blank(value) else _nonnegative_integer(value)


def _previous_shift(value: Any) -> str | None:
    if _is_blank(value):
        return None
    label = _text(value)
    if label not in (*SHIFT_COLUMNS, OFF, PUBLIC_OFF, PAID_LEAVE, RECOVERY):
        raise ValueError("前日の勤務区分は日勤・夜勤・準夜勤・深夜勤・休みから選んでください。")
    return label


def _boolean(value: Any) -> bool:
    if type(value) is bool:
        return value
    if type(value) is int and value in (0, 1):
        return bool(value)
    if isinstance(value, str) and value.strip() in ("有効", "無効", "1", "0", "確認済み", "未確認"):
        return value.strip() in ("有効", "1", "確認済み")
    raise ValueError("有効・無効、または整数1・0を指定してください。")


def _clock_minutes(value: Any, *, allow_24=False) -> int:
    if isinstance(value, time):
        if value.second or value.microsecond:
            raise ValueError("時刻は分単位にしてください。")
        return value.hour * 60 + value.minute
    if isinstance(value, str) and re.fullmatch(r"\d{1,2}:\d{2}(?::00)?", value.strip()):
        parts = value.strip().split(":")
        hour, minute = int(parts[0]), int(parts[1])
        if 0 <= minute < 60 and (0 <= hour < 24 or allow_24 and hour == 24 and minute == 0):
            return hour * 60 + minute
    raise ValueError("時刻はHH:MM形式で記入してください（開始00:00〜23:59、終了24:00も可）。")


def _read_times(workbook, path: Path) -> tuple[ShiftTime, ...]:
    if SHIFT_TIMES_SHEET not in workbook.sheetnames:
        return ()
    ws = workbook[SHIFT_TIMES_SHEET]
    results, issues = [], []
    first = next(ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=4))
    if any(cell.data_type in ("f", "e") for cell in first):
        raise InputError([Issue("数式", "勤務時刻の見出しに数式またはExcelエラーがあります。値を貼り付けてください。", Source(path.name, SHIFT_TIMES_SHEET, 1))])
    if _headers(ws, 1, path)[:4] != list(SHIFT_TIME_FIELDS):
        raise InputError([Issue("勤務時刻", "勤務時刻シートの見出しは勤務区分・開始時刻・終了時刻・翌日終了にしてください。", Source(path.name, SHIFT_TIMES_SHEET, 1))])
    for row_no, row in enumerate(ws.iter_rows(min_row=2, min_col=1, max_col=4), 2):
        if all(_is_blank(cell.value) for cell in row):
            continue
        values, bad = [], False
        converters = (_text, _clock_minutes, lambda value: _clock_minutes(value, allow_24=True), _boolean)
        for index, (cell, convert) in enumerate(zip(row, converters)):
            source = Source(path.name, SHIFT_TIMES_SHEET, row_no, f"{get_column_letter(index + 1)}{row_no}")
            if cell.data_type in ("f", "e"):
                issues.append(Issue("数式", f"勤務時刻の「{SHIFT_TIME_FIELDS[index]}」に数式またはExcelエラーがあります。値を貼り付けてください。", source))
                bad = True
            else:
                try:
                    values.append(convert(cell.value))
                except (ValueError, OverflowError) as exc:
                    issues.append(Issue("勤務時刻", f"{SHIFT_TIME_FIELDS[index]}: {exc}", source))
                    bad = True
        if not bad:
            label, start, end, next_day = values
            if end == 1440 and next_day:
                issues.append(Issue("勤務時刻", "終了24:00と翌日終了1を同時に指定できません。翌日00:00なら終了00:00・翌日終了1、当日24:00なら終了24:00・翌日終了0にしてください。", Source(path.name, SHIFT_TIMES_SHEET, row_no, f"C{row_no}")))
                continue
            end += 1440 if next_day else 0
            try:
                parsed = settings_from_payload({"shift_times": [{"label": label, "start_minute": start, "end_minute": end}]}, source=Source(path.name, SHIFT_TIMES_SHEET, row_no))
                results.extend(parsed.shift_times)
            except InputError as exc:
                issues.extend(exc.issues)
    if len(results) != len({item.label for item in results}):
        issues.append(Issue("勤務時刻", "勤務時刻の区分が重複しています。", Source(path.name, SHIFT_TIMES_SHEET)))
    if issues:
        raise InputError(issues)
    return tuple(results)


def _date(value: Any) -> date:
    if isinstance(value, datetime):
        if value.time().hour or value.time().minute or value.time().second or value.time().microsecond:
            raise ValueError("時刻を含めず日付だけを記入してください。")
        return value.date()
    if isinstance(value, date):
        return value
    if not isinstance(value, str) or not re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", value.strip()):
        raise ValueError("年を含む日付（例: 2026-11-07）を記入してください。")
    try:
        year, month, day = map(int, re.split("[-/]", value.strip()))
        return date(year, month, day)
    except ValueError as exc:
        raise ValueError("存在する日付を記入してください。") from exc


def _request_header(value: str) -> str:
    return value.translate(str.maketrans("１２３", "123"))


def detect_request_format(headers, requested="auto") -> str:
    """GUIと読込で共通の判定。混在した見出しから意味を推測しない。"""
    if requested not in REQUEST_FORMATS:
        raise ValueError("希望休の形式はauto・independent・alternativesから選んでください。")
    if requested != "auto":
        return requested
    normalized = {_request_header(str(header).strip()) for header in headers}
    wide = set(WIDE_REQUEST_FIELDS) - {"職員ID"}
    legacy = set(REQUIRED_FIELDS["希望休"]) - {"職員ID"}
    if normalized & wide and normalized & legacy:
        raise ValueError("3日分の独立申請と従来の代替候補の見出しが混在しています。希望休の形式を明示して列を対応付けてください。")
    if set(WIDE_REQUEST_FIELDS) <= normalized:
        return "independent"
    if set(REQUIRED_FIELDS["希望休"]) <= normalized:
        return "alternatives"
    raise ValueError("希望休の形式を自動判定できません。3日分（各日独立）または従来（代替候補）を選び、列を対応付けてください。")


def _read_independent_requests(sheet, headers, columns, header_row, path, employees, period_days):
    positions, issues, requests, metadata = {}, [], [], []
    for field in WIDE_REQUEST_FIELDS:
        actual = columns.get(field)
        if field not in columns:
            matches = [name for name in headers if _request_header(name) == field]
        else:
            matches = [name for name in headers if name == actual]
        if len(matches) != 1 or headers.count(matches[0]) != 1:
            issues.append(Issue("列", f"必要項目「{field}」の対応列がないか重複しています。", Source(path.name, sheet.title, header_row, field)))
        else:
            positions[field] = headers.index(matches[0])
    if len(positions) != len(WIDE_REQUEST_FIELDS):
        return requests, metadata, issues
    if len(set(positions.values())) != len(positions):
        return requests, metadata, [Issue("列", "複数の必要項目に同じ元列を指定できません。", Source(path.name, sheet.title, header_row))]
    first = next(sheet.iter_rows(min_row=header_row, max_row=header_row))
    for position in positions.values():
        if first[position].data_type in ("f", "e"):
            issues.append(Issue("数式", "希望休の使用見出しに数式またはExcelエラーがあります。値を貼り付けてください。", Source(path.name, sheet.title, header_row, f"{get_column_letter(position + 1)}{header_row}")))
    if issues:
        return requests, metadata, issues
    staff_ids = {employee.id for employee in employees}
    days = set(period_days)
    seen_staff = set()
    for row_no, row in enumerate(sheet.iter_rows(min_row=header_row + 1), header_row + 1):
        if all(_is_blank(row[index].value) for index in positions.values()):
            continue
        def source(field):
            return Source(path.name, sheet.title, row_no, f"{get_column_letter(positions[field] + 1)}{row_no}")
        id_cell = row[positions["職員ID"]]
        if id_cell.data_type in ("f", "e"):
            issues.append(Issue("数式", "職員IDに数式またはExcelエラーがあります。値を貼り付けてください。", source("職員ID")))
            continue
        try:
            employee_id = _text(id_cell.value)
        except ValueError as exc:
            issues.append(Issue("変換", f"職員ID: {exc}", source("職員ID")))
            continue
        if employee_id in seen_staff:
            issues.append(Issue("独立申請", f"職員ID「{employee_id}」が複数行にあります。1職員1行にしてください。", source("職員ID"), employee_id))
            continue
        seen_staff.add(employee_id)
        if employee_id not in staff_ids:
            issues.append(Issue("独立申請", f"未知の職員IDです: {employee_id}", source("職員ID"), employee_id))
            continue
        seen_days = {}
        for index in range(1, 4):
            date_field, kind_field = f"希望日{index}", f"区分{index}"
            date_cell, kind_cell = row[positions[date_field]], row[positions[kind_field]]
            bad = False
            for field, cell in ((date_field, date_cell), (kind_field, kind_cell)):
                if cell.data_type in ("f", "e"):
                    issues.append(Issue("数式", f"{field}に数式またはExcelエラーがあります。値を貼り付けてください。", source(field), employee_id))
                    bad = True
            if bad:
                continue
            blank_date, blank_kind = _is_blank(date_cell.value), _is_blank(kind_cell.value)
            if blank_date and blank_kind:
                continue
            if blank_date or blank_kind:
                field = date_field if blank_date else kind_field
                issues.append(Issue("独立申請", f"{date_field}と{kind_field}は両方記入してください。両方空欄なら未申請です。", source(field), employee_id))
                continue
            try:
                day = _date(date_cell.value)
            except ValueError as exc:
                issues.append(Issue("変換", f"{date_field}: {exc}", source(date_field), employee_id))
                continue
            if day in seen_days:
                issues.append(Issue("独立申請", f"同じ職員の同じ日付が重複しています。先のセル: {seen_days[day]}", source(date_field), employee_id, day))
                continue
            seen_days[day] = source(date_field).column
            if day not in days:
                issues.append(Issue("独立申請", f"{day.isoformat()}は対象期間外です。", source(date_field), employee_id, day))
                continue
            try:
                kind = {PUBLIC_REQUEST: PUBLIC_REQUEST, PUBLIC_OFF: PUBLIC_REQUEST,
                        PAID_REQUEST: PAID_REQUEST, PAID_LEAVE: PAID_REQUEST}.get(_text(kind_cell.value))
            except ValueError:
                kind = None
            if kind is None:
                issues.append(Issue("独立申請", f"{kind_field}は「公休希望」または「有給申請」にしてください。", source(kind_field), employee_id, day))
                continue
            request_id = f"W:{employee_id}:{day.isoformat()}"
            src = source(date_field)
            metadata.append(IndependentLeaveRequest(request_id, employee_id, day, kind, src))
            if kind == PUBLIC_REQUEST:
                requests.append(LeaveCandidate(request_id, employee_id, day, 1, src))
    return requests, metadata, issues


def read_input(path: str | Path, mappings: dict | None = None, *, settings_override: ShiftSettings | None = None) -> InputData:
    """4つの論理表を読み込む。mappings[論理表]={sheet,header_row,columns}。"""
    path = Path(path)
    if mappings is not None and not isinstance(mappings, dict):
        raise InputError([Issue("対応設定", "列対応設定は表ごとの設定を持つJSONオブジェクトにしてください。", Source(path.name))])
    workbook = _open(path)
    all_issues: list[Issue] = []
    logical_fields = {**REQUIRED_FIELDS, **OPTIONAL_SHEET_FIELDS}
    parsed: dict[str, list] = {key: [] for key in logical_fields}
    converters = {
        "職員": (_text, _text, _integer),
        "必要人数": (_date, _integer),
        "確定休": (_text, _date),
        "希望休": (_text, _text, _date, _integer),
        "有給": (_text, _date),
    }
    constructors = {"職員": Employee, "必要人数": DayRequirement, "確定休": FixedOff, "希望休": LeaveCandidate, "有給": PaidLeave}
    settings = ShiftSettings()
    independent_requests = []
    request_format = "alternatives"
    try:
        try:
            settings = _read_settings(workbook, path)
        except InputError as exc:
            all_issues.extend(exc.issues)
        if settings_override is not None:
            if not isinstance(settings_override, ShiftSettings):
                all_issues.append(Issue("勤務条件", "勤務条件の上書きはShiftSettingsを指定してください。", Source(path.name)))
            else:
                try:
                    settings = settings_from_payload(settings_to_payload(settings_override), source=Source(path.name, SETTINGS_SHEET))
                except InputError as exc:
                    all_issues.extend(exc.issues)
        for role, base_fields in logical_fields.items():
            config = (mappings or {}).get(role, {})
            if not isinstance(config, dict):
                all_issues.append(Issue("対応設定", f"「{role}」の対応設定が不正です。", Source(path.name, role)))
                continue
            sheet_name = config.get("sheet", role)
            if role in OPTIONAL_SHEET_FIELDS and (sheet_name == "" or role not in (mappings or {}) and sheet_name not in workbook.sheetnames):
                continue
            header_row = config.get("header_row", 1)
            columns = config.get("columns", {})
            if not isinstance(columns, dict):
                all_issues.append(Issue("対応設定", f"「{role}」の列対応設定が不正です。", Source(path.name, str(sheet_name))))
                continue
            if sheet_name not in workbook.sheetnames:
                all_issues.append(Issue("シート", f"「{role}」に使うシート「{sheet_name}」がありません。シートを対応付けてください。", Source(path.name, str(sheet_name))))
                continue
            sheet = workbook[sheet_name]
            try:
                headers = _headers(sheet, header_row, path)
            except InputError as exc:
                all_issues.extend(exc.issues)
                continue
            if role == "希望休":
                try:
                    requested = config.get("format", "auto")
                    # 旧版で保存した別名列対応にはformatキーがない。旧4項目の
                    # 明示対応が揃い、横形式の見出しがなければ従来の意味を保つ。
                    wide_headers = set(WIDE_REQUEST_FIELDS) - {"職員ID"}
                    if ("format" not in config and set(REQUIRED_FIELDS[role]) <= set(columns)
                        and all(isinstance(columns[field], str) and columns[field] in headers for field in REQUIRED_FIELDS[role])
                        and not any(_request_header(header) in wide_headers for header in headers)):
                        requested = "alternatives"
                    request_format = detect_request_format(headers, requested)
                except ValueError as exc:
                    all_issues.append(Issue("申請形式", str(exc), Source(path.name, sheet_name, header_row)))
                    continue
                if request_format == "independent":
                    requests, metadata, issues = _read_independent_requests(
                        sheet, headers, columns, header_row, path, parsed["職員"], (item.day for item in parsed["必要人数"]))
                    parsed[role].extend(requests)
                    independent_requests.extend(metadata)
                    all_issues.extend(issues)
                    continue
            fields = list(base_fields)
            row_converters = list(converters[role])
            derived = settings.leave_enabled and settings.day_count_mode == DERIVE_PUBLIC_DAYS
            if role == "職員" and derived:
                work_field = "対象期間の勤務日数"
                actual = columns.get(work_field, work_field)
                if actual == "" or work_field not in columns and actual not in headers:
                    index = fields.index(work_field)
                    fields.pop(index)
                    row_converters.pop(index)
                else:
                    row_converters[fields.index(work_field)] = lambda value: None if _is_blank(value) else _nonnegative_integer(value)
            optional = {PREVIOUS_DAYS: _previous_days, PREVIOUS_SHIFT: _previous_shift,
                        PUBLIC_HOLIDAYS: lambda value: None if _is_blank(value) else _nonnegative_integer(value)} if role == "職員" else {}
            if role == "必要人数":
                for label, field in SHIFT_COLUMNS.items():
                    def convert_count(value, label=label):
                        if _is_blank(value) and label not in SHIFT_MODES[settings.mode]:
                            return None  # 無効区分の空欄は使用しない。0へ補完しない。
                        return _nonnegative_integer(value)
                    optional[field] = convert_count
            for field, convert in optional.items():
                # GUIの「（使用しない）」は、同名の元列が存在しても使わない。
                # キー自体を省略した場合だけ、従来どおり同名列を自動検出する。
                if field in columns and columns[field] == "":
                    continue
                if field in columns or field in headers:
                    fields.append(field)
                    row_converters.append(convert)
            positions = []
            for field in fields:
                actual = columns.get(field, field)
                if actual not in headers:
                    all_issues.append(Issue("列", f"必要項目「{field}」に対応する列「{actual}」がありません。", Source(path.name, sheet_name, header_row, field)))
                elif headers.count(actual) > 1:
                    all_issues.append(Issue("列", f"見出し「{actual}」が重複しています。一意な見出しに直してください。", Source(path.name, sheet_name, header_row, field)))
                else:
                    positions.append(headers.index(actual))
            if len(positions) != len(fields):
                continue
            header_cells = next(sheet.iter_rows(min_row=header_row, max_row=header_row))
            if any(header_cells[position].data_type == "f" for position in positions):
                all_issues.append(Issue("数式", "使用列の見出しに数式があります。値を貼り付けた別ファイルを使用してください。", Source(path.name, sheet_name, header_row)))
                continue
            if len(set(positions)) != len(positions):
                all_issues.append(Issue("列", "複数の必要項目に同じ元列を指定できません。", Source(path.name, sheet_name, header_row)))
                continue
            for row_no, row in enumerate(sheet.iter_rows(min_row=header_row + 1), header_row + 1):
                used = [row[index] for index in positions]
                if all(_is_blank(cell.value) for cell in used):
                    continue
                values = []
                bad = False
                for position, field, convert, cell in zip(positions, fields, row_converters, used):
                    source = Source(path.name, sheet_name, row_no, f"{get_column_letter(position + 1)}{row_no}")
                    if cell.data_type == "f":
                        all_issues.append(Issue("数式", f"「{field}」に数式があります。Excelで値を貼り付けた別ファイルを使用してください。", source))
                        bad = True
                        continue
                    if cell.data_type == "e":
                        all_issues.append(Issue("セルエラー", f"「{field}」にExcelエラーがあります: {cell.value}", source))
                        bad = True
                        continue
                    try:
                        values.append(convert(cell.value))
                    except (ValueError, OverflowError) as exc:
                        all_issues.append(Issue("変換", f"「{field}」: {exc}", source))
                        bad = True
                if not bad:
                    source = Source(path.name, sheet_name, row_no)
                    if role == "職員":
                        previous = values[fields.index(PREVIOUS_DAYS)] if PREVIOUS_DAYS in fields else 0
                        previous_shift = values[fields.index(PREVIOUS_SHIFT)] if PREVIOUS_SHIFT in fields else None
                        work_days = values[fields.index("対象期間の勤務日数")] if "対象期間の勤務日数" in fields else None
                        public_days = values[fields.index(PUBLIC_HOLIDAYS)] if PUBLIC_HOLIDAYS in fields else None
                        parsed[role].append(Employee(values[fields.index("職員ID")], values[fields.index("氏名")], work_days,
                                                     source=source, previous_work_days=previous, previous_shift=previous_shift, public_holidays=public_days))
                    elif role == "必要人数":
                        counts = tuple((label, values[fields.index(field)]) for label, field in SHIFT_COLUMNS.items()
                                       if field in fields and values[fields.index(field)] is not None)
                        parsed[role].append(DayRequirement(*values[:2], source=source, shift_counts=counts))
                    else:
                        parsed[role].append(constructors[role](*values, source=source))
        if not all_issues:
            all_issues.extend(_attendance_period_issues(workbook, path, parsed["必要人数"]))
    finally:
        workbook.close()
    if all_issues:
        raise InputError(all_issues)
    return InputData(tuple(parsed["職員"]), tuple(parsed["必要人数"]), tuple(parsed["確定休"]), tuple(parsed["希望休"]), settings=settings,
                     paid_leaves=tuple(parsed["有給"]), independent_requests=tuple(independent_requests), request_format=request_format)


def _put(sheet, row: int, column: int, value: Any):
    cell = sheet.cell(row=row, column=column)
    cell.value = value
    if isinstance(value, str):
        # '='などで始まる取り込み文字列も数式にしない。先頭文字は改変しない。
        cell.data_type = "s"
    if isinstance(value, (date, datetime)):
        cell.number_format = "yyyy-mm-dd"
    return cell


def _table(sheet, headers, rows, *, header_row=1, title=None, note=None):
    if title is not None:
        _put(sheet, 1, 1, title)
        sheet.cell(1, 1).font = Font(name="Yu Gothic", size=15, bold=True, color=NAVY)
        sheet.row_dimensions[1].height = 30
    if note:
        _put(sheet, 2, 1, note)
        sheet.cell(2, 1).font = Font(name="Yu Gothic", size=11, color=NAVY)
        sheet.row_dimensions[2].height = 24
    for j, value in enumerate(headers, 1):
        cell = _put(sheet, header_row, j, value)
        cell.font = Font(name="Yu Gothic", bold=True, color=WHITE, size=11)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for i, values in enumerate(rows, header_row + 1):
        for j, value in enumerate(values, 1):
            cell = _put(sheet, i, j, value)
            cell.font = Font(name="Yu Gothic", size=11, color=NAVY)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=GREEN)
        sheet.row_dimensions[i].height = 26
    sheet.row_dimensions[header_row].height = 32
    sheet.freeze_panes = f"A{header_row + 1}"
    sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{max(header_row, sheet.max_row)}"
    for j in range(1, len(headers) + 1):
        def width(value):
            return sum(2 if ord(char) > 255 else 1 for char in str(value))
        content = [sheet.cell(i, j).value for i in range(header_row, sheet.max_row + 1)]
        sheet.column_dimensions[get_column_letter(j)].width = min(64, max(14, max((width(v) for v in content if v is not None), default=10) + 3))
    if title is not None and header_row > 1:
        last_column = get_column_letter(len(headers))
        sheet.merge_cells(f"A1:{last_column}1")
        sheet.cell(1, 1).alignment = Alignment(vertical="center", wrap_text=True)
        if header_row > 2:
            sheet.merge_cells(f"A2:{last_column}2")
            sheet.cell(2, 1).alignment = Alignment(vertical="center", wrap_text=True)
            if note:
                available = max(1, sum(sheet.column_dimensions[get_column_letter(j)].width for j in range(1, len(headers) + 1)) - 3)
                lines = sum(max(1, math.ceil(width(line) / available)) for line in str(note).split("\n"))
                sheet.row_dimensions[2].height = max(28, lines * 17 + 8)
    # 自動行高はExcel以外の閲覧環境で不安定なため、折返し量に合わせて明示する。
    for row_no in range(header_row + 1, sheet.max_row + 1):
        lines = 1
        for col_no in range(1, len(headers) + 1):
            value = sheet.cell(row_no, col_no).value
            if value is not None:
                available = max(1, sheet.column_dimensions[get_column_letter(col_no)].width - 3)
                lines = max(lines, sum(max(1, math.ceil(width(line) / available)) for line in str(value).split("\n")))
        sheet.row_dimensions[row_no].height = max(26, lines * 17 + 7)
    sheet.sheet_view.showGridLines = False
    sheet.print_options.horizontalCentered = True
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = f"1:{header_row}"


def _save_unique(workbook: Workbook, folder: str | Path, basename: str) -> Path:
    """O_EXCLで同時実行時にも原本・既存結果を上書きしない。"""
    folder = Path(folder).expanduser().resolve()
    if not folder.is_dir():
        raise ValueError("保存先として存在するフォルダを選んでください。")
    stem, suffix = Path(basename).stem, Path(basename).suffix
    index = 0
    while True:
        target = folder / (f"{stem}{suffix}" if index == 0 else f"{stem}_{index}{suffix}")
        try:
            descriptor = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            break
        except FileExistsError:
            index += 1
    try:
        with os.fdopen(descriptor, "wb") as stream:
            workbook.save(stream)
    except BaseException:
        target.unlink(missing_ok=True)
        raise
    finally:
        workbook.close()
    return target


def export_template(folder: str | Path, settings: ShiftSettings | None = None, *, request_format: str = "alternatives") -> Path:
    """記入例を保存。既定は旧4表、条件変更時は勤務条件と任意列を追加する。"""
    from .sample import sample_data
    data = sample_data()
    advanced = settings is not None and settings != ShiftSettings()
    if settings is not None and settings.leave_enabled:
        settings = settings_from_payload(settings_to_payload(settings))
        data = _leave_sample_data(settings)
        settings = data.settings
        advanced = True
    elif advanced:
        from .planning import configure_input
        settings = settings_from_payload(settings_to_payload(settings))
        data = configure_input(data, settings)
    if request_format not in ("alternatives", "independent"):
        raise ValueError("ひな形の申請形式はalternativesまたはindependentを指定してください。")
    return _write_input_template(folder, data, advanced, request_format)


def _write_input_template(folder, data, advanced, request_format):
    settings = data.settings
    active = SHIFT_MODES[data.settings.mode]
    workbook = Workbook()
    workbook.remove(workbook.active)
    rows = {
        "職員": [(e.id, e.name, e.work_days, *([e.previous_work_days, e.previous_shift or (PUBLIC_OFF if settings.leave_enabled else OFF)] if advanced else []),
                  *([e.public_holidays] if advanced and settings.leave_enabled else [])) for e in data.employees],
        "必要人数": [(d.day, d.required, *([dict(d.shift_counts)[label] for label in active] if advanced else [])) for d in data.requirements],
        "確定休": [(f.employee_id, f.day) for f in data.fixed_off],
        "希望休": [(r.request_id, r.employee_id, r.day, r.priority) for r in data.requests],
    }
    if request_format == "independent":
        grouped = {employee.id: {} for employee in data.employees}
        if data.independent_requests:
            for item in data.independent_requests:
                grouped[item.employee_id][item.day] = item.leave_type
        else:
            for item in data.requests:
                grouped[item.employee_id][item.day] = PUBLIC_REQUEST
            for item in data.paid_leaves:
                if item.day in grouped[item.employee_id]:
                    raise ValueError("記入例に同日の公休希望と有給が重複しています。")
                grouped[item.employee_id][item.day] = PAID_REQUEST
        if any(len(items) > 3 for items in grouped.values()):
            raise ValueError("横形式の記入例は1職員3日までです。")
        rows["希望休"] = [[employee.id, *[value for pair in list(grouped[employee.id].items()) + [(None, None)] * (3 - len(grouped[employee.id])) for value in pair]] for employee in data.employees]
    for role, fields in REQUIRED_FIELDS.items():
        ws = workbook.create_sheet(role)
        if role == "希望休" and request_format == "independent":
            fields = WIDE_REQUEST_FIELDS
        if advanced and role == "職員":
            fields = (*fields, PREVIOUS_DAYS, PREVIOUS_SHIFT)
            if settings.leave_enabled:
                fields = (*fields, PUBLIC_HOLIDAYS)
        elif advanced and role == "必要人数":
            fields = (*fields, *(SHIFT_COLUMNS[label] for label in active))
        _table(ws, fields, rows[role])
        id_columns = {"職員": ("A",), "必要人数": (), "確定休": ("A",), "希望休": ("A",) if request_format == "independent" else ("A", "B")}[role]
        for column in id_columns:
            ws.column_dimensions[column].number_format = "@"
            for cells in ws.iter_rows(min_row=2, min_col=ws[column + "1"].column, max_col=ws[column + "1"].column):
                cells[0].number_format = "@"
        date_column = {"職員": None, "必要人数": "A", "確定休": "B", "希望休": None if request_format == "independent" else "C"}[role]
        if date_column:
            ws.column_dimensions[date_column].number_format = "yyyy-mm-dd"
        if role == "希望休" and request_format == "independent":
            _style_wide_template(ws)
    if advanced:
        conditions = [
            ("勤務区分", settings.mode),
            ("連続勤務上限", settings.max_consecutive_days),
            ("必要人数の指定方法", "Excel人数" if settings.weekday_ratios is None else "曜日別比率"),
            *[(key, None if settings.weekday_ratios is None else str(settings.weekday_ratios[index])) for index, key in enumerate(WEEKDAY_KEYS)],
            *[(f"{label}配分", settings.shift_weights[index] if settings.shift_weights else None) for index, label in enumerate(active)],
            ("夜勤後休み", "有効" if settings.night_rest else "無効"),
            ("禁止勤務順序", json.dumps(settings.forbidden_transitions, ensure_ascii=False)),
            ("公休・有給・明け", "有効" if settings.leave_enabled else "無効"),
            ("日数指定方法", settings.day_count_mode),
            ("共通公休日数", settings.default_public_holidays),
            ("勤務時刻確認済み", "確認済み" if settings.shift_times_confirmed else "未確認"),
        ]
        _table(workbook.create_sheet(SETTINGS_SHEET), ["項目", "値"], conditions)
        if settings.leave_enabled:
            fixed_paid = [] if request_format == "independent" else [(item.employee_id, item.day) for item in data.paid_leaves]
            _table(workbook.create_sheet("有給"), OPTIONAL_SHEET_FIELDS["有給"], fixed_paid)
        if settings.shift_times:
            rows = [(item.label, _format_clock(item.start_minute), _format_clock(item.end_minute if item.end_minute <= 1440 else item.end_minute - 1440), int(item.end_minute > 1440)) for item in settings.shift_times]
            _table(workbook.create_sheet(SHIFT_TIMES_SHEET), SHIFT_TIME_FIELDS, rows)
    if settings.weekly_ratios or settings.date_ratios:
        adjustments = [("週", day, str(percentage)) for day, percentage in settings.weekly_ratios]
        adjustments.extend(("日付", day, str(percentage)) for day, percentage in settings.date_ratios)
        _table(workbook.create_sheet(ATTENDANCE_SHEET), ATTENDANCE_FIELDS, adjustments)
    return _save_unique(workbook, folder, "sample_input.xlsx")


def _style_wide_template(sheet):
    for column in ("B", "D", "F"):
        sheet.column_dimensions[column].number_format = "yyyy-mm-dd"
        sheet.column_dimensions[column].width = 17
    for column in ("C", "E", "G"):
        sheet.column_dimensions[column].width = 17
        dropdown = DataValidation(type="list", formula1='"公休希望,有給申請"', allow_blank=True)
        dropdown.errorTitle = "区分を選んでください"
        dropdown.error = "公休希望または有給申請を選んでください。"
        dropdown.showErrorMessage = True
        dropdown.errorStyle = "stop"
        sheet.add_data_validation(dropdown)
        dropdown.add(f"{column}2:{column}1000")
    note = _put(sheet, 1, 9, "希望日1・2・3は、3日分の独立した申請です。\n『どれか1日』の代替候補ではありません。\n各日付と区分をセットで記入し、両方空欄なら未申請です。有給申請は読込後に担当者が確認してから固定します。")
    sheet.merge_cells("I1:L4")
    note.font = Font(name="Yu Gothic", size=11, color=NAVY)
    note.fill = PatternFill("solid", fgColor=GREEN)
    note.alignment = Alignment(wrap_text=True, vertical="top")
    for column in ("I", "J", "K", "L"):
        sheet.column_dimensions[column].width = 13


def export_independent_sample(folder: str | Path) -> Path:
    """30日3職員・公休9日/有給申請1日/勤務20日の確認前サンプル。"""
    settings = ShiftSettings(leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS, default_public_holidays=9)
    days = tuple(date(2026, 11, 1) + timedelta(days=index) for index in range(30))
    employees, requests, metadata = [], [], []
    for index in range(3):
        employee_id = f"{index + 1:03d}"
        off_days = [day for offset, day in enumerate(days) if offset % 3 == index]
        employees.append(Employee(employee_id, f"サンプル職員{index + 1}", 20, previous_shift=PUBLIC_OFF, public_holidays=9))
        for slot, day in enumerate(off_days[:3]):
            kind = PAID_REQUEST if slot == 2 else PUBLIC_REQUEST
            request_id = f"W:{employee_id}:{day.isoformat()}"
            metadata.append(IndependentLeaveRequest(request_id, employee_id, day, kind))
            if kind == PUBLIC_REQUEST:
                requests.append(LeaveCandidate(request_id, employee_id, day, 1))
    data = InputData(tuple(employees), tuple(DayRequirement(day, 2, shift_counts=(("日勤", 2),)) for day in days),
                     requests=tuple(requests), settings=settings, independent_requests=tuple(metadata), request_format="independent")
    return _write_input_template(folder, data, True, "independent")


def _format_clock(minute: int) -> str:
    return f"{minute // 60:02d}:{minute % 60:02d}"


def _leave_sample_data(settings: ShiftSettings) -> InputData:
    """架空の入力例。結果は固定せず、この条件をCP-SATで解く。"""
    if settings.default_public_holidays is None:
        settings = replace(settings, default_public_holidays=8)
    if settings.mode == TWO_SHIFT and not settings.shift_times:
        settings = replace(settings, shift_times=(ShiftTime("日勤", 480, 1020), ShiftTime("夜勤", 960, 1980)), shift_times_confirmed=False)
    elif settings.mode == THREE_SHIFT and not settings.shift_times:
        settings = replace(settings, shift_times=(ShiftTime("日勤", 480, 960), ShiftTime("準夜勤", 960, 1440), ShiftTime("深夜勤", 0, 480)), shift_times_confirmed=False)
    active = SHIFT_MODES[settings.mode]
    crossing = [item.label for item in settings.shift_times if item.label in active and item.end_minute > 1440]
    if settings.mode == DAY_ONLY:
        pattern = (PUBLIC_OFF, "日勤", "日勤")
    elif crossing:
        other = [label for label in active if label not in crossing] or [active[0]]
        pattern = (crossing[0], RECOVERY, PUBLIC_OFF, other[0], other[-1], PUBLIC_OFF)
    elif settings.mode == TWO_SHIFT:
        pattern = ("夜勤", PUBLIC_OFF, "日勤", "日勤", "日勤", PUBLIC_OFF)
    else:
        pattern = ("深夜勤", PUBLIC_OFF, "日勤", "日勤", "準夜勤", PUBLIC_OFF)
    days = tuple(date(2026, 11, 1) + timedelta(days=index) for index in range(30))
    employees, paid, requests = [], [], []
    for index in range(len(pattern)):
        employee_id = f"{index + 1:03d}"
        values = [pattern[(offset + index) % len(pattern)] for offset in range(30)]
        public_days = [day for day, label in zip(days, values) if label == PUBLIC_OFF]
        paid.extend(PaidLeave(employee_id, day) for day in public_days[:2])
        previous = pattern[(index - 1) % len(pattern)]
        previous_days = 0
        while pattern[(index - 1 - previous_days) % len(pattern)] != PUBLIC_OFF:
            previous_days += 1
        work_days = sum(label in active for label in values) if settings.day_count_mode == SPECIFY_WORK_DAYS else None
        employees.append(Employee(employee_id, f"サンプル職員{index + 1}", work_days,
                                  previous_work_days=previous_days, previous_shift=previous, public_holidays=8))
        requests.append(LeaveCandidate(f"R{index + 1:03d}", employee_id, public_days[2], 1))
    counts = tuple((label, pattern.count(label)) for label in active)
    requirements = tuple(DayRequirement(day, sum(count for _, count in counts), shift_counts=counts) for day in days)
    return InputData(tuple(employees), requirements, requests=tuple(requests), settings=settings, paid_leaves=tuple(paid))


def export_leave_sample(folder: str | Path, settings: ShiftSettings | None = None) -> Path:
    """公休・有給の30日例。交代制の未指定時刻は未確認の例として保存する。"""
    if settings is None:
        settings = ShiftSettings(leave_enabled=True, day_count_mode=DERIVE_PUBLIC_DAYS, default_public_holidays=8)
    return export_template(folder, replace(settings, leave_enabled=True))


def generate_sample(folder: str | Path, settings: ShiftSettings | None = None) -> Path:
    """export_templateの別名。引数はファイル名ではなく既存フォルダ。"""
    return export_template(folder, settings=settings)


def export_result(data: InputData, schedule: Schedule, solve_result: SolveResult, folder: str | Path, *, time_limit: float = 30, edited: bool = False, input_path: str | Path | None = None) -> Path:
    """必須条件を再検証した案のみ保存。手修正前の最適性を流用しない。"""
    from .evaluation import evaluate_schedule
    from .planning import configure_input
    from .planning import public_holiday_target, crossing_shifts
    from .planning import attendance_ratio_for_day, attendance_source_for_day
    from .validation import validate_input
    data = configure_input(data)
    issues = validate_input(data)
    if issues:
        raise InputError(issues)
    evaluation = evaluate_schedule(data, schedule)
    if evaluation.issues:
        raise InputError(evaluation.issues)
    if solve_result.status not in (SolverStatus.OPTIMAL, SolverStatus.FEASIBLE):
        raise ValueError("実行可能な計算結果だけ保存できます。シフト案を作り直してください。")
    folder = Path(folder).expanduser().resolve()
    if input_path and folder == Path(input_path).expanduser().resolve().parent:
        raise ValueError("原本とは別のフォルダを選んでください。")
    edited = edited or solve_result.schedule != schedule
    workbook = Workbook()
    workbook.remove(workbook.active)
    days = data.days
    active = SHIFT_MODES[data.settings.mode]
    advanced = data.settings.leave_enabled
    multi = len(active) > 1
    weekday = "月火水木金土日"
    day_names = [f"{d.isoformat()}（{weekday[d.weekday()]}）" for d in days]
    rows = []
    for employee in data.employees:
        counts = [evaluation.employee_shift_counts[employee.id][label] for label in active] if multi else []
        metrics = [evaluation.employee_counts[employee.id], employee.work_days]
        if advanced:
            target = public_holiday_target(employee, data.settings)
            metrics = [evaluation.employee_counts[employee.id], evaluation.employee_actual_work_days[employee.id],
                       employee.work_days if employee.work_days is not None else "公休・有給・明けから決定",
                       target if target is not None else "指定なし", evaluation.employee_public_counts[employee.id],
                       evaluation.employee_paid_counts[employee.id], evaluation.employee_recovery_counts[employee.id]]
        rows.append([employee.id, employee.name, *[assignment_label(schedule[(employee.id, d)], data.settings) for d in days], *metrics, *counts])
    tail = [None] * ((7 if advanced else 2) + (len(active) if multi else 0))
    rows.append(["日別勤務人数", "", *[evaluation.daily_counts[d] for d in days], *tail])
    requirements = {d.day: d.required for d in data.requirements}
    shift_requirements = {r.day: dict(r.shift_counts) for r in data.requirements}
    rows.append(["必要配置人数", "", *[requirements[d] for d in days], *tail])
    if multi:
        for label in active:
            rows.append([f"{label}人数", "", *[evaluation.daily_shift_counts[d][label] for d in days], *tail])
            rows.append([f"{label}必要人数", "", *[shift_requirements[d][label] for d in days], *tail])
    if advanced:
        for label in (PUBLIC_OFF, PAID_LEAVE, RECOVERY):
            rows.append([f"{label}人数", "", *[sum(assignment_label(schedule[e.id, d], data.settings) == label for e in data.employees) for d in days], *tail])
    ws = workbook.create_sheet("勤務表")
    metrics_headers = ["勤務開始回数", "実勤務暦日（明け含む）", "指定勤務開始回数", "公休目標", "公休実績", "有給日数", "明け日数"] if advanced else ["勤務日数", "指定日数"]
    headers = ["職員ID", "氏名", *day_names, *metrics_headers, *([f"{label}日数" for label in active] if multi else [])]
    limit_note = "未設定" if data.settings.max_consecutive_days is None else f"{data.settings.max_consecutive_days}日"
    _table(ws, headers, rows, header_row=3, title=DRAFT,
           note=f"必須条件: 対象期間内は確認済み / {data.settings.mode} / 連続勤務上限: {limit_note} / 勤務時間・休憩時間・資格・法令適合は対象外")
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{len(data.employees) + 3}"
    # Excel上の変更は再計算されない。編集する場合はアプリ内で検証して再保存する。
    note_row = len(rows) + 6
    crossing = crossing_shifts(data.settings) if advanced else frozenset()
    pending_recovery = [e.id for e in data.employees if advanced and assignment_label(schedule[e.id, days[-1]], data.settings) in crossing]
    pending_night = [employee.id for employee in data.employees
                     if data.settings.night_rest and assignment_label(schedule[employee.id, days[-1]], data.settings) in ("夜勤", "深夜勤")
                     and (not advanced or assignment_label(schedule[employee.id, days[-1]], data.settings) not in crossing)]
    next_day = days[-1] + timedelta(days=1)
    note = "変更はアプリ内で実施してください。Excel上の直接編集では条件チェック・採用結果は更新されません。"
    if pending_night:
        note += f"\n【期間外未確認】{next_day.isoformat()}の休みを担当者が確認してください。対象職員ID: {' / '.join(pending_night)}（最終日の夜勤・深夜勤の翌日）。"
    if pending_recovery:
        note += f"\n【期間外未確認】{next_day.isoformat()}は明けとなる勤務です。対象職員ID: {' / '.join(pending_recovery)}。明けを公休・有給として計上しないでください。"
    note_cell = _put(ws, note_row, 1, note)
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=len(headers))
    note_cell.font = Font(name="Yu Gothic", size=11, color=NAVY)
    note_cell.alignment = Alignment(vertical="center", wrap_text=True)
    available = max(1, sum(ws.column_dimensions[get_column_letter(j)].width for j in range(1, len(headers) + 1)) - 3)
    note_lines = sum(max(1, math.ceil(sum(2 if ord(char) > 255 else 1 for char in line) / available)) for line in note.split("\n"))
    ws.row_dimensions[note_row].height = max(32, note_lines * 17 + 10)
    ws = workbook.create_sheet("日別人数")
    daily_headers = ["日付", "必要配置人数", "勤務人数", "差", "判定", *[column for label in active for column in (f"{label}必要人数", f"{label}人数")], "有効出勤率％", "指定根拠", "職員数"]
    daily_rows = [(d, requirements[d], evaluation.daily_counts[d], evaluation.daily_counts[d] - requirements[d], "一致",
                   *[value for label in active for value in (shift_requirements[d][label], evaluation.daily_shift_counts[d][label])],
                   attendance_ratio_for_day(d, data.settings), attendance_source_for_day(d, data.settings), len(data.employees)) for d in days]
    _table(ws, daily_headers, daily_rows, header_row=3, title=DRAFT)
    ws = workbook.create_sheet("希望休の採用結果")
    employee_names = {e.id: e.name for e in data.employees}
    candidates = {}
    for candidate in data.requests:
        candidates.setdefault(candidate.request_id, []).append(f"{candidate.day.isoformat()} 第{candidate.priority}希望")
    independent = data.request_format == "independent" or bool(data.independent_requests)
    kinds = {item.request_id: item.leave_type for item in data.independent_requests}
    originals = {item.request_id: item.original_leave_type or item.leave_type for item in data.independent_requests}
    has_conversion = bool(data.converted_public_request_ids)
    if independent:
        candidates.update({item.request_id: [item.day.isoformat()] for item in data.independent_requests})
    public_total = len(evaluation.request_results)
    paid_results = evaluation.paid_request_results if independent else []
    paid_total = sum(item.leave_type == PAID_REQUEST for item in data.independent_requests)
    if independent:
        request_headers = ["申請ID", "職員ID", "氏名", "希望日", "結果日", "日数", "結果", "区分", *(["元の希望"] if has_conversion else [])]
        request_days = {item.request_id: item.day for item in data.independent_requests}
        request_rows = [[r.request_id, r.employee_id, employee_names[r.employee_id], request_days[r.request_id], r.adopted_day, 1,
                         r.status if kinds[r.request_id] == PAID_REQUEST else "採用" if r.adopted_day is not None else "未採用",
                         kinds[r.request_id], *([originals[r.request_id]] if has_conversion else [])] for r in [*evaluation.request_results, *paid_results]]
    else:
        request_headers = ["申請ID", "職員ID", "氏名", "候補日・順位", "採用日", "採用順位", "採用状況"]
        request_rows = [[r.request_id, r.employee_id, employee_names[r.employee_id], " / ".join(candidates[r.request_id]), r.adopted_day, r.priority, r.status]
                        for r in evaluation.request_results]
    summary = (f"公休希望採用 {evaluation.adopted_count}/{public_total}件・未採用 {public_total - evaluation.adopted_count}件 / 有給確定 {len(paid_results)}/{paid_total}件" if independent else
               f"採用 {evaluation.adopted_count} / {public_total}申請・未採用 {public_total - evaluation.adopted_count}申請")
    _table(ws, request_headers, request_rows, header_row=3, title=DRAFT, note=summary)
    ws = workbook.create_sheet("設定条件")
    metadata = [
        ("用途", DRAFT), ("勤務区分", data.settings.mode), ("区分", " / ".join(assignment_labels(data.settings))),
        ("対象期間", f"{days[0].isoformat()} ～ {days[-1].isoformat()}"),
        ("職員入力件数", len(data.employees)), ("必要人数入力件数", len(data.requirements)),
        ("確定休入力件数", len(data.fixed_off)), ("有給入力件数", len(data.paid_leaves)), ("希望休候補入力件数", len(data.requests)),
        ("希望休入力形式", "横形式・各日独立申請（3枠）" if independent else "従来の代替候補形式"),
        ("独立申請件数", len(data.independent_requests)),
        ("独立公休希望件数", sum(item.leave_type == PUBLIC_REQUEST for item in data.independent_requests)),
        ("独立有給申請件数", paid_total), ("独立有給確定件数", len(paid_results)),
        ("元の有給申請件数", sum(kind == PAID_REQUEST for kind in originals.values())),
        ("有給希望から公休希望へ変更した件数", len(data.converted_public_request_ids)),
        ("公休希望へ変更した申請ID", json.dumps(sorted(data.converted_public_request_ids), ensure_ascii=False)),
        ("有給申請の担当者確認", f"全件反映済み・有給確定{len(paid_results)}件・公休希望への変更{len(data.converted_public_request_ids)}件" if paid_total or has_conversion else "対象申請なし"),
        ("確認済み有給申請ID", json.dumps(sorted(data.confirmed_paid_request_ids), ensure_ascii=False)),
        ("有給確認の範囲", "アプリ内で担当者が明示確認した申請IDを記録。人事上の承認・残高更新を証明するものではない。"),
        ("既存の固定有給", "有給シート等の既存固定有給は希望の配分では変更しない。"),
        ("元からの固定有給件数", sum(not item.request_id for item in data.paid_leaves)),
        ("希望休申請件数", len(evaluation.request_results)),
        ("結果セル件数", len(schedule)), ("勤務セル件数", sum(is_working(value, data.settings) for value in schedule.values())),
        ("実勤務暦日セル件数", sum(is_actual_working(value, data.settings) for value in schedule.values())),
        *[(f"{label}セル件数", sum(assignment_label(value, data.settings) == label for value in schedule.values())) for label in assignment_labels(data.settings)],
        ("希望休採用件数", evaluation.adopted_count), *([] if independent else [("採用順位合計", evaluation.priority_sum)]),
        ("計算状態", "手修正後・必須条件確認済み（最適性は未確認）" if edited else solve_result.status.value),
        ("採用件数の最適性", "未確認" if edited or not solve_result.count_optimal else "確認済み"),
        *([] if independent else [("順位合計の最適性", "未確認" if edited or not solve_result.priority_optimal else "確認済み")]),
        ("全体制限時間（秒）", time_limit), ("計算経過時間（秒）", round(solve_result.elapsed_seconds, 3)),
        ("必須条件", "対象期間内の各人1日1区分・日別総人数と各区分人数一致・選択した日数条件（勤務開始回数または公休枠）・確定休と有給遵守・明け・連続勤務上限・夜勤後休み・禁止勤務順序（有効にした条件）"),
        ("希望条件", "3枠は独立した公休希望。採用件数を最大化。有給は担当者確認後の固定条件。" if independent else "採用申請数を最大化した後、その件数を保ち順位合計を最小化。同一職員・同一日の二重採用なし。"),
        ("必要人数の指定方法", "日付・週別の調整あり" if data.settings.weekly_ratios or data.settings.date_ratios else "Excel人数" if data.settings.weekday_ratios is None else "曜日別比率"),
        ("出勤率の適用順序", "日付指定 > 月曜日から日曜日の週指定 > 曜日指定 > Excel人数"),
        ("曜日別比率（月〜日％）", "使用しない" if data.settings.weekday_ratios is None else " / ".join(map(str, data.settings.weekday_ratios))),
        ("週別出勤率", json.dumps(settings_to_payload(data.settings)["weekly_ratios"], ensure_ascii=False)),
        ("日付別出勤率", json.dumps(settings_to_payload(data.settings)["date_ratios"], ensure_ascii=False)),
        ("比率からの人数", "日付・週・曜日の優先順で選んだ有効出勤率を使用し、全職員数×出勤率÷100を切り上げ。勤務日数指定方式の個人勤務開始回数は変更しない。公休算出方式は公休・有給・明け条件から勤務を決める。"),
        ("区分配分", " / ".join(f"{label}:{weight}" for label, weight in zip(active, data.settings.shift_weights or (1,) * len(active)))),
        ("区分への丸め", "比率の指定がない日は明示人数を使用。自動配分は最大剰余法。同率は区分の表示順。日付・週・曜日いずれかの出勤率を適用する日は、明示人数も再配分。"),
        ("連続勤務上限", "未設定" if data.settings.max_consecutive_days is None else data.settings.max_consecutive_days),
        ("夜勤後休み", ("有効（夜勤・深夜勤の翌日に新規勤務開始不可。日跨ぎの翌日は明け。明けのさらに翌日の公休は自動強制しない）" if advanced else "有効（夜勤・深夜勤の翌日は休み。連続夜勤も不可）") if data.settings.night_rest else "無効"),
        ("禁止勤務順序", json.dumps(data.settings.forbidden_transitions, ensure_ascii=False)),
        ("公休・有給・明け", "有効" if advanced else "無効（従来の休み区分）"),
        ("日数指定方法", data.settings.day_count_mode),
        ("共通公休日数", data.settings.default_public_holidays if data.settings.default_public_holidays is not None else "未設定"),
        ("公休日数の根拠", "個別公休日数を優先し、未設定なら共通公休日数。公休から算出方式のみ目標に使用。勤務日数指定と独立した二重条件にはしない。"),
        *[(f"公休基準/{employee.id}", (public_holiday_target(employee, data.settings) - employee.public_holiday_adjustment) if public_holiday_target(employee, data.settings) is not None else "指定なし") for employee in data.employees],
        *[(f"公休への振替加算/{employee.id}", employee.public_holiday_adjustment) for employee in data.employees],
        *[(f"適用公休目標/{employee.id}", public_holiday_target(employee, data.settings) if public_holiday_target(employee, data.settings) is not None else "指定なし") for employee in data.employees],
        *([(f"休暇合計（公休＋有給）/{employee.id}", evaluation.employee_public_counts[employee.id] + evaluation.employee_paid_counts[employee.id]) for employee in data.employees] if advanced else []),
        *([(f"暦日合計/{employee.id}", f"勤務開始{evaluation.employee_counts[employee.id]}＋明け{evaluation.employee_recovery_counts[employee.id]}＋公休{evaluation.employee_public_counts[employee.id]}＋有給{evaluation.employee_paid_counts[employee.id]}＝{len(days)}日") for employee in data.employees] if advanced else []),
        ("有給の扱い", "日付固定の全日有給。公休や明けに置換しない。有給残高は更新しない。残高と承認は外部の正本を参照。"),
        ("明けの扱い", "日跨ぎ勤務の翌日。公休・有給と別区分。勤務開始人数には含めず、実勤務暦日と連続勤務には含める。" if advanced else "従来方式は明けを使用しない"),
        ("勤務時刻確認", "確認済み" if data.settings.shift_times_confirmed else "未確認"),
        *[(f"勤務時刻/{item.label}", f"{_format_clock(item.start_minute)}〜{'翌' if item.end_minute > 1440 else ''}{_format_clock(item.end_minute if item.end_minute <= 1440 else item.end_minute - 1440)}") for item in data.settings.shift_times],
        ("夜勤・順序判定の単位", "確認した時刻から日跨ぎを識別し、各日セル間の区分・明けを検証。勤務間隔の法令適合を判定する機能ではない。"),
        ("前期間の扱い", "前期間の連続勤務日数を期間先頭へ接続する。列なし・空欄・使用しないは0日＝対象期間の前日は休みと仮定。高度方式の旧履歴区分「休み」は前日の公休相当へ正規化し、禁止順序を判定する。対象期間内の公休日数には計上しない。"),
        *[(f"前期間連続勤務日数/{e.id}", e.previous_work_days) for e in data.employees],
        *[(f"前日の勤務区分/{e.id}", e.previous_shift or (OFF if e.previous_work_days == 0 else "日勤（履歴区分未記入）")) for e in data.employees],
        ("期間外の翌日確認", "未確認。最終日の勤務に対する翌日の休み・明けを担当者が確認してください。" if pending_night or pending_recovery else "最終日の勤務に伴う翌日休み・明けの確認対象なし。対象期間より後の一般的な制約は未検証。"),
        ("翌日確認対象日", next_day if pending_night or pending_recovery else "なし"),
        ("翌日休み確認対象の職員ID", " / ".join(pending_night) if pending_night else "なし"),
        ("翌日明け確認対象の職員ID", " / ".join(pending_recovery) if pending_recovery else "なし"),
        ("対象外", ("夜勤明け専用区分（旧方式）・" if not advanced else "") + "有給残高更新・半日休・時間休・勤務間隔や勤務時間・休憩・資格別配置・法令適合の自動判定。高度方式では明けを含め暦日の実勤務で連勤を数える。"),
        ("手修正", "あり" if edited else "なし"),
        ("入力原本", str(Path(input_path).resolve()) if input_path else "サンプルまたはメモリ入力"),
        ("作成日時", datetime.now(timezone(timedelta(hours=9))).isoformat(timespec="seconds")),
        ("備考", "実際の配置は担当者が確認してください。Excelの直接編集はアプリ内検証へ反映されません。"),
    ]
    # 元ファイル・シート・行を残し、列対応後も入力の出所を追えるようにする。
    source_rows = []
    for label, records in (("職員", data.employees), ("必要人数", data.requirements), ("確定休", data.fixed_off), ("希望休", data.requests), ("有給", data.paid_leaves), ("独立申請", data.independent_requests)):
        for record in records:
            src = record.source
            source_rows.append((f"入力出所/{label}", f"{src.file} / {src.sheet} / 元行{src.row}"))
    _table(ws, ["項目", "値"], metadata + source_rows, header_row=3, title=DRAFT)
    return _save_unique(workbook, folder, "sinta_shift_result.xlsx")
