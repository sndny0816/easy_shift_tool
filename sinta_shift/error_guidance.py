"""Issueの原文を保ち、利用者が行う修正手順と明示的な移動先を返す。

GUI・ファイル操作・条件変更は行わない。分類不能なエラーの原因や、
条件不成立の具体的な原因を、このモジュールで推測しない。
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from .models import Issue, InputData, InputError, Source, SolveResult, SolverStatus, PUBLIC_DAYS


@dataclass(frozen=True)
class ErrorGuidance:
    title: str
    location: str
    reason: str
    steps: tuple[str, ...]
    next_action: str
    target: str
    action_label: str
    category: str


ACTIONS = {
    "input": "ファイル・列対応を確認",
    "settings": "勤務条件を確認",
    "leave": "公休・有給の配分を確認",
    "schedule": "該当する勤務表セルを確認",
    "save": "保存先を選び直す",
    "run": "計算条件・制限時間を確認",
    "support": "元の理由・技術詳細を確認",
}
INPUT_NEXT = "Excelを修正して保存した後、［Excelを選ぶ］で修正版を読み直し、［条件を確認］から再実行してください。"
SETTING_NEXT = "設定を直した後、［人数をプレビュー・条件確認］を実行し、条件がそろってからシフト案を作ってください。"
MANUAL_NEXT = "変更後の違反一覧を確認してください。全必須条件を満たすまで完成版は保存できません。"
SAVE_NEXT = "保存先を選び直し、［結果を保存］を再実行してください。入力データの人数・日数は変更しません。"


def _name(value):
    return str(value).replace("\\", "/").rstrip("/").rsplit("/", 1)[-1]


def _location(issue: Issue, target: str, input_path=None):
    source = issue.source
    if target == "schedule":
        parts = ["アプリ内の勤務表"]
        if source.sheet == "勤務表" and re.fullmatch(r"[A-Za-z]+[1-9][0-9]*", source.column):
            parts.append(f"Excel出力時のセル {source.column}（参考）")
        if issue.employee_id:
            parts.append(f"職員ID {issue.employee_id}")
        if issue.day is not None:
            parts.append(issue.day.isoformat() if hasattr(issue.day, "isoformat") else str(issue.day))
        return " / ".join(parts)
    parts = []
    if source.file:
        parts.append(_name(source.file))
    elif input_path is not None and target == "input":
        parts.append(f"{_name(input_path)}（選択中の入力）")
    if source.sheet:
        parts.append(f"シート「{source.sheet}」")
    if source.row:
        parts.append(f"元行 {source.row}")
    if source.column:
        cell = bool(re.fullmatch(r"[A-Za-z]+[1-9][0-9]*", source.column))
        parts.append(f"{'セル' if cell else '項目'} {source.column}")
    if issue.employee_id:
        parts.append(f"職員ID {issue.employee_id}")
    if issue.day is not None:
        parts.append(issue.day.isoformat() if hasattr(issue.day, "isoformat") else str(issue.day))
    if parts:
        return " / ".join(parts)
    if target == "settings":
        return "勤務条件（元のファイル・行・セルは特定されていません）"
    return "場所は特定されていません。元の理由を確認してください。"


def issues_for_solve_result(result: SolveResult) -> tuple[Issue, ...]:
    """既存の具体的Issueを優先し、結果がない失敗状態だけを補う。"""
    if result.issues:
        return tuple(result.issues)
    if result.status in (SolverStatus.UNKNOWN, SolverStatus.INFEASIBLE, SolverStatus.MODEL_ERROR):
        return (Issue(result.status.value, result.message),)
    return ()


def issues_for_exception(exc: Exception, *, operation=None) -> tuple[Issue, ...]:
    """例外をGUIの既存Issue一覧へ渡す補助。原文と実際のfilenameを保持。"""
    if isinstance(exc, InputError):
        return tuple(exc.issues)
    kind = ("保存エラー" if operation == "save" else "読込エラー" if operation == "read" else
            "ファイルアクセス" if isinstance(exc, OSError) else "処理エラー")
    return (Issue(kind, str(exc), Source(file=str(getattr(exc, "filename", "") or ""))),)


def guidance_for_issue(issue: Issue, *, data: InputData | None = None,
                       has_schedule: bool = False, input_path: str | Path | None = None,
                       operation: str | None = None) -> ErrorGuidance:
    """reasonはIssue.messageと完全一致。手順は原本・設定・案の修正を区別する。"""
    message, kind, source = issue.message, issue.kind, issue.source
    settings = data.settings if data is not None else None
    from_excel = source.sheet in ("勤務条件", "勤務時刻", "出勤率調整") and bool(source.row or re.fullmatch(r"[A-Za-z]+[1-9][0-9]*", source.column))

    def make(title, category, target, steps, next_action=None):
        next_action = next_action or (INPUT_NEXT if target == "input" else SETTING_NEXT if target == "settings" else
                                     MANUAL_NEXT if target == "schedule" else SAVE_NEXT if target == "save" else
                                     "表示された元の理由を確認し、修正または担当者への確認後に再実行してください。")
        location_path = None if operation in ("mapping_load", "mapping_save", "settings_load", "settings_save", "template_save") else input_path
        return ErrorGuidance(title, _location(issue, target, location_path), message, tuple(steps),
                             next_action, target, ACTIONS[target], category)

    def input_guide(title, category, *steps):
        return make(title, category, "input", (*steps, "元ファイルは残し、修正版を別名の.xlsxとして保存してください。"))

    def setting_guide(title, category, *steps):
        if from_excel:
            return input_guide(title, category, *steps, "設定をファイルで管理している場合は、該当する設定表またはJSONも修正してください。")
        return make(title, category, "settings", (*steps,
            "ファイルから復元した設定を使っている場合は、元の勤務条件シート・勤務時刻シート・設定JSONも確認してください。"))

    if operation == "template_save":
        guide = make("入力ひな形の保存先を確認してください", "template_save", "save", (
            "ひな形を保存する、存在する書き込み可能なフォルダを選んでください。",
            "使用中・アクセス拒否の場合は対象Excelを閉じ、フォルダの権限を確認してください。"),
            "［入力ひな形を保存］から保存先を選び直してください。")
        return ErrorGuidance(guide.title, guide.location, guide.reason, guide.steps,
                             guide.next_action, guide.target, "入力ひな形の保存先を選び直す", guide.category)

    # JSONの再選択・保存は、入力Excelの再読込と混同しない。
    if operation in ("mapping_load", "mapping_save", "settings_load", "settings_save"):
        mapping = operation.startswith("mapping_")
        saving = operation.endswith("_save")
        name = "列対応" if mapping else "勤務条件"
        action = "保存先を選び直す" if saving else "JSONファイルを選び直す"
        guide = make(f"{name}JSONの{'保存先' if saving else '読込内容'}を確認してください", operation,
            "input" if mapping else "settings", (
                f"{name}用のJSONであることと、元の理由に示された項目を確認してください。",
                "保存先の存在と書き込み権限を確認し、JSONの保存先を選び直してください。" if saving else
                "JSONを修正して保存するか、正しいJSONファイルを選び直してください。入力Excelの値を変える必要はありません。"),
            f"［{name}{'を保存' if saving else 'を読込'}］から{action}を行ってください。")
        return ErrorGuidance(guide.title, guide.location, guide.reason, guide.steps,
                             guide.next_action, guide.target, action, guide.category)

    if (operation == "solve" and "制限時間" in message and
            kind not in (SolverStatus.UNKNOWN.value, SolverStatus.INFEASIBLE.value, SolverStatus.MODEL_ERROR.value)):
        return make("計算の制限時間を入力してください", "time_limit", "run", (
            "制限時間を0より大きく300以下の秒数で入力してください。初期値は30秒です。",
            "入力値の不備で計算を開始できていません。勤務条件や人数を変更する必要はありません。"),
            "制限時間を直して［シフト案を作る］を実行してください。")

    # 時間切れを『条件不成立』や入力誤りへ読み替えない。
    if kind == SolverStatus.UNKNOWN.value or "制限時間内に案が見つかりません" in message:
        return make("時間内に案を発見できませんでした", "timeout", "run", (
            "計算時間を延ばして再実行してください（最大300秒）。",
            "案がないと確定した状態ではありません。確定休・承認済み有給はそのままにしてください。"),
            "制限時間を確認して［シフト案を作る］を再実行してください。")
    if kind == SolverStatus.MODEL_ERROR.value:
        return make("計算モデルを処理できませんでした", "model_error", "support", (
            "元の理由・技術詳細を全文確認し、入力・設定・アプリのバージョンを控えて保守担当者へ確認してください。",
            "条件不成立の判定ではありません。勤務条件を変えずに確認してください。"))

    # 保存失敗は勤務条件の修正では解決しない。
    if "実行可能な計算結果だけ保存" in message:
        return make("先にシフト案を作成してください", "infeasible", "run", (
            "実行可能な案がないため、この結果は完成版として保存できません。",
            "計算状態と、ほかの入力エラー・条件違反の案内を確認してください。"))
    if ("原本とは別" in message or "保存先として存在するフォルダ" in message or
            kind == "保存エラー" or (operation == "save" and kind in ("処理エラー", "ファイルアクセス"))):
        return make("結果の保存先を確認してください", "save_path", "save", (
            "原本と別の、存在する書き込み可能なフォルダを選んでください。",
            "理由に使用中・アクセス拒否がある場合は、対象Excelを閉じ、保存先への書き込み権限を確認してください。",
            "理由に空き容量不足がある場合は、別の保存先または空き容量を確認してください。入力や必須条件は変更しません。"))
    if kind in ("ファイルアクセス", "読込エラー") or any(word in message for word in ("Permission denied", "Access is denied", "アクセスが拒否", "[Errno 13]")):
        target = "save" if operation == "save" else "input"
        return make("ファイルへのアクセスを確認してください", "file_access", target, (
            "元の理由に示されたファイルの存在と、読み取り・書き込み権限を確認してください。使用中の場合は対象ファイルを閉じてください。",
            "アクセスできる場所からファイルを選び直してください。勤務日数や必要人数は変更しません。"),
            SAVE_NEXT if target == "save" else "ファイルを開ける状態にした後、［Excelを選ぶ］から読み直してください。")

    if kind in ("数式", "セルエラー"):
        return input_guide("数式・Excelエラーを値に直してください", "formula",
            "表示されたセルまたは使用列をExcelで確認してください。古い計算結果は利用しません。",
            "数式の場合は、計算結果を確認してコピーし、別の入力用ファイルへ［形式を選択して貼り付け］→［値］で貼り付けてください。",
            "Excelエラーの場合は、その原因をExcel側で解消し、確認した値を記入してください。")
    if kind in ("列", "見出し", "対応設定", "シート"):
        return make("シート・見出し・列対応を確認してください", "columns", "input", (
            "［読み込み・列対応］で、使用するシートと見出し行を選び、［見出しを反映］を押してください。",
            "必要項目を元の列へ対応付けてください。同じ元列を複数項目へ使わず、重複した見出しはExcelで直してください。",
            "Excel自体を修正した場合は別名で保存し、修正版を選び直してください。対応だけ直した場合は条件確認を再実行してください。"),
            "対応設定を反映して［条件を確認］を実行してください。Excelを修正した場合は先にファイルを読み直してください。")
    if kind in ("申請形式", "希望休形式"):
        return make("希望休の形式を選び直してください", "request_format", "input", (
            "横形式の希望日1・2・3は3日分の独立申請です。従来の同じ申請IDの候補日は1日分の代替候補です。",
            "意図した形式を明示し、その形式の列を対応付けてください。形式を推測して休暇日数を変えないでください。"),
            "形式と列対応を直して［条件を確認］を実行してください。")
    if kind == "有給申請未確認":
        return make("有給希望の配分が未決定です", "leave_confirmation", "leave", (
            "［公休・有給］の一覧で、対象の職員ID・日付と元の希望を確認してください。",
            "職場の承認内容を確認したうえで、有給に確定するか公休希望へ変更するかを明示してください。保留のままでは計算できません。",
            "有給残高の更新や人事上の承認はこのアプリでは行いません。"),
            "配分を明示した後、［条件を確認］から再実行してください。")

    manual = {
        "日別人数違反": ("日別人数を合わせてください", "manual_day_count", "該当日の列を確認し、勤務開始人数を必要人数と照合してください。明け・公休・有給は配置人数に含めません。"),
        "勤務区分人数違反": ("勤務区分別の人数を合わせてください", "manual_shift_count", "日勤・夜勤など区分ごとの実際人数と必要人数を確認し、該当日の区分を修正してください。"),
        "勤務日数違反": ("個人の勤務開始回数を確認してください", "manual_work_days", "該当職員の行を確認し、勤務開始回数を指定日数と照合してください。有給・明けで勤務開始数を補わないでください。"),
        "公休日数違反": ("公休と有給・明けを分けてください", "manual_public_days", "該当職員の公休日数を目標と照合してください。有給・明けは公休へ数えません。元有給から公休へ変更した件数も確認してください。"),
        "連続勤務違反": ("連続勤務の並びを修正してください", "manual_consecutive", "前期間からの連勤も含め、実労働のある暦日を確認してください。明けも連勤へ数え、必要な休みを配置してください。"),
        "明け違反": ("日跨ぎ勤務と明けの対応を確認してください", "manual_recovery", "日跨ぎ勤務の翌日は明け、前日に日跨ぎ勤務がなければ明け不可です。初日は前日の履歴も確認し、公休・有給と重ねないでください。"),
        "夜勤翌日違反": ("夜勤翌日の勤務開始を修正してください", "manual_night_rest", "夜勤・深夜勤の翌日に新しい勤務を始めていないか確認してください。日跨ぎ勤務なら翌日は明けです。"),
        "勤務順序違反": ("禁止された勤務順序を修正してください", "manual_transition", "前日と当日の両セルを、設定した禁止順序と照合してください。期間初日は前日の履歴を確認してください。"),
        "確定休違反": ("確定休を守る区分へ戻してください", "manual_fixed_off", "確定休の日を勤務に変えず、ほかの勤務配置を調整してください。高度方式の確定休は公休として扱います。"),
        "有給違反": ("確定済み有給へ戻してください", "manual_paid", "元の有給一覧または担当者が確定した有給の日付を確認し、そのセルを有給へ戻してください。"),
        "未承認有給": ("有給を勝手に追加しないでください", "manual_paid", "有給一覧・明示確認済み申請にない有給セルを見直してください。有給の承認・配分を確認してから案を作り直してください。"),
        "区分違反": ("有効な勤務区分を選んでください", "manual_shift_count", "日付セルの区分メニューから、選択中の勤務モードの区分を1つ選んでください。"),
        "範囲違反": ("対象外のセルを見直してください", "manual_shift_count", "対象職員と期間の範囲を確認し、現在の入力から案を作り直してください。"),
        "セル確認": ("表示されたセルと行・列を確認してください", "manual_day_count", "該当セルだけでなく、職員の行と日付の列の人数・勤務日数を照合してください。"),
    }
    if kind in manual:
        title, category, step = manual[kind]
        if has_schedule:
            return make(title, category, "schedule", (step, "確定休・承認済み有給や必要人数の基準を勝手に変えて、違反を隠さないでください。"))
        return make(title, category, "run", (step, "現在操作できる勤務表がないため、入力と設定を確認して新しい案を作成してください。"))

    if "勤務日数合計" in message or ("日数が一致しません" in message and "勤務開始" in message):
        derived = settings is not None and settings.leave_enabled and settings.day_count_mode == PUBLIC_DAYS
        overrides = settings is not None and bool(settings.weekday_ratios is not None or settings.weekly_ratios or settings.date_ratios)
        if derived:
            crossing = any(isinstance(getattr(t, "end_minute", None), int) and t.end_minute > 1440 for t in settings.shift_times)
            formula = ("勤務開始＋明け＋公休＋確定有給＝全員の対象暦日数を確認してください。日跨ぎがあるため、暦日数から公休と有給だけを引いて勤務開始数を決めないでください。" if crossing else
                       "対象期間の日数−公休目標−確定有給＝勤務開始回数を確認してください。公休へ変更した元有給希望の調整分も公休目標に含めます。")
            return setting_guide("公休・有給・明けを含む日数を突き合わせてください", "totals", formula,
                "公休算出方式では入力勤務日数を独立した基準として直しません。共通/個別公休、確定有給、勤務時刻、必要人数の誤記を確認してください。",
                "必要人数・確定休・承認済み有給の業務上の基準を勝手に減らして一致させないでください。")
        steps = ("職員表の対象期間勤務日数（勤務開始回数）の合計と、日別必要人数の合計を照合してください。",
                 "実際の勤務予定と配置基準を確認し、誤記が確認できた側だけを修正してください。")
        if overrides:
            return setting_guide("勤務日数と有効な配置人数を突き合わせてください", "totals", *steps,
                "曜日・週・日付の出勤率上書きがあるため、元Excelの人数ではなく画面の有効な必要人数を使って比較してください。",
                "Excelの個人勤務日数を直す場合は、修正版を保存して読み直してください。")
        return input_guide("職員と必要人数の合計を合わせてください", "totals", *steps)

    if kind in ("条件不成立", SolverStatus.INFEASIBLE.value):
        known_capacity = "勤務可能" in message and "必要" in message
        return make("必要人数と配置可能な要員を確認してください" if known_capacity else "必須条件を同時に満たせません", "capacity" if known_capacity else "infeasible", "settings", (
            "理由に示された日付・職員と、必要人数・配置可能人数を確認してください。" if known_capacity else "この結果だけでは、どの条件の組合せが原因かは特定できません。",
            "確定休・確定有給・前日履歴の誤記がないか原入力と照合し、実際の配置基準と要員確保を担当者が確認してください。",
            "確定休・承認済み有給を自動解除したり、必要人数・連勤条件を勝手に緩めたりして成立させないでください。"),
            "誤記をExcelで直した場合は保存して読み直し、勤務条件を直した場合は条件確認後に再計算してください。")

    if kind in ("勤務時刻",) or (kind == "勤務条件" and "時刻" in message) or any(word in message for word in ("勤務時刻", "時刻を確認済み", "終了24:00", "開始0〜1439")):
        return setting_guide("勤務時刻と日跨ぎを確認してください", "time_settings",
            "有効な各勤務区分の開始・終了と翌日終了の指定を確認してください。入力済み時刻を職場の実際の時刻と照合してから確認済みにします。",
            "当日24:00は終了24:00＋当日、翌日00:00は終了00:00＋翌日を選びます。終了は開始より後で24時間以内にしてください。",
            "時刻は日跨ぎ・明けの分類に使用します。確認せずにチェックだけ入れて計算を通さないでください。")
    if kind == "出勤率" or any(word in message for word in ("出勤率", "曜日別", "週指定", "日付指定", "比率％", "配置比率")):
        return setting_guide("出勤率の指定と優先順位を確認してください", "attendance_settings",
            "各比率は0〜100％の有限の数値、週指定は月曜日を開始日として設定してください。重複・期間外の指定を直してください。",
            "日付指定→週指定→曜日指定→元Excel人数の順に優先されます。意図しない上書きはその指定を解除してください。",
            "勤務開始人数は職員数×有効な比率を切り上げます。比率を変えても個人の勤務日数は自動変更しません。")
    if "連続勤務上限" in message:
        return setting_guide("連勤上限の値を確認してください", "consecutive_settings",
            "職場で決めた上限を1以上の整数で入力してください。明けを含む実労働暦日の連勤として扱います。",
            "成立させるためだけに上限を引き上げず、勤務配置と前期間履歴を確認してください。")
    if "禁止" in message and "勤務順序" in message:
        return setting_guide("禁止する勤務順序を確認してください", "sequence_settings",
            "現在の勤務モードと休み区分から、前日→翌日の組を指定してください。重複と対象外区分を直してください。",
            "職場で必要な禁止順序を、計算を通すためだけに解除しないでください。")

    if "期間開始前" in message or ("前日" in message and "履歴" in message):
        return input_guide("前期間の連勤・前日区分を照合してください", "history",
            "職員表の前期間連続勤務日数と前日の勤務区分を、前期間の実記録で確認してください。",
            "公休・有給など非勤務の前日と、連勤1日以上を同時指定しないでください。明けは実労働のある暦日として数えます。",
            "過去の勤務や承認内容を変えて条件を合わせず、誤記・不明な履歴を確認してください。")
    if kind in ("有給確認設定", "有給配分"):
        return make("有給希望の決定と対象IDを確認してください", "leave_choices", "leave", (
            "元の有給希望だけを対象にし、同じ申請を有給と公休の両方へ指定しないでください。",
            "元の固定有給はこの配分では変更しません。対象の職員ID・日付・申請IDを確認してください。",
            "未決定分に対する今回の件数と、既に確定した決定を確認してから明示確定してください。"),
            "配分を直して条件確認を再実行してください。入力の重複がある場合はExcelを修正・保存して読み直してください。")
    if "重複" in message or "同じ職員・同じ日" in message or "同じ職員・日付" in message:
        return input_guide("重複した入力を確認してください", "ids" if "ID" in message else "dates",
            "表示された行と、同じIDまたは同じ職員・日付を持つ行をExcelで照合してください。",
            "同じ休みを公休と有給の両方に記入しないでください。誤った重複だけを修正し、承認済みの休暇を一括で削除しないでください。",
            "従来形式の同じ申請IDの代替候補は有効です。同一候補日の重複や、申請ID内の別職員混在を直してください。")
    if "未知の職員ID" in message or "職員IDを文字列" in message or ("職員ID" in message and ("空欄" in message or "読み取" in message)):
        return input_guide("職員IDを職員表と合わせてください", "ids",
            "申請・確定休・有給の職員IDを職員表のIDと照合し、空欄・誤記・別職員のIDを直してください。",
            "ID列は文字列で入力し、先頭ゼロと文字列NAを保持してください。既に失われたゼロを推測で補わないでください。")
    if "申請ID" in message and any(word in message for word in ("空欄", "混在", "不正")):
        return input_guide("申請IDと対象職員を確認してください", "ids",
            "申請IDを空欄にせず、同じ申請IDの代替候補は同じ職員に統一してください。",
            "異なる申請を同じIDへ結合せず、職員IDと候補日を照合してください。")
    if (kind == "変換" and any(word in message for word in ("日付", "候補日", "希望日", "時刻を含めず", "年を含む"))) or any(word in message for word in ("対象期間外", "連続していません", "存在する日付", "年を含む日付", "日付が重複")):
        return input_guide("日付と対象期間を確認してください", "dates",
            "年を含む完全な日付（例：2026-11-07）で記入し、必要人数表の対象期間と照合してください。",
            "必要人数は対象期間の全日を連続して指定し、期間外の日付や同日の誤った重複を直してください。")
    if "両方記入" in message:
        return input_guide("希望日と区分をセットで入力してください", "request_format",
            "希望日1と区分1など、同じ番号の2セルを両方記入してください。未申請の枠は両方空欄にします。",
            "区分は公休希望または有給申請を選びます。記載だけで有給確定にはなりません。")

    if kind == "勤務条件":
        if source.row and source.sheet not in ("勤務条件", "勤務時刻", "出勤率調整"):
            return input_guide("入力表の人数・日数を確認してください", "leave_counts" if "公休" in message or "有給" in message else "counts",
                "表示された元行と項目を確認し、人数・日数・区分別人数を元の理由に合う値へ直してください。",
                "公休算出方式で個別値を使う場合は職員表の公休日数、共通値を使う場合は勤務条件も確認してください。")
        return setting_guide("勤務条件の入力値を確認してください", "leave_counts" if "公休" in message or "有給" in message else "counts",
            "元の理由に示された項目を確認し、現在の勤務モード・日数指定方式に合う値を入力してください。",
            "公休から算出する場合は共通または個別公休日数を指定します。勤務日数直接方式では入力勤務開始回数を基準にします。")
    if operation == "mapping":
        return make("画面の列対応を確認してください", "columns", "input", (
            "読み込み・列対応でシート、見出し行、必要項目の対応先を確認してください。",
            "元の理由に示された画面の入力を直し、列対応を反映してください。"),
            "列対応を直して［条件を確認］を実行してください。")
    if operation == "settings":
        return make("画面の勤務条件を確認してください", "settings", "settings", (
            "勤務条件画面で、元の理由に示された入力値を確認してください。",
            "数値の空欄・入力形式・選択内容を直してください。業務上の基準は確認した値を使用します。"))
    if kind == "変換" or any(word in message for word in ("整数", "空欄", "0件")):
        return input_guide("入力値の型・空欄を確認してください", "counts",
            "理由に示された項目へ有効な整数・文字列を記入してください。欠損や変換失敗を0で埋めないでください。",
            "人数・勤務日数・順位の範囲は、元の理由に示された条件と照合してください。")
    if kind in ("読込", "形式"):
        return input_guide("入力Excelを開き直してください", "file_access",
            "対象の.xlsxをExcelで開けるか確認し、値の表として別名保存してください。.xls・マクロ・PDFは対象外です。",
            "元の理由にパスワード・破損・アクセス制限が示されている場合は、管理者へ確認してください。推測で勤務条件を変えないでください。")
    if kind in ("入力不備", "独立申請") or (source.row and source.sheet != "勤務表" and operation in (None, "read")):
        return input_guide("表示された入力箇所を照合してください", "unknown",
            "元の理由と表示されたシート・元行・項目をExcelで照合してください。この情報だけで別の原因は断定できません。",
            "誤記が確認できた箇所だけを修正してください。判断できない場合は入力の担当者へ確認してください。")
    return make("元の理由を確認してください", "unknown", "support", (
        "このエラーの具体的な原因を案内だけで特定できません。元の理由・技術詳細を全文確認してください。",
        "再現した操作と設定を控え、必要に応じて保守担当者へ確認してください。確定済みの勤務条件を推測で変更しないでください。"))
