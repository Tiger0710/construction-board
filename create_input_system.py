"""入力システム生成: ガントチャート + 日次入力 + VBA自動同期

VBA付きテンプレート(template/gantt_template.xlsm)をベースに、
月別のガントチャート入力ファイルを生成する。

使い方:
  python create_input_system.py                       # 当月、空データ
  python create_input_system.py --month 2604          # 4月分
  python create_input_system.py --persons 田中,山本,鈴木 --month 2604
"""
import os
import sys
import argparse
import calendar
import datetime
import shutil

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

import config

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template", "gantt_template.xlsm")

WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]

# ---------- スタイル ----------
NAVY = "1A2E5A"
HEADER_FONT = Font(name="Noto Sans JP", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_FONT = Font(name="Noto Sans JP", size=11)
CELL_ALIGN = Alignment(vertical="center", wrap_text=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
GANTT_DATE_FONT = Font(name="Noto Sans JP", size=9, bold=True)
GANTT_WD_FONT = Font(name="Noto Sans JP", size=9, color="555555")
GANTT_LABEL_FONT = Font(name="Noto Sans JP", size=10, bold=True)
LABEL_FILL = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
SAT_FILL = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
SUN_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
BAR_FILL = PatternFill(start_color="93C5FD", end_color="93C5FD", fill_type="solid")

# ガントチャート列定義
GANTT_COLS = [
    ("客先", 12), ("工事件名", 48), ("現場担当者", 10), ("安品担当者", 10),
    ("協力会社名", 16), ("協力会社担当者", 10), ("開始", 10), ("終了", 10),
]
GANTT_LEFT_COLS = len(GANTT_COLS)  # 8列 (A-H)
DATE_START_COL = GANTT_LEFT_COLS + 1  # I列から日付
GANTT_EMPTY_ROWS = 30


def parse_month(month_str):
    """'YYMM' 文字列 → (year, month) タプル"""
    yy = int(month_str[:2])
    mm = int(month_str[2:])
    year = 2000 + yy
    return year, mm


def get_month_dates(year, month):
    """指定月の全日付リスト"""
    days_in_month = calendar.monthrange(year, month)[1]
    return [datetime.date(year, month, d) for d in range(1, days_in_month + 1)]


def apply_header(ws, row, col_start, col_end, fill=None):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = fill or HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER


def create_input_file(person_name, month_str):
    """VBA付き入力ファイル生成 (テンプレートベース)"""
    year, month = parse_month(month_str)
    dates = get_month_dates(year, month)
    gantt_days = len(dates)

    # テンプレート読み込み (VBA保持)
    if not os.path.exists(TEMPLATE_PATH):
        print(f"  ERROR: テンプレートなし: {TEMPLATE_PATH}")
        print("  → python create_template.py を実行してください")
        return None

    wb = load_workbook(TEMPLATE_PATH, keep_vba=True)

    # ==================== ガントチャート ====================
    ws = wb["ガントチャート"]

    # 列幅
    for i, (_, width) in enumerate(GANTT_COLS):
        ws.column_dimensions[get_column_letter(i + 1)].width = width

    # Row 1: 日付
    ws.row_dimensions[1].height = 28
    for c in range(1, GANTT_LEFT_COLS + 1):
        ws.cell(row=1, column=c).border = THIN_BORDER

    for i, d in enumerate(dates):
        col = DATE_START_COL + i
        c1 = ws.cell(row=1, column=col, value=d)
        c1.number_format = "M/D"
        c1.font = GANTT_DATE_FONT
        c1.alignment = Alignment(horizontal="center")
        c1.border = THIN_BORDER
        if d.weekday() == 5:
            c1.fill = SAT_FILL
        elif d.weekday() == 6:
            c1.fill = SUN_FILL
        ws.column_dimensions[get_column_letter(col)].width = 5.5

    # Row 2: ラベル + 曜日
    for i, (label, _) in enumerate(GANTT_COLS):
        cell = ws.cell(row=2, column=i + 1, value=label)
        cell.font = GANTT_LABEL_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER
        cell.fill = LABEL_FILL

    for i, d in enumerate(dates):
        col = DATE_START_COL + i
        c2 = ws.cell(row=2, column=col, value=WEEKDAYS_JP[d.weekday()])
        c2.font = GANTT_WD_FONT
        c2.alignment = Alignment(horizontal="center")
        c2.border = THIN_BORDER
        if d.weekday() == 5:
            c2.fill = SAT_FILL
        elif d.weekday() == 6:
            c2.fill = SUN_FILL

    # 空行（入力エリア）
    GANTT_NO_WRAP = Alignment(vertical="center", wrap_text=False)
    GANTT_ROW_HEIGHT = 22
    max_data_row = GANTT_EMPTY_ROWS + 2

    for r in range(3, max_data_row + 1):
        for c in range(1, GANTT_LEFT_COLS + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            if c in (7, 8):
                cell.number_format = "M/D"
                cell.alignment = CENTER_ALIGN
        for col in range(DATE_START_COL, DATE_START_COL + gantt_days):
            ws.cell(row=r, column=col).border = THIN_BORDER
        ws.row_dimensions[r].height = GANTT_ROW_HEIGHT

    # 条件付き書式: 開始〜終了のバー
    last_col = get_column_letter(DATE_START_COL + gantt_days - 1)
    fmt_range = f"I3:{last_col}{max_data_row}"
    ws.conditional_formatting.add(fmt_range, FormulaRule(
        formula=['AND(I$1>=$G3,I$1<=$H3)'],
        fill=BAR_FILL,
    ))

    ws.freeze_panes = "I3"

    # ==================== 日次入力 ====================
    ws_d = wb["日次入力"]
    d_headers = ["日付", "客先", "工事件名", "昼/夜", "工事内容", "重点工事"]
    d_widths = [12, 12, 28, 8, 40, 10]
    for c, (h, w) in enumerate(zip(d_headers, d_widths), 1):
        cell = ws_d.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER
        ws_d.column_dimensions[get_column_letter(c)].width = w

    # プルダウン
    dv_dn = DataValidation(type="list", formula1='"昼,夜,なし"', allow_blank=True, showDropDown=False)
    ws_d.add_data_validation(dv_dn)
    dv_dn.add("D2:D500")

    dv_pri = DataValidation(type="list", formula1='"有,無"', allow_blank=True, showDropDown=False)
    ws_d.add_data_validation(dv_pri)
    dv_pri.add("F2:F500")

    ws_d.freeze_panes = "A2"

    # ガントチャートをアクティブに
    wb.active = 0

    # 月サブフォルダに保存
    out_dir = os.path.join(config.DATA_DIR, month_str)
    os.makedirs(out_dir, exist_ok=True)
    xlsm_path = os.path.join(out_dir, f"入力_{person_name}.xlsm")
    wb.save(xlsm_path)
    print(f"  作成: {month_str}/入力_{person_name}.xlsm (ガント: {gantt_days}日間)")
    return xlsm_path


def main():
    parser = argparse.ArgumentParser(description="工事予定表 入力ファイル生成")
    parser.add_argument("--month", default=None,
                        help="対象月 YYMM形式 (例: 2604=2026年4月、デフォルト: 当月)")
    parser.add_argument("--persons", default="担当者A,担当者B,担当者C",
                        help="担当者名 (カンマ区切り)")
    parser.add_argument("--empty", action="store_true",
                        help="空ファイル生成 (サンプルデータなし)")
    args = parser.parse_args()

    if args.month:
        month_str = args.month
    else:
        today = datetime.date.today()
        month_str = f"{today.year % 100:02d}{today.month:02d}"

    year, month = parse_month(month_str)
    persons = [p.strip() for p in args.persons.split(",")]

    os.makedirs(config.DATA_DIR, exist_ok=True)
    print(f"=== 入力システム生成 ({year}年{month}月) ===\n")

    for person in persons:
        create_input_file(person, month_str)

    print(f"\n完了! {config.DATA_DIR}/{month_str}/")
    print(f"  VBA自動同期: ガントチャート保存時に日次入力シートが自動更新")


if __name__ == "__main__":
    main()
