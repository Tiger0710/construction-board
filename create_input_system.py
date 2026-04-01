"""入力システム生成: ガントチャート + 日次入力 + VBA自動同期

使い方:
  python create_input_system.py                       # 当月、サンプルデータ付き
  python create_input_system.py --month 2604 --empty  # 4月分、空データ
  python create_input_system.py --persons 田中,山本,鈴木 --month 2604 --empty
"""
import os
import sys
import argparse
import calendar
import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

import config

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

# ガントチャート列定義: (列名, 幅)
GANTT_COLS = [
    ("客先", 12), ("工事件名", 38), ("現場担当者", 10), ("安品担当者", 10),
    ("協力会社名", 16), ("協力会社担当者", 10), ("開始", 10), ("終了", 10),
]
GANTT_LEFT_COLS = len(GANTT_COLS)  # 8列 (A-H)
DATE_START_COL = GANTT_LEFT_COLS + 1  # I列から日付
GANTT_EMPTY_ROWS = 30  # 空ファイルの入力可能行数


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


def create_input_file(person_name, month_str, projects=None, daily_data=None):
    """入力ファイル生成

    Args:
        person_name: 担当者名
        month_str: 'YYMM' (例: '2604')
        projects: 案件リスト。Noneで空ファイル
        daily_data: 日次入力サンプル。Noneで空
    """
    year, month = parse_month(month_str)
    dates = get_month_dates(year, month)
    gantt_days = len(dates)

    wb = Workbook()

    # ==================== ガントチャート ====================
    ws = wb.active
    ws.title = "ガントチャート"

    # 列幅
    for i, (_, width) in enumerate(GANTT_COLS):
        ws.column_dimensions[get_column_letter(i + 1)].width = width

    # Row 1: ボタン配置エリア + 日付
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

    # データ行
    GANTT_NO_WRAP = Alignment(vertical="center", wrap_text=False)
    GANTT_ROW_HEIGHT = 22

    data_rows = len(projects) if projects else 0
    max_data_row = max(data_rows, GANTT_EMPTY_ROWS) + 2

    if projects:
        first_date = dates[0]
        for idx, proj in enumerate(projects):
            row = idx + 3
            client, title, our, safety, partner, pp, s_off, e_off = proj
            start_d = first_date + datetime.timedelta(days=s_off)
            end_d = first_date + datetime.timedelta(days=e_off)
            vals = [client, title, our, safety, partner, pp, start_d, end_d]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=c, value=v)
                cell.font = CELL_FONT
                cell.border = THIN_BORDER
                if c in (3, 4, 6):
                    cell.alignment = CENTER_ALIGN
                elif c in (7, 8):
                    cell.number_format = "M/D"
                    cell.alignment = CENTER_ALIGN
                else:
                    cell.alignment = GANTT_NO_WRAP
            for col in range(DATE_START_COL, DATE_START_COL + gantt_days):
                ws.cell(row=row, column=col).border = THIN_BORDER
            ws.row_dimensions[row].height = GANTT_ROW_HEIGHT

    # 空行（入力エリア）
    empty_start = (data_rows + 3) if projects else 3
    for r in range(empty_start, max_data_row + 1):
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
    ws_d = wb.create_sheet("日次入力")
    d_headers = ["日付", "客先", "工事件名", "昼/夜", "工事内容", "重点工事"]
    d_widths = [12, 12, 28, 8, 40, 10]
    for c, (h, w) in enumerate(zip(d_headers, d_widths), 1):
        ws_d.cell(row=1, column=c, value=h)
        ws_d.column_dimensions[get_column_letter(c)].width = w
    apply_header(ws_d, 1, 1, len(d_headers))

    # サンプルデータ展開
    if projects and daily_data:
        first_date = dates[0]
        entries = []
        for idx, proj in enumerate(projects):
            client, title, _, _, _, _, s_off, e_off = proj
            for day in range(s_off, e_off + 1):
                d = first_date + datetime.timedelta(days=day)
                dn, work, pri = daily_data.get((day, idx), ("", "", ""))
                entries.append((d, client, title, dn, work, pri))
        entries.sort(key=lambda x: (x[0], x[1], x[2]))

        for i, (d, client, title, dn, work, pri) in enumerate(entries, 2):
            ws_d.cell(row=i, column=1, value=d).number_format = "M/D"
            ws_d.cell(row=i, column=2, value=client)
            ws_d.cell(row=i, column=3, value=title)
            ws_d.cell(row=i, column=4, value=dn)
            ws_d.cell(row=i, column=5, value=work)
            ws_d.cell(row=i, column=6, value=pri)
            for c in range(1, 7):
                cell = ws_d.cell(row=i, column=c)
                cell.font = CELL_FONT
                cell.border = THIN_BORDER
                cell.alignment = CENTER_ALIGN if c in (1, 4, 6) else CELL_ALIGN

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

    xlsx_path = os.path.join(config.DATA_DIR, f"入力_{person_name}_{month_str}.xlsx")
    wb.save(xlsx_path)
    print(f"  作成: {os.path.basename(xlsx_path)} (ガント: {gantt_days}日間)")
    return xlsx_path


def main():
    parser = argparse.ArgumentParser(description="工事予定表 入力ファイル生成")
    parser.add_argument("--month", default=None,
                        help="対象月 YYMM形式 (例: 2604=2026年4月、デフォルト: 当月)")
    parser.add_argument("--persons", default="担当者A,担当者B,担当者C",
                        help="担当者名 (カンマ区切り)")
    parser.add_argument("--empty", action="store_true",
                        help="空ファイル生成 (サンプルデータなし)")
    args = parser.parse_args()

    # 月の決定
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

    print(f"\n完了! {config.DATA_DIR}")
    print(f"  ファイル名パターン: 入力_担当者名_{month_str}.xlsx")
    print("  ※ 担当者名は自由に変更可能")
    print("  ※ ガントチャートを記入→保存するだけで自動反映（VBA不要）")


if __name__ == "__main__":
    main()
