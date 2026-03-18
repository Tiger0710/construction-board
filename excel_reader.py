"""工事予定表Excelファイル読み込みモジュール"""
import os
import datetime
import traceback

import openpyxl

import config

# ステータスマップ（日本語 → コード）
STATUS_MAP = {
    "予定": "scheduled",
    "準備中": "preparing",
    "進行中": "in_progress",
    "完了": "completed",
    "遅延": "delayed",
    "中止": "cancelled",
}

# ソート優先度（小さいほど上に表示）
STATUS_PRIORITY = {
    "in_progress": 0,
    "preparing": 1,
    "scheduled": 1,
    "delayed": 2,
    "completed": 3,
    "cancelled": 4,
}


def _to_date_str(val):
    """date/datetime/文字列 → 'YYYY-MM-DD' 文字列に変換"""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date().isoformat()
    if isinstance(val, datetime.date):
        return val.isoformat()
    return str(val).strip()


def _to_time_str(val):
    """time/datetime/文字列 → 'HH:MM' 文字列に変換"""
    if val is None:
        return ""
    if isinstance(val, datetime.datetime):
        return val.strftime("%H:%M")
    if isinstance(val, datetime.time):
        return val.strftime("%H:%M")
    return str(val).strip()


def _sort_key(item, today_str):
    """ソートキーを生成: 当日進行中 → 当日予定/準備中 → 未来 → 完了 → 中止"""
    code = item.get("status_code", "scheduled")
    priority = STATUS_PRIORITY.get(code, 5)
    is_today = item.get("date") == today_str

    if is_today and code == "in_progress":
        group = 0
    elif is_today and code in ("scheduled", "preparing"):
        group = 1
    elif code in ("scheduled", "preparing", "in_progress", "delayed"):
        group = 2
    elif code == "completed":
        group = 3
    else:  # cancelled
        group = 4

    return (group, item.get("date", ""), item.get("start_time", ""))


def load_construction_data():
    """工事予定表Excelを読み込み、辞書で返す"""
    today_str = datetime.date.today().isoformat()

    # ファイル存在チェック
    if not os.path.exists(config.EXCEL_PATH):
        return {
            "items": [],
            "updated_at": datetime.datetime.now().isoformat(),
            "total": 0,
            "today": today_str,
            "error": f"ファイルが見つかりません: {config.EXCEL_PATH}",
        }

    # ロックファイルチェック
    lock_file = os.path.join(
        os.path.dirname(config.EXCEL_PATH),
        "~$" + os.path.basename(config.EXCEL_PATH),
    )
    if os.path.exists(lock_file):
        # ロックファイルがあっても読み取り専用なので読める場合が多い
        pass

    try:
        wb = openpyxl.load_workbook(
            config.EXCEL_PATH, read_only=True, data_only=True
        )
        ws = wb["工事予定"]

        items = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            # A列(日付)がNoneの行はスキップ
            if row[0] is None:
                continue

            date_str = _to_date_str(row[0])
            if date_str is None:
                continue

            status_raw = str(row[6]).strip() if row[6] else "予定"
            status_code = STATUS_MAP.get(status_raw, "scheduled")

            progress_val = row[7] if row[7] is not None else 0
            try:
                progress_val = int(progress_val)
            except (ValueError, TypeError):
                progress_val = 0

            item = {
                "date": date_str,
                "start_time": _to_time_str(row[1]),
                "end_time": _to_time_str(row[2]),
                "name": str(row[3]).strip() if row[3] else "",
                "location": str(row[4]).strip() if row[4] else "",
                "person": str(row[5]).strip() if row[5] else "",
                "status": status_raw,
                "status_code": status_code,
                "progress": progress_val,
                "note": str(row[8]).strip() if len(row) > 8 and row[8] else "",
            }
            items.append(item)

        wb.close()

        # ソート
        items.sort(key=lambda x: _sort_key(x, today_str))

        return {
            "items": items,
            "updated_at": datetime.datetime.now().isoformat(),
            "total": len(items),
            "today": today_str,
            "error": None,
        }

    except Exception as e:
        return {
            "items": [],
            "updated_at": datetime.datetime.now().isoformat(),
            "total": 0,
            "today": today_str,
            "error": f"読み込みエラー: {e}\n{traceback.format_exc()}",
        }
