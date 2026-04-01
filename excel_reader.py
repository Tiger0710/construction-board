"""工事予定表Excelファイル読み込みモジュール"""
import os
import datetime
import traceback

import openpyxl

import config


def _to_date_str(val):
    """date/datetime/文字列 → 'YYYY-MM-DD' 文字列に変換"""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date().isoformat()
    if isinstance(val, datetime.date):
        return val.isoformat()
    return str(val).strip()


def _cell_str(val):
    """セル値を文字列に変換（None → 空文字）"""
    if val is None:
        return ""
    return str(val).strip()


def load_construction_data():
    """工事予定表Excelを読み込み、辞書で返す

    Excel列: A=日付, B=客先, C=工事件名, D=弊社担当者,
             E=安全品質管理部担当者, F=協力会社名, G=協力会社担当者,
             H=当日の作業内容, I=作業時間, J=重点作業有無
    """
    today = datetime.date.today()
    today_str = today.isoformat()
    tomorrow_str = (today + datetime.timedelta(days=1)).isoformat()

    # ファイル存在チェック
    if not os.path.exists(config.EXCEL_PATH):
        return {
            "items": [],
            "updated_at": datetime.datetime.now().isoformat(),
            "total": 0,
            "today": today_str,
            "tomorrow": tomorrow_str,
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

            priority_raw = _cell_str(row[9]) if len(row) > 9 else ""
            is_priority = priority_raw == "有"

            item = {
                "date": date_str,
                "client": _cell_str(row[1]),
                "title": _cell_str(row[2]),
                "our_person": _cell_str(row[3]),
                "safety_person": _cell_str(row[4]),
                "partner": _cell_str(row[5]),
                "partner_person": _cell_str(row[6]),
                "work_content": _cell_str(row[7]),
                "work_time": _cell_str(row[8]),
                "priority": priority_raw,
                "is_priority": is_priority,
            }
            items.append(item)

        wb.close()

        # 日付順 → 重点作業を先頭 → 工事件名順
        items.sort(key=lambda x: (x["date"], 0 if x["is_priority"] else 1, x["title"]))

        return {
            "items": items,
            "updated_at": datetime.datetime.now().isoformat(),
            "total": len(items),
            "today": today_str,
            "tomorrow": tomorrow_str,
            "error": None,
        }

    except Exception as e:
        return {
            "items": [],
            "updated_at": datetime.datetime.now().isoformat(),
            "total": 0,
            "today": today_str,
            "tomorrow": tomorrow_str,
            "error": f"読み込みエラー: {e}\n{traceback.format_exc()}",
        }
