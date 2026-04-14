"""手島のExcel → JSON同期スクリプト (GitHub Actions用)

DirectCloudから入力_手島.xlsmをダウンロードし、
_YYMM_手島.json に変換してリポジトリに書き出す。
差分があればGitHub Actionsがcommit & pushする。

Usage:
    python sync_teshima.py
"""

import datetime
import json
import os
import sys

from download_directcloud import list_files, get_download_url, download_file
from sync_directcloud import get_token, list_folders
from migrate_to_json import read_input_file, clean_daily

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")


def get_target_months():
    """前月 + 当月 + 翌月"""
    today = datetime.date.today()
    current = f"{today.year % 100:02d}{today.month:02d}"
    next_month = today.replace(day=28) + datetime.timedelta(days=4)
    next_m = f"{next_month.year % 100:02d}{next_month.month:02d}"
    prev_month = today.replace(day=1) - datetime.timedelta(days=1)
    prev_m = f"{prev_month.year % 100:02d}{prev_month.month:02d}"
    return [prev_m, current, next_m]


def main():
    base_node = os.environ.get("DIRECTCLOUD_NODE")
    if not base_node:
        print("DIRECTCLOUD_NODE が未設定")
        return 1

    token = get_token()
    target_months = get_target_months()
    print(f"対象月: {target_months}")

    # フォルダ一覧
    folders = list_folders(token, base_node)
    month_folders = {}
    for f in folders:
        name = f.get("name", "")
        if name in target_months:
            month_folders[name] = f["node"]

    if not month_folders:
        print(f"対象月のフォルダなし (検出: {[f.get('name') for f in folders]})")
        return 0

    updated = 0
    for yymm, folder_node in sorted(month_folders.items()):
        print(f"\n=== {yymm} ===")
        files = list_files(token, folder_node)
        print(f"  ファイル数: {len(files)}")
        if files:
            for fi in files[:5]:
                fname = fi.get("name", "") or fi.get("file_name", "")
                print(f"    - {fname}")
        if not files:
            # デバッグ: 直接APIを叩いてレスポンスを確認
            import urllib.request as _ur
            from sync_directcloud import _encode_node, API_BASE
            _enc = _encode_node(folder_node)
            _req = _ur.Request(f"{API_BASE}/openapp/v1/files/index/{_enc}")
            _req.add_header("access_token", token)
            with _ur.urlopen(_req) as _resp:
                _raw = json.loads(_resp.read())
            print(f"  DEBUG raw keys: {list(_raw.keys())}")
            if "data" in _raw:
                print(f"  DEBUG data keys: {list(_raw['data'].keys()) if isinstance(_raw['data'], dict) else type(_raw['data'])}")
            print(f"  DEBUG raw preview: {json.dumps(_raw, ensure_ascii=False, default=str)[:500]}")
            continue

        # 入力_手島.xlsm を探す
        teshima_file = None
        for fi in files:
            name = fi.get("name", "") or fi.get("file_name", "")
            if name == "入力_手島.xlsm" or name == "入力_手島.xlsx":
                teshima_file = fi
                break

        if not teshima_file:
            print("  入力_手島 なし")
            continue

        file_seq = teshima_file.get("file_seq") or teshima_file.get("seq")
        fname = teshima_file.get("name", "") or teshima_file.get("file_name", "")
        print(f"  ダウンロード: {fname}")

        dl_url = get_download_url(token, folder_node, file_seq)
        if not dl_url:
            print("  ダウンロードURL取得失敗")
            continue

        dest = os.path.join(DOWNLOAD_DIR, yymm, fname)
        try:
            download_file(dl_url, dest)
        except Exception as e:
            print(f"  ダウンロードエラー: {e}")
            continue

        # Excel → JSON変換
        print("  JSON変換中...")
        try:
            projects, daily = read_input_file(dest)
        except Exception as e:
            print(f"  Excel読み取りエラー: {e}")
            continue

        seen = {}
        for p in projects:
            seen[p["id"]] = p
        result = {"projects": list(seen.values()), "daily": clean_daily(daily)}

        out_path = os.path.join(BASE_DIR, f"_{yymm}_手島.json")

        # 既存JSONと比較（差分がなければスキップ）
        new_json = json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True)
        if os.path.exists(out_path):
            with open(out_path, "r", encoding="utf-8") as f:
                old_json = json.dumps(json.load(f), ensure_ascii=False, indent=2, sort_keys=True)
            if old_json == new_json:
                print(f"  変更なし (案件{len(result['projects'])}件)")
                continue

        with open(out_path, "w", encoding="utf-8") as f:
            f.write(json.dumps(result, ensure_ascii=False, indent=2))

        print(f"  更新: 案件{len(result['projects'])}件, 日次{len(result['daily'])}件")
        updated += 1

    print(f"\n完了: {updated}ファイル更新")
    return 0


if __name__ == "__main__":
    sys.exit(main())
