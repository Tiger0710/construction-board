"""DirectCloudダウンロードスクリプト

GitHub Actions から呼び出し、DirectCloud の工事予定表フォルダから
入力_*.xlsm を downloads/{YYMM}/ にダウンロードする。

API構成:
  - 認証:       POST /openapi/jauth/token (sync_directcloud.pyと共有)
  - フォルダ一覧: GET  /openapi/v2/folders/lists?node=XXX
  - ファイル一覧: GET  /openapp/v1/files/index/{node}
  - ダウンロード: POST /openapp/v1/files/download/{node}

必要な環境変数:
  DIRECTCLOUD_SERVICE      - API Service
  DIRECTCLOUD_SERVICE_KEY  - API Service Key
  DIRECTCLOUD_CODE         - 会社コード
  DIRECTCLOUD_ID           - ユーザーID
  DIRECTCLOUD_PASSWORD     - ユーザーパスワード
  DIRECTCLOUD_NODE         - 工事予定表フォルダのnode値
"""

import json
import os
import sys
import urllib.parse
import urllib.request

from merge_inputs import get_target_months
from sync_directcloud import API_BASE, _encode_node, get_token, list_folders

DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")


def list_files(token, folder_node):
    """フォルダ内のファイル一覧を取得"""
    encoded = _encode_node(folder_node)
    url = f"{API_BASE}/openapp/v1/files/index/{encoded}"
    req = urllib.request.Request(url)
    req.add_header("access_token", token)
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    # レスポンス構造をログ出力（API仕様が限定的なため）
    if not result.get("success"):
        print(f"  ファイル一覧エラー: {json.dumps(result, ensure_ascii=False, default=str)[:500]}")
        return []

    # files キーまたは data.files を試行
    files = result.get("files") or result.get("data", {}).get("files", [])
    return files


def get_download_url(token, folder_node, file_seq):
    """ファイルのダウンロードURLを取得"""
    encoded = _encode_node(folder_node)
    url = f"{API_BASE}/openapp/v1/files/download/{encoded}"

    data = urllib.parse.urlencode({"file_seq": file_seq}).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("access_token", token)
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    if not result.get("success"):
        print(f"  ダウンロードURL取得エラー: {json.dumps(result, ensure_ascii=False, default=str)[:500]}")
        return None

    # download_url または url キーを試行
    return result.get("download_url") or result.get("url")


def download_file(url, dest_path):
    """URLからファイルをダウンロード"""
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req) as resp:
        data = resp.read()

    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    with open(dest_path, "wb") as f:
        f.write(data)

    size_kb = len(data) / 1024
    print(f"  ダウンロード完了: {os.path.basename(dest_path)} ({size_kb:.0f}KB)")


def main():
    base_node = os.environ.get("DIRECTCLOUD_NODE")
    if not base_node:
        print("DIRECTCLOUD_NODE が未設定")
        return 1

    token = get_token()
    target_months = get_target_months()
    print(f"対象月: {target_months}")

    # 工事予定表フォルダ配下のサブフォルダ一覧
    folders = list_folders(token, base_node)
    month_folders = {}
    for f in folders:
        name = f.get("name", "")
        if name in target_months:
            month_folders[name] = f["node"]

    if not month_folders:
        print(f"対象月のフォルダなし (検出フォルダ: {[f.get('name') for f in folders]})")
        return 0

    total = 0
    for yymm, folder_node in sorted(month_folders.items()):
        print(f"\n[{yymm}]")
        files = list_files(token, folder_node)

        if not files:
            print("  ファイルなし")
            continue

        # 初回はレスポンス構造をログ出力
        if files:
            sample = files[0]
            print(f"  APIレスポンスサンプル: {json.dumps(sample, ensure_ascii=False, default=str)[:300]}")

        for file_info in files:
            name = file_info.get("name", "") or file_info.get("file_name", "")
            file_seq = file_info.get("file_seq") or file_info.get("seq")

            if not name.startswith("\u5165\u529b_"):  # 入力_
                continue
            if not (name.endswith(".xlsm") or name.endswith(".xlsx")):
                continue
            if name.startswith("~$"):
                continue

            print(f"  取得中: {name} (file_seq={file_seq})")
            dl_url = get_download_url(token, folder_node, file_seq)
            if not dl_url:
                print(f"  スキップ: {name} (ダウンロードURL取得失敗)")
                continue

            dest = os.path.join(DOWNLOAD_DIR, yymm, name)
            try:
                download_file(dl_url, dest)
                total += 1
            except Exception as e:
                print(f"  ダウンロードエラー: {name} -> {e}")

    print(f"\n完了: {total}ファイルをダウンロード -> {DOWNLOAD_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
