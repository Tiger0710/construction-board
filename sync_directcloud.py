"""DirectCloud同期スクリプト

GitHub Actions から呼び出し、data/{YYMM}/入力_*.xlsm を
DirectCloud の工事予定表フォルダにアップロードする。

API構成:
  - 認証:       POST /openapi/jauth/token (v1/v2共通)
  - フォルダ一覧: GET  /openapi/v2/folders/lists?node=XXX (v2)
  - フォルダ作成: POST /openapp/v1/folders/create/{node}  (v1, URL encode必須)
  - アップロード: POST /openapp/v1/files/upload/{node}     (v1, URL encode必須)

必要な環境変数:
  DIRECTCLOUD_SERVICE      - API Service
  DIRECTCLOUD_SERVICE_KEY  - API Service Key
  DIRECTCLOUD_CODE         - 会社コード
  DIRECTCLOUD_ID           - ユーザーID
  DIRECTCLOUD_PASSWORD     - ユーザーパスワード
  DIRECTCLOUD_NODE         - 工事予定表フォルダのnode値
"""

import glob
import os
import sys
import urllib.request
import urllib.parse
import json

API_BASE = "https://api.directcloud.jp"


def get_token():
    """認証トークンを取得"""
    url = f"{API_BASE}/openapi/jauth/token"
    data = urllib.parse.urlencode({
        "service": os.environ["DIRECTCLOUD_SERVICE"],
        "service_key": os.environ["DIRECTCLOUD_SERVICE_KEY"],
        "code": os.environ["DIRECTCLOUD_CODE"],
        "id": os.environ["DIRECTCLOUD_ID"],
        "password": os.environ["DIRECTCLOUD_PASSWORD"],
    }).encode()

    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    if not result.get("success"):
        print(f"認証エラー: {result}")
        sys.exit(1)

    print("DirectCloud認証OK")
    return result["access_token"]


def _encode_node(node):
    """nodeをURLパス用にエンコード（{} 等の特殊文字対応）"""
    return urllib.parse.quote(node, safe="")


def list_folders(token, parent_node):
    """v2 APIでフォルダ一覧を取得"""
    encoded = _encode_node(parent_node)
    url = f"{API_BASE}/openapi/v2/folders/lists?node={encoded}&limit=100"
    req = urllib.request.Request(url)
    req.add_header("access_token", token)
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    if result.get("result") != "success":
        return []
    return result.get("data", {}).get("folders", [])


def create_folder(token, parent_node, name):
    """v1 APIでサブフォルダを作成"""
    encoded = _encode_node(parent_node)
    url = f"{API_BASE}/openapp/v1/folders/create/{encoded}"
    boundary = "----FormBoundary7MA4YWxkTrZu0gW"
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="name"\r\n\r\n'
        f"{name}\r\n"
        f"--{boundary}--\r\n"
    ).encode()

    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("access_token", token)
    req.add_header("Content-Type", f"multipart/form-data; boundary={boundary}")
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    if not result.get("success"):
        print(f"  フォルダ作成エラー: {result}")
        return None

    print(f"  フォルダ作成: {name} (node={result.get('node')})")
    return result.get("node")


def find_or_create_folder(token, parent_node, name):
    """既存フォルダを探すか、なければ作成"""
    folders = list_folders(token, parent_node)
    for f in folders:
        if f.get("name") == name:
            print(f"  既存フォルダ: {name} (node={f['node']})")
            return f["node"]
    return create_folder(token, parent_node, name)


def upload_file(token, node, filepath):
    """v1 APIでファイルをアップロード"""
    encoded = _encode_node(node)
    url = f"{API_BASE}/openapp/v1/files/upload/{encoded}"
    filename = os.path.basename(filepath)

    boundary = "----FormBoundary7MA4YWxkTrZu0gW"
    with open(filepath, "rb") as f:
        file_data = f.read()

    body = bytearray()
    body.extend(f"--{boundary}\r\n".encode())
    body.extend(
        f'Content-Disposition: form-data; name="Filedata"; filename="{filename}"\r\n'.encode()
    )
    body.extend(b"Content-Type: application/octet-stream\r\n\r\n")
    body.extend(file_data)
    body.extend(f"\r\n--{boundary}--\r\n".encode())

    req = urllib.request.Request(url, data=bytes(body), method="POST")
    req.add_header("access_token", token)
    req.add_header("Content-Type", f"multipart/form-data; boundary={boundary}")

    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    if result.get("success"):
        print(f"  アップロード完了: {filename}")
    else:
        print(f"  アップロードエラー: {filename} -> {result}")
    return result.get("success", False)


def main():
    base_node = os.environ.get("DIRECTCLOUD_NODE")
    if not base_node:
        print("DIRECTCLOUD_NODE が未設定、DirectCloud同期スキップ")
        return 0

    # data/ 配下の月サブフォルダを検出
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    month_dirs = sorted(glob.glob(os.path.join(data_dir, "[0-9][0-9][0-9][0-9]")))

    if not month_dirs:
        print("アップロード対象の月フォルダなし")
        return 0

    token = get_token()

    total = 0
    for month_dir in month_dirs:
        month = os.path.basename(month_dir)
        files = glob.glob(os.path.join(month_dir, "入力_*.xlsm")) + \
                glob.glob(os.path.join(month_dir, "入力_*.xlsx"))

        if not files:
            continue

        print(f"\n[{month}] {len(files)}ファイル")
        month_node = find_or_create_folder(token, base_node, month)
        if not month_node:
            print(f"  フォルダ取得失敗、スキップ")
            continue

        for f in files:
            if upload_file(token, month_node, f):
                total += 1

    print(f"\n完了: {total}ファイルをDirectCloudに同期")
    return 0


if __name__ == "__main__":
    sys.exit(main())
