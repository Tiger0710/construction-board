"""入力ファイル監視 → 自動統合・HTML生成 → Netlifyデプロイ"""
import os
import sys
import glob
import time
import datetime
import subprocess
import traceback

import config
from merge_inputs import merge_all, write_output
from generate_html import generate_html, read_css
from excel_reader import load_construction_data

POLL_INTERVAL = 15  # 秒ごとにチェック
DEBOUNCE_SEC = 5    # 変更検出後、追加変更を待つ秒数


def get_input_files():
    """入力ファイル一覧 (DirectCloud + ローカル、ロックファイル除外)"""
    files = []
    for d in (config.INPUT_DIR, config.DATA_DIR):
        if not os.path.isdir(d):
            continue
        for ext in ("xlsx", "xlsm"):
            files.extend(glob.glob(os.path.join(d, f"入力_*.{ext}")))
    return [f for f in files if not os.path.basename(f).startswith("~$")]


def get_mtimes(files):
    """各ファイルの更新時刻を取得"""
    mtimes = {}
    for f in files:
        try:
            mtimes[f] = os.path.getmtime(f)
        except OSError:
            pass
    return mtimes


def deploy_to_netlify(html_path):
    """HTMLをNetlifyにデプロイ"""
    os.makedirs(config.DEPLOY_DIR, exist_ok=True)

    # HTMLをデプロイフォルダにコピー
    deploy_html = os.path.join(config.DEPLOY_DIR, "index.html")
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()
    with open(deploy_html, "w", encoding="utf-8") as f:
        f.write(content)

    netlify_cmd = os.path.join(os.path.expanduser("~"),
                               "AppData", "Roaming", "npm", "netlify.cmd")
    if not os.path.exists(netlify_cmd):
        netlify_cmd = "netlify"

    try:
        result = subprocess.run(
            [netlify_cmd, "deploy", "--prod",
             "--dir", config.DEPLOY_DIR,
             "--site", config.NETLIFY_SITE_ID],
            capture_output=True, text=True, timeout=60,
            cwd=config.BASE_DIR,
        )
        if result.returncode == 0:
            print("    Netlifyデプロイ完了!")
        else:
            print(f"    Netlifyデプロイ失敗: {result.stderr[:200]}")
    except FileNotFoundError:
        print("    netlify CLI未インストール → デプロイスキップ")
    except subprocess.TimeoutExpired:
        print("    Netlifyデプロイタイムアウト")


def run_pipeline():
    """merge → HTML生成 → Netlifyデプロイ"""
    now = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"\n[{now}] 変更検出 → パイプライン実行")

    # 1. merge
    print("  [1/3] 統合 → 工事予定表.xlsx...")
    merged = merge_all()
    if merged:
        write_output(merged)
    else:
        print("    データなし")
        return

    # 2. generate HTML
    print("  [2/3] HTML生成...")
    data = load_construction_data()
    css = read_css()
    html = generate_html(data, css)
    html_path = os.path.join(config.DEPLOY_DIR, "index.html")
    os.makedirs(config.DEPLOY_DIR, exist_ok=True)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    # デスクトップにもコピー
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "工事予定表.html")
    with open(desktop_path, "w", encoding="utf-8") as f:
        f.write(html)

    # 3. Netlifyデプロイ
    print("  [3/3] Netlifyデプロイ...")
    deploy_to_netlify(html_path)

    now2 = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"  [{now2}] 完了! {len(merged)}件")
    print(f"    URL: {config.NETLIFY_URL}")


def main():
    print("=" * 55)
    print("  工事予定表 自動更新モニター")
    print(f"  入力: {config.INPUT_DIR}")
    print(f"  チェック間隔: {POLL_INTERVAL}秒")
    print(f"  公開: {config.NETLIFY_URL}")
    print("  停止: Ctrl+C")
    print("=" * 55)

    # 初回実行
    run_pipeline()

    # 前回のmtimeを記録
    prev_mtimes = get_mtimes(get_input_files())

    print(f"\n監視中... (Excelで保存すると自動でデプロイ)")

    try:
        while True:
            time.sleep(POLL_INTERVAL)

            current_files = get_input_files()
            current_mtimes = get_mtimes(current_files)

            # 変更検出
            changed = False
            if set(current_mtimes.keys()) != set(prev_mtimes.keys()):
                changed = True
            else:
                for f, mtime in current_mtimes.items():
                    if prev_mtimes.get(f) != mtime:
                        changed = True
                        break

            if changed:
                time.sleep(DEBOUNCE_SEC)
                try:
                    run_pipeline()
                except Exception as e:
                    print(f"  エラー: {e}")
                    traceback.print_exc()
                prev_mtimes = get_mtimes(get_input_files())

    except KeyboardInterrupt:
        print("\n\n監視を停止しました。")


if __name__ == "__main__":
    main()
