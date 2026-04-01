"""DirectCloud監視 → 自動統合・HTML生成 → Netlifyデプロイ

タスクスケジューラから自動起動される常駐プロセス。
DirectCloud上の入力ファイルを監視し、変更検出時にパイプラインを実行。
"""
import os
import sys
import glob
import time
import shutil
import datetime
import subprocess
import traceback
import logging
from logging.handlers import RotatingFileHandler

import config
from merge_inputs import merge_all, write_output
from generate_html import generate_html, read_css
from excel_reader import load_construction_data

POLL_INTERVAL = 15   # 秒ごとにチェック
DEBOUNCE_SEC = 5     # 変更検出後、追加変更を待つ秒数
MAX_ERRORS = 5       # 連続エラー上限 (超えたら長めにwait)
ERROR_COOLDOWN = 120  # エラー連続時のクールダウン(秒)
GIT_SYNC_INTERVAL = 8  # N回ポーリングごとにgit pull (8 * 15秒 = 2分)

# ログ設定
LOG_DIR = os.path.join(config.BASE_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logger = logging.getLogger("watcher")
logger.setLevel(logging.INFO)

fh = RotatingFileHandler(
    os.path.join(LOG_DIR, "watcher.log"),
    maxBytes=1024 * 1024,  # 1MB
    backupCount=7,
    encoding="utf-8",
)
fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(fh)

ch = logging.StreamHandler()
ch.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(ch)


def get_input_files():
    """入力ファイル一覧 (サブフォルダ + 旧フラット、ロックファイル除外)"""
    files = []
    for d in (config.INPUT_DIR, config.DATA_DIR):
        if not os.path.isdir(d):
            continue
        # 新形式: {base}/{YYMM}/入力_*.xlsm
        for ext in ("xlsm", "xlsx"):
            files.extend(glob.glob(os.path.join(d, "*", f"入力_*.{ext}")))
        # 旧形式: {base}/入力_*.xlsx
        for ext in ("xlsm", "xlsx"):
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

    deploy_html = os.path.join(config.DEPLOY_DIR, "index.html")
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()
    with open(deploy_html, "w", encoding="utf-8") as f:
        f.write(content)

    # 静的ページもコピー
    for page in ("upload.html", "sample.html", "admin.html"):
        src = os.path.join(config.BASE_DIR, "static", page)
        if os.path.exists(src):
            import shutil
            shutil.copy2(src, os.path.join(config.DEPLOY_DIR, page))

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
            logger.info("Netlifyデプロイ完了")
        else:
            logger.warning(f"Netlifyデプロイ失敗: {result.stderr[:200]}")
    except FileNotFoundError:
        logger.warning("netlify CLI未インストール → デプロイスキップ")
    except subprocess.TimeoutExpired:
        logger.warning("Netlifyデプロイタイムアウト")


def sync_repo_to_directcloud():
    """git pull → 新ファイルをDirectCloudにコピー"""
    if config.INPUT_DIR == config.DATA_DIR:
        return  # DirectCloud未接続時はスキップ

    try:
        result = subprocess.run(
            ["git", "pull", "--ff-only"],
            capture_output=True, text=True, timeout=30,
            cwd=config.BASE_DIR,
        )
        if result.returncode != 0:
            logger.debug(f"git pull失敗: {result.stderr[:100]}")
            return
        if "Already up to date" in result.stdout:
            return

        logger.info(f"git pull: {result.stdout.strip()}")

        # data/{YYMM}/入力_*.xlsm → DirectCloud/{YYMM}/ にコピー
        copied = 0
        for month_dir in glob.glob(os.path.join(config.DATA_DIR, "*")):
            if not os.path.isdir(month_dir):
                continue
            yymm = os.path.basename(month_dir)
            if not yymm.isdigit() or len(yymm) != 4:
                continue

            dc_dir = os.path.join(config.INPUT_DIR, yymm)
            os.makedirs(dc_dir, exist_ok=True)

            for src in glob.glob(os.path.join(month_dir, "入力_*.xlsm")):
                dest = os.path.join(dc_dir, os.path.basename(src))
                if not os.path.exists(dest):
                    shutil.copy2(src, dest)
                    copied += 1
                    logger.info(f"DirectCloudへコピー: {yymm}/{os.path.basename(src)}")

        if copied:
            logger.info(f"DirectCloud同期: {copied}ファイルをコピー")

    except subprocess.TimeoutExpired:
        logger.warning("git pullタイムアウト")
    except Exception as e:
        logger.debug(f"git sync エラー: {e}")


def run_pipeline():
    """merge → HTML生成 → Netlifyデプロイ"""
    logger.info("変更検出 → パイプライン実行")

    # 1. merge
    logger.info("統合開始...")
    merged = merge_all()
    if merged:
        write_output(merged)
    else:
        write_output([])
        logger.info("データなし → 空のHTMLを生成")

    # 2. generate HTML
    logger.info("HTML生成...")
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
    logger.info("Netlifyデプロイ...")
    deploy_to_netlify(html_path)

    item_count = len(data.get("items", []))
    logger.info(f"完了! {item_count}件 → {config.NETLIFY_URL}")


def main():
    logger.info("=" * 50)
    logger.info("工事予定表 自動更新モニター 起動")
    logger.info(f"入力: {config.INPUT_DIR}")
    logger.info(f"チェック間隔: {POLL_INTERVAL}秒")
    logger.info(f"公開: {config.NETLIFY_URL}")
    logger.info("=" * 50)

    print("=" * 55)
    print("  工事予定表 自動更新モニター")
    print(f"  入力: {config.INPUT_DIR}")
    print(f"  チェック間隔: {POLL_INTERVAL}秒")
    print(f"  公開: {config.NETLIFY_URL}")
    print(f"  ログ: {LOG_DIR}")
    print("  停止: Ctrl+C")
    print("=" * 55)

    # 初回実行
    try:
        run_pipeline()
    except Exception as e:
        logger.error(f"初回実行エラー: {e}")
        traceback.print_exc()

    # 初回: repoからDirectCloudへ同期
    try:
        sync_repo_to_directcloud()
    except Exception as e:
        logger.debug(f"初回git sync: {e}")

    # 前回のmtimeを記録
    prev_mtimes = get_mtimes(get_input_files())
    error_count = 0
    poll_count = 0

    logger.info("監視開始...")
    print(f"\n監視中... (DirectCloudで保存すると自動デプロイ)")

    try:
        while True:
            time.sleep(POLL_INTERVAL)
            poll_count += 1

            try:
                # 定期的にgit pull → DirectCloudへ同期
                if poll_count % GIT_SYNC_INTERVAL == 0:
                    sync_repo_to_directcloud()

                current_files = get_input_files()
                current_mtimes = get_mtimes(current_files)

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
                    run_pipeline()
                    error_count = 0

                prev_mtimes = get_mtimes(get_input_files())

            except Exception as e:
                error_count += 1
                logger.error(f"エラー ({error_count}/{MAX_ERRORS}): {e}")
                if error_count >= MAX_ERRORS:
                    logger.warning(f"連続エラー → {ERROR_COOLDOWN}秒クールダウン")
                    time.sleep(ERROR_COOLDOWN)
                    error_count = 0

    except KeyboardInterrupt:
        logger.info("監視停止 (Ctrl+C)")
        print("\n監視を停止しました。")


if __name__ == "__main__":
    main()
