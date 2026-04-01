"""CI用パイプライン: merge → HTML生成 → deploy/ に配置"""
import os
import sys
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import config
from merge_inputs import merge_all, write_output
from generate_html import generate_html, read_css
from excel_reader import load_construction_data


def main():
    print("=== CI Pipeline ===\n")

    # 1. Merge input files
    print("[1/3] 入力ファイル統合...")
    merged = merge_all()
    if not merged:
        print("  データなし → 空のHTMLを生成します")
    else:
        write_output(merged)
        print(f"  -> {len(merged)}件 統合完了")

    # 2. Generate HTML
    print("\n[2/3] HTML生成...")
    data = load_construction_data()
    css = read_css()
    html = generate_html(data, css)

    os.makedirs(config.DEPLOY_DIR, exist_ok=True)
    html_path = os.path.join(config.DEPLOY_DIR, "index.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  -> {html_path} ({len(html):,} bytes)")

    # 3. Copy upload page to deploy/
    print("\n[3/3] 静的ファイルコピー...")
    upload_src = os.path.join(BASE_DIR, "static", "upload.html")
    if os.path.exists(upload_src):
        shutil.copy2(upload_src, os.path.join(config.DEPLOY_DIR, "upload.html"))
        print("  -> upload.html")

    item_count = len(data.get("items", []))
    print(f"\n完了! {item_count}件のデータ → deploy/ に配置済み")
    return 0


if __name__ == "__main__":
    sys.exit(main())
