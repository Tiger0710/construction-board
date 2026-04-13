"""CI用パイプライン: static/ → deploy/ にコピー"""
import os
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

import config


def main():
    print("=== CI Pipeline ===\n")

    # static/ → deploy/ をコピー
    print("[1/1] static/ → deploy/ コピー...")
    src = os.path.join(BASE_DIR, "static")
    dst = config.DEPLOY_DIR

    if os.path.exists(dst):
        shutil.rmtree(dst)
    shutil.copytree(src, dst)

    count = sum(len(files) for _, _, files in os.walk(dst))
    print(f"  -> {count} ファイルをコピーしました")
    print(f"\n完了! deploy/ に配置済み")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
