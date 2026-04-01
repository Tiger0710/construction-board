"""工事予定表サイネージ設定"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

# 入力ファイル: DirectCloud or ローカル (CI時はDATA_DIRにフォールバック)
INPUT_DIR = os.environ.get("INPUT_DIR", r"E:\★工事共有\工事予定表")
if not os.path.isdir(INPUT_DIR):
    INPUT_DIR = DATA_DIR

EXCEL_PATH = os.path.join(DATA_DIR, "工事予定表.xlsx")
DEPLOY_DIR = os.path.join(BASE_DIR, "deploy")

# Netlify
NETLIFY_SITE_ID = "e7da59df-2ff8-4a01-bd77-7246c4ea9e02"
NETLIFY_URL = "https://hsj-construction-board.netlify.app"

REFRESH_INTERVAL_SEC = 300
ROWS_PER_PAGE = 10
PAGE_ROTATE_SEC = 15
PORT = 5555
