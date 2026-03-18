"""工事予定表サイネージ設定"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_PATH = os.path.join(DATA_DIR, "工事予定表.xlsx")

REFRESH_INTERVAL_SEC = 300  # 5分
ROWS_PER_PAGE = 8
PAGE_ROTATE_SEC = 15
PORT = 5555
