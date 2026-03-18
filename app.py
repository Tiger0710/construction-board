"""工事予定表サイネージ - Flaskサーバー"""
from flask import Flask, jsonify, render_template

import config
import excel_reader

app = Flask(__name__)


@app.route("/")
def index():
    return render_template(
        "index.html",
        refresh_interval=config.REFRESH_INTERVAL_SEC,
    )


@app.route("/api/data")
def api_data():
    data = excel_reader.load_construction_data()
    return jsonify(data)


@app.route("/api/config")
def api_config():
    return jsonify(
        {
            "refresh_interval": config.REFRESH_INTERVAL_SEC,
            "rows_per_page": config.ROWS_PER_PAGE,
            "page_rotate_sec": config.PAGE_ROTATE_SEC,
        }
    )


if __name__ == "__main__":
    print(f"工事予定表サイネージサーバー起動: http://0.0.0.0:{config.PORT}")
    app.run(host="0.0.0.0", port=config.PORT, debug=False)
