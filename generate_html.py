"""工事予定表スタンドアロンHTML生成スクリプト

ExcelデータとCSS/JSを全て埋め込んだ単一HTMLファイルをデスクトップに生成する。
"""
import json
import os
import sys

# プロジェクトルートをパスに追加
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from excel_reader import load_construction_data

CSS_PATH = os.path.join(BASE_DIR, "static", "css", "style.css")
OUTPUT_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "工事予定表.html")


def read_css():
    """style.css を読み込む"""
    with open(CSS_PATH, "r", encoding="utf-8") as f:
        return f.read()


def generate_html(data: dict, css: str) -> str:
    """スタンドアロンHTMLを生成する"""
    data_json = json.dumps(data, ensure_ascii=False, indent=2)

    return f'''<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>工事予定表</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Noto+Sans+JP:wght@300;400;500;600;700&family=Noto+Serif+JP:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
{css}
  </style>
</head>
<body>

  <header class="board-header">
    <div class="header-left">
      <h1 class="header-title">工事予定表</h1>
      <span class="header-subtitle">Construction Schedule</span>
    </div>
    <div class="header-center">
      <div class="header-today" id="header-today"></div>
    </div>
    <div class="header-right">
      <div id="clock" class="header-clock">00:00:00</div>
      <div id="date" class="header-date"></div>
    </div>
  </header>

  <div class="board-table-header">
    <div class="col col-time">時間</div>
    <div class="col col-name">工事名</div>
    <div class="col col-location">場所</div>
    <div class="col col-status">ステータス</div>
    <div class="col col-person">担当者</div>
    <div class="col col-progress">進捗</div>
  </div>

  <div class="board-body" id="board-body"></div>

  <footer class="board-footer">
    <div class="footer-left">
      <span class="footer-live-dot"></span>
      最終更新 <span id="update-time" class="footer-value">--:--:--</span>
    </div>
    <div class="footer-center" id="page-indicator"></div>
    <div class="footer-right">
      全 <span id="total-count" class="footer-value">0</span> 件
    </div>
  </footer>

  <script>
    // ===== 埋め込みデータ =====
    const BOARD_DATA = {data_json};
    const ROWS_PER_PAGE = 8;
    const PAGE_ROTATE_SEC = 15;

    // ===== グローバル変数 =====
    let currentPage = 0;
    let totalPages = 1;
    let displayItems = [];
    let pageRotateTimer = null;
    let cycleComplete = false;     // 全ページ巡回完了フラグ
    let lastTimeCheckMinute = -1;  // 前回の時刻チェック分

    const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

    // ===== HTMLエスケープ =====
    function escapeHtml(str) {{
      if (str == null) return "";
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }}

    function escapeAttr(str) {{
      return escapeHtml(str);
    }}

    // ===== 時計機能 =====
    function startClock() {{
      function update() {{
        const now = new Date();
        const hh = String(now.getHours()).padStart(2, "0");
        const mm = String(now.getMinutes()).padStart(2, "0");
        const ss = String(now.getSeconds()).padStart(2, "0");
        document.getElementById("clock").textContent = hh + ":" + mm + ":" + ss;

        const y = now.getFullYear();
        const m = now.getMonth() + 1;
        const d = now.getDate();
        const w = WEEKDAYS[now.getDay()];
        document.getElementById("date").textContent = y + "年" + m + "月" + d + "日(" + w + ")";

        const todayEl = document.getElementById("header-today");
        if (todayEl) {{
          todayEl.textContent = m + "/" + d + " " + w + "曜日";
        }}
      }}
      update();
      setInterval(update, 1000);
    }}

    // ===== 「今日」フィルタリング =====
    function filterToday(items, todayStr) {{
      if (!items || items.length === 0) return [];
      const todayItems = items.filter(function(item) {{ return item.date === todayStr; }});
      if (todayItems.length > 0) return todayItems;
      return items;
    }}

    // ===== 現在時刻に最も近い工事のページを探す =====
    function findCurrentTimePage(items) {{
      if (!items || items.length === 0) return 0;

      const now = new Date();
      const currentHH = String(now.getHours()).padStart(2, "0");
      const currentMM = String(now.getMinutes()).padStart(2, "0");
      const currentTime = currentHH + ":" + currentMM;

      // 現在時刻の範囲内にある工事、または最も近い工事を探す
      let bestIndex = 0;
      let bestDistance = Infinity;

      for (let i = 0; i < items.length; i++) {{
        var item = items[i];
        var startTime = item.start_time || "";
        var endTime = item.end_time || "";

        // start_time <= currentTime <= end_time なら完全一致
        if (startTime && endTime && startTime <= currentTime && currentTime <= endTime) {{
          bestIndex = i;
          bestDistance = 0;
          break;
        }}

        // start_time との距離を計算（分単位）
        if (startTime) {{
          var dist = timeDistanceMinutes(currentTime, startTime);
          if (dist < bestDistance) {{
            bestDistance = dist;
            bestIndex = i;
          }}
        }}
      }}

      // そのアイテムが含まれるページ番号を返す
      return Math.floor(bestIndex / ROWS_PER_PAGE);
    }}

    // ===== 時刻間の距離（分）を計算 =====
    function timeDistanceMinutes(t1, t2) {{
      // t1, t2 は "HH:MM" 形式
      var parts1 = t1.split(":");
      var parts2 = t2.split(":");
      if (parts1.length < 2 || parts2.length < 2) return Infinity;
      var m1 = parseInt(parts1[0], 10) * 60 + parseInt(parts1[1], 10);
      var m2 = parseInt(parts2[0], 10) * 60 + parseInt(parts2[1], 10);
      return Math.abs(m1 - m2);
    }}

    // ===== 行の高さ計算 =====
    function calcRowHeight(rowCount) {{
      if (rowCount === 0) return 64;
      var available = window.innerHeight - 120 - 60 - 52;
      return Math.floor(available / Math.max(rowCount, 1));
    }}

    // ===== 行レンダリング =====
    function renderRow(item, index, rowHeight, changed, rowId) {{
      var evenOdd = index % 2 === 0 ? "row-even" : "row-odd";
      var flipClass = changed ? " flip-enter" : "";
      var progressComplete = item.progress >= 100 ? " complete" : "";
      var timeDisplay = item.end_time
        ? escapeHtml(item.start_time) + ' <span class="time-dash">&#x2015;</span> ' + escapeHtml(item.end_time)
        : escapeHtml(item.start_time);

      return '<div class="board-row ' + evenOdd + flipClass + '" data-id="' + escapeAttr(rowId) + '" style="height:' + rowHeight + 'px;">' +
        '<div class="col col-time">' + timeDisplay + '</div>' +
        '<div class="col col-name">' + escapeHtml(item.name) + '</div>' +
        '<div class="col col-location">' + escapeHtml(item.location) + '</div>' +
        '<div class="col col-status"><span class="status-badge status-' + item.status_code + '"><span class="status-dot"></span>' + escapeHtml(item.status) + '</span></div>' +
        '<div class="col col-person">' + escapeHtml(item.person) + '</div>' +
        '<div class="col col-progress">' +
          '<div class="progress-wrapper"><div class="progress-bar"><div class="progress-fill" style="width:' + item.progress + '%"></div></div></div>' +
          '<span class="progress-text' + progressComplete + '">' + item.progress + '%</span>' +
        '</div>' +
      '</div>';
    }}

    // ===== ボード表示 =====
    function renderBoard(data) {{
      var body = document.getElementById("board-body");

      // エラー表示
      if (data.error) {{
        body.innerHTML =
          '<div class="board-message">' +
            '<div class="board-message-icon">&#9888;</div>' +
            '<div class="board-message-error">' + escapeHtml(data.error) + '</div>' +
          '</div>';
        updateFooter(data, 0);
        stopPageRotation();
        return;
      }}

      // 「今日」のフィルタリング
      displayItems = filterToday(data.items, data.today);

      // 空の場合
      if (displayItems.length === 0) {{
        body.innerHTML =
          '<div class="board-message">' +
            '<div class="board-message-icon">&#128736;</div>' +
            '<div class="board-message-text">本日の工事予定はありません</div>' +
          '</div>';
        updateFooter(data, 0);
        stopPageRotation();
        return;
      }}

      // 時間順にソート（start_time昇順）
      displayItems.sort(function(a, b) {{
        var ta = a.start_time || "";
        var tb = b.start_time || "";
        if (ta < tb) return -1;
        if (ta > tb) return 1;
        return 0;
      }});

      // ページ計算
      totalPages = Math.ceil(displayItems.length / ROWS_PER_PAGE);

      // 初期ページを現在時刻ベースで決定
      currentPage = findCurrentTimePage(displayItems);
      if (currentPage >= totalPages) currentPage = 0;

      // 現在ページを描画
      renderCurrentPage();

      // フッター更新
      updateFooter(data, displayItems.length);

      // ページローテーション開始
      if (totalPages > 1) {{
        startPageRotation();
      }}
    }}

    // ===== 現在ページの描画 =====
    function renderCurrentPage() {{
      if (displayItems.length === 0) return;

      var body = document.getElementById("board-body");
      var start = currentPage * ROWS_PER_PAGE;
      var pageItems = displayItems.slice(start, start + ROWS_PER_PAGE);
      var rowHeight = calcRowHeight(pageItems.length);

      var rows = pageItems.map(function(item, idx) {{
        var globalIdx = start + idx;
        var rowId = item.name + "|" + item.date + "|" + item.start_time;
        return renderRow(item, globalIdx, rowHeight, false, rowId);
      }});

      body.innerHTML = rows.join("");

      // ページインジケーター更新
      var indicator = document.getElementById("page-indicator");
      if (totalPages > 1) {{
        indicator.textContent = (currentPage + 1) + " / " + totalPages;
      }} else {{
        indicator.textContent = "";
      }}
    }}

    // ===== ページローテーション =====
    function startPageRotation() {{
      if (pageRotateTimer) return;
      cycleComplete = false;
      var startPage = currentPage;  // 巡回開始ページを記憶

      pageRotateTimer = setInterval(function() {{
        var body = document.getElementById("board-body");
        body.classList.add("fade-out");
        body.classList.remove("fade-in");

        setTimeout(function() {{
          var nextPage = (currentPage + 1) % totalPages;

          // 全ページ巡回完了チェック
          // 巡回開始ページに戻ったら、現在時刻で再計算
          if (nextPage === startPage) {{
            currentPage = findCurrentTimePage(displayItems);
            if (currentPage >= totalPages) currentPage = 0;
            startPage = currentPage;  // 新しい起点を設定
          }} else {{
            currentPage = nextPage;
          }}

          renderCurrentPage();
          body.classList.remove("fade-out");
          body.classList.add("fade-in");
        }}, 300);
      }}, PAGE_ROTATE_SEC * 1000);
    }}

    function stopPageRotation() {{
      if (pageRotateTimer) {{
        clearInterval(pageRotateTimer);
        pageRotateTimer = null;
      }}
    }}

    // ===== 1分ごとの現在時刻再評価 =====
    function startTimeCheck() {{
      setInterval(function() {{
        var now = new Date();
        var currentMinute = now.getHours() * 60 + now.getMinutes();

        if (currentMinute !== lastTimeCheckMinute) {{
          lastTimeCheckMinute = currentMinute;
          var bestPage = findCurrentTimePage(displayItems);
          if (bestPage !== currentPage && displayItems.length > 0) {{
            // 時間が進んだのでページジャンプ
            stopPageRotation();
            currentPage = bestPage;
            if (currentPage >= totalPages) currentPage = 0;

            var body = document.getElementById("board-body");
            body.classList.add("fade-out");
            body.classList.remove("fade-in");
            setTimeout(function() {{
              renderCurrentPage();
              body.classList.remove("fade-out");
              body.classList.add("fade-in");
              if (totalPages > 1) {{
                startPageRotation();
              }}
            }}, 300);
          }}
        }}
      }}, 60000);
    }}

    // ===== フッター更新 =====
    function updateFooter(data, displayCount) {{
      if (data.updated_at) {{
        var dt = new Date(data.updated_at);
        var hh = String(dt.getHours()).padStart(2, "0");
        var mm = String(dt.getMinutes()).padStart(2, "0");
        var ss = String(dt.getSeconds()).padStart(2, "0");
        document.getElementById("update-time").textContent = hh + ":" + mm + ":" + ss;
      }}
      document.getElementById("total-count").textContent = displayCount;

      var indicator = document.getElementById("page-indicator");
      if (totalPages > 1) {{
        indicator.textContent = (currentPage + 1) + " / " + totalPages;
      }} else {{
        indicator.textContent = "";
      }}
    }}

    // ===== 初期化 =====
    function initBoard() {{
      startClock();
      renderBoard(BOARD_DATA);
      startTimeCheck();
    }}

    document.addEventListener("DOMContentLoaded", initBoard);
  </script>
</body>
</html>'''


def main():
    print("工事予定表スタンドアロンHTML生成")
    print("=" * 40)

    # 1. Excelデータ読み込み
    print("Excelデータ読み込み中...")
    data = load_construction_data()
    item_count = len(data.get("items", []))
    print(f"  -> {item_count} 件のデータを読み込みました")

    if data.get("error"):
        print(f"  [警告] {data['error']}")

    # 2. CSS読み込み
    print("CSS読み込み中...")
    css = read_css()
    print(f"  -> {len(css)} 文字")

    # 3. HTML生成
    print("HTML生成中...")
    html = generate_html(data, css)

    # 4. デスクトップに保存
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"  -> {len(html)} 文字")
    print()
    print(f"完了! {item_count} 件のデータを含むHTMLを生成しました。")
    print(f"出力先: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
