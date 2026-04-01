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
    </div>
  </header>

  <div class="board-table-header">
    <div class="col col-client">客先</div>
    <div class="col col-title">工事件名</div>
    <div class="col col-our-person">現場</div>
    <div class="col col-safety">安品</div>
    <div class="col col-partner">協力会社</div>
    <div class="col col-partner-person">担当</div>
    <div class="col col-work">作業内容</div>
    <div class="col col-time">昼/夜</div>
    <div class="col col-priority">重点作業</div>
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
    var BOARD_DATA = {data_json};
    var MIN_ROW_HEIGHT = 72;
    var PAGE_ROTATE_SEC = 15;
    var GROUP_ROTATE_CYCLES = 1;

    // ===== グローバル変数 =====
    var ROWS_PER_PAGE = 10;
    var currentPage = 0;
    var totalPages = 1;
    var dateGroups = [];
    var currentGroupIdx = 0;
    var groupCycleCount = 0;
    var pageRotateTimer = null;

    var WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

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
        var now = new Date();
        var hh = String(now.getHours()).padStart(2, "0");
        var mm = String(now.getMinutes()).padStart(2, "0");
        var ss = String(now.getSeconds()).padStart(2, "0");
        document.getElementById("clock").textContent = hh + ":" + mm + ":" + ss;
      }}
      update();
      setInterval(update, 1000);
    }}

    // ===== 日付フォーマット =====
    function formatDateLabel(dateStr) {{
      var parts = dateStr.split("-");
      var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      var m = d.getMonth() + 1;
      var day = d.getDate();
      var w = WEEKDAYS[d.getDay()];
      return m + "/" + day + " " + w + "曜日";
    }}

    // ===== 今日+明日フィルタリング =====
    function filterTodayTomorrow(items, todayStr, tomorrowStr) {{
      if (!items || items.length === 0) return [];
      var filtered = items.filter(function(item) {{
        return item.date === todayStr || item.date === tomorrowStr;
      }});
      if (filtered.length > 0) return filtered;
      return items;
    }}

    // ===== 日付グループ構築 =====
    function buildDateGroups(items, todayStr, tomorrowStr) {{
      var todayItems = items.filter(function(item) {{ return item.date === todayStr; }});
      var tomorrowItems = items.filter(function(item) {{ return item.date === tomorrowStr; }});
      var otherItems = items.filter(function(item) {{
        return item.date !== todayStr && item.date !== tomorrowStr;
      }});
      var groups = [];
      if (todayItems.length > 0) {{
        groups.push({{ date: todayStr, label: "本日", dateFormatted: formatDateLabel(todayStr), items: todayItems }});
      }}
      if (tomorrowItems.length > 0) {{
        groups.push({{ date: tomorrowStr, label: "明日", dateFormatted: formatDateLabel(tomorrowStr), items: tomorrowItems }});
      }}
      if (otherItems.length > 0 && groups.length === 0) {{
        var dates = [];
        otherItems.forEach(function(i) {{ if (dates.indexOf(i.date) === -1) dates.push(i.date); }});
        dates.sort();
        dates.forEach(function(dt) {{
          groups.push({{ date: dt, label: "", dateFormatted: formatDateLabel(dt), items: otherItems.filter(function(i) {{ return i.date === dt; }}) }});
        }});
      }}
      return groups;
    }}

    // ===== 表示件数を画面サイズから自動計算 =====
    function calcRowsPerPage() {{
      var available = window.innerHeight - 88 - 62 - 40;
      return Math.max(4, Math.floor(available / MIN_ROW_HEIGHT));
    }}

    // ===== 夜間作業判定 =====
    function isNightWork(workTime) {{
      return workTime === "夜";
    }}

    // ===== 行レンダリング =====
    function renderRow(item, index, rowHeight, rowId) {{
      var evenOdd = index % 2 === 0 ? "row-even" : "row-odd";
      var priorityClass = item.is_priority ? " row-priority" : "";
      var nightClass = isNightWork(item.work_time) ? " row-night" : "";
      var priorityBadge = item.is_priority
        ? '<span class="priority-badge priority-yes">有</span>'
        : '<span class="priority-badge priority-no">無</span>';

      return '<div class="board-row ' + evenOdd + priorityClass + nightClass + '" data-id="' + escapeAttr(rowId) + '" style="min-height:' + rowHeight + 'px;">' +
        '<div class="col col-client">' + escapeHtml(item.client) + '</div>' +
        '<div class="col col-title">' + escapeHtml(item.title) + '</div>' +
        '<div class="col col-our-person">' + escapeHtml(item.our_person) + '</div>' +
        '<div class="col col-safety">' + escapeHtml(item.safety_person) + '</div>' +
        '<div class="col col-partner">' + escapeHtml(item.partner) + '</div>' +
        '<div class="col col-partner-person">' + escapeHtml(item.partner_person) + '</div>' +
        '<div class="col col-work">' + escapeHtml(item.work_content) + '</div>' +
        '<div class="col col-time">' + escapeHtml(item.work_time) + '</div>' +
        '<div class="col col-priority">' + priorityBadge + '</div>' +
      '</div>';
    }}

    // ===== ヘッダー日付表示更新 =====
    function updateHeaderForGroup(group) {{
      var todayEl = document.getElementById("header-today");
      if (!todayEl) return;
      var labelClass = group.label === "明日" ? "header-today-label label-tomorrow" : "header-today-label";
      var label = group.label || "";
      todayEl.innerHTML = '<span class="' + labelClass + '">' + escapeHtml(label) + '</span>' + escapeHtml(group.dateFormatted);
    }}

    // ===== ボード表示 =====
    function renderBoard(data) {{
      var body = document.getElementById("board-body");

      if (data.error) {{
        body.innerHTML =
          '<div class="board-message">' +
            '<div class="board-message-icon">&#9888;</div>' +
            '<div class="board-message-error">' + escapeHtml(data.error) + '</div>' +
          '</div>';
        updateFooter(data, 0);
        stopAllTimers();
        return;
      }}

      var tomorrowStr = data.tomorrow || "";
      var filtered = filterTodayTomorrow(data.items, data.today, tomorrowStr);

      if (filtered.length === 0) {{
        body.innerHTML =
          '<div class="board-message">' +
            '<div class="board-message-icon">&#128736;</div>' +
            '<div class="board-message-text">本日の工事予定はありません</div>' +
          '</div>';
        updateFooter(data, 0);
        stopAllTimers();
        return;
      }}

      dateGroups = buildDateGroups(filtered, data.today, tomorrowStr);
      dateGroups.forEach(function(g) {{
        g.totalPages = Math.ceil(g.items.length / ROWS_PER_PAGE);
      }});

      if (currentGroupIdx >= dateGroups.length) currentGroupIdx = 0;
      var group = dateGroups[currentGroupIdx];
      totalPages = group.totalPages;
      if (currentPage >= totalPages) currentPage = 0;

      renderCurrentGroupPage();
      updateHeaderForGroup(group);
      updateFooter(data, filtered.length);

      stopAllTimers();
      if (totalPages > 1 || dateGroups.length > 1) {{
        startPageRotation();
      }}
    }}

    // ===== 現在グループ・ページ描画 =====
    function renderCurrentGroupPage() {{
      if (dateGroups.length === 0) return;
      var group = dateGroups[currentGroupIdx];
      var body = document.getElementById("board-body");
      var start = currentPage * ROWS_PER_PAGE;
      var pageItems = group.items.slice(start, start + ROWS_PER_PAGE);

      var available = window.innerHeight - 88 - 62 - 40;
      var rowHeight = pageItems.length > 0 ? Math.min(64, Math.floor(available / pageItems.length)) : 64;

      var html = pageItems.map(function(item, idx) {{
        var rowId = item.title + "|" + item.date + "|" + item.partner;
        return renderRow(item, idx, rowHeight, rowId);
      }});

      body.innerHTML = html.join("");
      updatePageIndicator();
    }}

    // ===== ページローテーション =====
    function startPageRotation() {{
      if (pageRotateTimer) return;
      groupCycleCount = 0;
      pageRotateTimer = setInterval(function() {{
        var body = document.getElementById("board-body");
        var group = dateGroups[currentGroupIdx];
        body.classList.add("fade-out");
        body.classList.remove("fade-in");

        setTimeout(function() {{
          var nextPage = currentPage + 1;
          if (nextPage >= group.totalPages) {{
            groupCycleCount++;
            if (dateGroups.length > 1 && groupCycleCount >= GROUP_ROTATE_CYCLES) {{
              groupCycleCount = 0;
              currentGroupIdx = (currentGroupIdx + 1) % dateGroups.length;
              currentPage = 0;
              var newGroup = dateGroups[currentGroupIdx];
              totalPages = newGroup.totalPages;
              updateHeaderForGroup(newGroup);
            }} else {{
              currentPage = 0;
            }}
          }} else {{
            currentPage = nextPage;
          }}
          renderCurrentGroupPage();
          body.classList.remove("fade-out");
          body.classList.add("fade-in");
        }}, 300);
      }}, PAGE_ROTATE_SEC * 1000);
    }}

    function stopAllTimers() {{
      if (pageRotateTimer) {{
        clearInterval(pageRotateTimer);
        pageRotateTimer = null;
      }}
    }}

    function updatePageIndicator() {{
      var indicator = document.getElementById("page-indicator");
      var group = dateGroups[currentGroupIdx];
      if (!group) return;
      if (group.totalPages > 1) {{
        indicator.textContent = (currentPage + 1) + " / " + group.totalPages;
      }} else {{
        indicator.textContent = "";
      }}
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
      updatePageIndicator();
    }}

    // ===== フルスクリーン =====
    function requestFullscreen() {{
      var el = document.documentElement;
      if (el.requestFullscreen) el.requestFullscreen();
      else if (el.webkitRequestFullscreen) el.webkitRequestFullscreen();
      else if (el.msRequestFullscreen) el.msRequestFullscreen();
    }}

    function toggleFullscreen() {{
      if (document.fullscreenElement || document.webkitFullscreenElement) {{
        if (document.exitFullscreen) document.exitFullscreen();
        else if (document.webkitExitFullscreen) document.webkitExitFullscreen();
      }} else {{
        requestFullscreen();
      }}
    }}

    // ===== リサイズ対応 =====
    function onResize() {{
      var newRows = calcRowsPerPage();
      if (newRows !== ROWS_PER_PAGE) {{
        ROWS_PER_PAGE = newRows;
        stopAllTimers();
        if (dateGroups.length > 0) {{
          dateGroups.forEach(function(g) {{
            g.totalPages = Math.ceil(g.items.length / ROWS_PER_PAGE);
          }});
          var group = dateGroups[currentGroupIdx];
          totalPages = group.totalPages;
          if (currentPage >= totalPages) currentPage = 0;
          renderCurrentGroupPage();
          if (totalPages > 1 || dateGroups.length > 1) startPageRotation();
        }}
      }}
    }}

    // ===== 初期化 =====
    function initBoard() {{
      ROWS_PER_PAGE = calcRowsPerPage();
      startClock();
      renderBoard(BOARD_DATA);
      window.addEventListener("resize", onResize);
      document.addEventListener("click", toggleFullscreen);
      requestFullscreen();
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
