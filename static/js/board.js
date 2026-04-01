/* =============================================
   工事予定表サイネージ - Warm Corporate Style
   ============================================= */

// ---------- グローバル変数 ----------
const MIN_ROW_HEIGHT = 72; // 1行の最低高さ(px)
let ROWS_PER_PAGE = 10;
let PAGE_ROTATE_SEC = 15;
let currentPage = 0; // 現在の日付グループ内ページ
let totalPages = 1;
let previousDataJSON = "";
let previousItems = [];
let pageRotateTimer = null;
let dateGroups = []; // [{date, label, dateFormatted, items, pages}]
let currentGroupIdx = 0;
let groupRotateTimer = null;
const GROUP_ROTATE_CYCLES = 1; // 各グループを何周するか
let groupCycleCount = 0;

const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

// ---------- 時計機能 ----------
function startClock() {
  function update() {
    const now = new Date();

    // 時刻
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    document.getElementById("clock").textContent = `${hh}:${mm}:${ss}`;

    // 日付
    const y = now.getFullYear();
    const m = now.getMonth() + 1;
    const d = now.getDate();
    const w = WEEKDAYS[now.getDay()];
    const dateEl = document.getElementById("date");
    if (dateEl) dateEl.textContent = `${y}年${m}月${d}日(${w})`;

    // ヘッダー中央の TODAY 表示
    const todayEl = document.getElementById("header-today");
    if (todayEl) {
      todayEl.textContent = `${m}/${d} ${w}曜日`;
    }
  }

  update();
  setInterval(update, 1000);
}

// ---------- 設定取得 ----------
async function fetchConfig() {
  try {
    const res = await fetch("/api/config");
    if (res.ok) {
      const cfg = await res.json();
      PAGE_ROTATE_SEC = cfg.page_rotate_sec || 15;
    }
  } catch (e) {
    // デフォルト値のまま
  }
  // 画面サイズから表示件数を自動計算
  ROWS_PER_PAGE = calcRowsPerPage();
}

// ---------- データ取得・表示 ----------
async function fetchAndRender() {
  try {
    const res = await fetch("/api/data");
    if (!res.ok) {
      renderError("サーバーエラー: " + res.status);
      return;
    }
    const data = await res.json();
    renderBoard(data);
  } catch (e) {
    renderError("通信エラー: サーバーに接続できません");
  }
}

// ---------- 日付フォーマット ----------
function formatDateLabel(dateStr) {
  const parts = dateStr.split("-");
  const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const w = WEEKDAYS[d.getDay()];
  return `${m}/${day} ${w}曜日`;
}

// ---------- 今日+明日フィルタリング ----------
function filterTodayTomorrow(items, todayStr, tomorrowStr) {
  if (!items || items.length === 0) return [];
  const filtered = items.filter(function(item) {
    return item.date === todayStr || item.date === tomorrowStr;
  });
  if (filtered.length > 0) return filtered;
  return items;
}

// ---------- 日付グループ構築 ----------
function buildDateGroups(items, todayStr, tomorrowStr) {
  const todayItems = items.filter(function(item) { return item.date === todayStr; });
  const tomorrowItems = items.filter(function(item) { return item.date === tomorrowStr; });
  const otherItems = items.filter(function(item) {
    return item.date !== todayStr && item.date !== tomorrowStr;
  });

  const groups = [];
  if (todayItems.length > 0) {
    groups.push({
      date: todayStr,
      label: "本日",
      dateFormatted: formatDateLabel(todayStr),
      items: todayItems,
    });
  }
  if (tomorrowItems.length > 0) {
    groups.push({
      date: tomorrowStr,
      label: "明日",
      dateFormatted: formatDateLabel(tomorrowStr),
      items: tomorrowItems,
    });
  }
  if (otherItems.length > 0 && groups.length === 0) {
    const dates = [...new Set(otherItems.map(function(i) { return i.date; }))].sort();
    dates.forEach(function(d) {
      groups.push({
        date: d,
        label: "",
        dateFormatted: formatDateLabel(d),
        items: otherItems.filter(function(i) { return i.date === d; }),
      });
    });
  }
  return groups;
}

function renderBoard(data) {
  const body = document.getElementById("board-body");

  // エラー表示
  if (data.error) {
    body.innerHTML = `
      <div class="board-message">
        <div class="board-message-icon">&#9888;</div>
        <div class="board-message-error">${escapeHtml(data.error)}</div>
      </div>`;
    updateFooter(data, 0);
    stopAllTimers();
    return;
  }

  // 今日+明日のフィルタリング
  const tomorrowStr = data.tomorrow || "";
  const filtered = filterTodayTomorrow(data.items, data.today, tomorrowStr);

  // 空の場合
  if (filtered.length === 0) {
    body.innerHTML = `
      <div class="board-message">
        <div class="board-message-icon">&#128736;</div>
        <div class="board-message-text">本日の工事予定はありません</div>
      </div>`;
    updateFooter(data, 0);
    stopAllTimers();
    return;
  }

  // 差分検出
  const newDataJSON = JSON.stringify(filtered);
  const hasChanged = newDataJSON !== previousDataJSON;
  previousDataJSON = newDataJSON;
  previousItems = filtered;

  // 日付グループ構築
  dateGroups = buildDateGroups(filtered, data.today, tomorrowStr);

  // 各グループのページ数を計算
  dateGroups.forEach(function(g) {
    g.totalPages = Math.ceil(g.items.length / ROWS_PER_PAGE);
  });

  // 初期表示
  if (currentGroupIdx >= dateGroups.length) currentGroupIdx = 0;
  const group = dateGroups[currentGroupIdx];
  totalPages = group.totalPages;
  if (currentPage >= totalPages) currentPage = 0;

  renderCurrentGroupPage(hasChanged);
  updateHeaderForGroup(group);
  updateFooter(data, filtered.length);

  // タイマー開始
  stopAllTimers();
  if (totalPages > 1 || dateGroups.length > 1) {
    startPageRotation();
  }
}

// ---------- ヘッダーの日付表示を現在グループに合わせる ----------
function updateHeaderForGroup(group) {
  const todayEl = document.getElementById("header-today");
  if (!todayEl) return;
  const labelClass = group.label === "明日" ? "header-today-label label-tomorrow" : "header-today-label";
  const label = group.label || "";
  todayEl.innerHTML = `<span class="${labelClass}">${escapeHtml(label)}</span>${escapeHtml(group.dateFormatted)}`;
}

// ---------- 現在グループ・現在ページの描画 ----------
function renderCurrentGroupPage(hasChanged) {
  if (dateGroups.length === 0) return;

  const group = dateGroups[currentGroupIdx];
  const body = document.getElementById("board-body");
  const start = currentPage * ROWS_PER_PAGE;
  const pageItems = group.items.slice(start, start + ROWS_PER_PAGE);

  const available = window.innerHeight - 88 - 62 - 40;
  const rowHeight = pageItems.length > 0 ? Math.min(64, Math.floor(available / pageItems.length)) : 64;

  let rowIdx = 0;
  const html = pageItems.map(function(item) {
    const rowId = item.title + "|" + item.date + "|" + item.partner;
    const prevItem = findPreviousItem(rowId);
    const changed = hasChanged && (!prevItem || JSON.stringify(prevItem) !== JSON.stringify(item));
    const result = renderRow(item, rowIdx, rowHeight, changed, rowId);
    rowIdx++;
    return result;
  });

  body.innerHTML = html.join("");

  // フリップアニメーション完了後にクラス除去
  body.querySelectorAll(".flip-enter").forEach(function(el) {
    el.addEventListener("animationend", function() {
      el.classList.remove("flip-enter");
    }, { once: true });
  });

  updatePageIndicator();
}

// ---------- 夜間作業判定 ----------
function isNightWork(workTime) {
  return workTime === "夜";
}

// ---------- 行レンダリング ----------
function renderRow(item, index, rowHeight, changed, rowId) {
  const evenOdd = index % 2 === 0 ? "row-even" : "row-odd";
  const flipClass = changed ? " flip-enter" : "";
  const priorityClass = item.is_priority ? " row-priority" : "";
  const nightClass = isNightWork(item.work_time) ? " row-night" : "";

  const priorityBadge = item.is_priority
    ? `<span class="priority-badge priority-yes">有</span>`
    : `<span class="priority-badge priority-no">無</span>`;

  return `<div class="board-row ${evenOdd}${flipClass}${priorityClass}${nightClass}" data-id="${escapeAttr(rowId)}" style="min-height:${rowHeight}px;">
  <div class="col col-client">${escapeHtml(item.client)}</div>
  <div class="col col-title">${escapeHtml(item.title)}</div>
  <div class="col col-our-person">${escapeHtml(item.our_person)}</div>
  <div class="col col-safety">${escapeHtml(item.safety_person)}</div>
  <div class="col col-partner">${escapeHtml(item.partner)}</div>
  <div class="col col-partner-person">${escapeHtml(item.partner_person)}</div>
  <div class="col col-work">${escapeHtml(item.work_content)}</div>
  <div class="col col-time">${escapeHtml(item.work_time)}</div>
  <div class="col col-priority">${priorityBadge}</div>
</div>`;
}

// ---------- 表示件数・行の高さを画面サイズから自動計算 ----------
function calcRowsPerPage() {
  const available = window.innerHeight - 88 - 62 - 40; // header + table-header + footer
  return Math.max(4, Math.floor(available / MIN_ROW_HEIGHT));
}

// ---------- 前回データ検索 ----------
function findPreviousItem(rowId) {
  for (const item of previousItems) {
    const id = item.title + "|" + item.date + "|" + item.partner;
    if (id === rowId) return item;
  }
  return null;
}

// ---------- ページローテーション ----------
function startPageRotation() {
  if (pageRotateTimer) return;
  groupCycleCount = 0;
  pageRotateTimer = setInterval(function() {
    rotatePage();
  }, PAGE_ROTATE_SEC * 1000);
}

function stopAllTimers() {
  if (pageRotateTimer) {
    clearInterval(pageRotateTimer);
    pageRotateTimer = null;
  }
}

function rotatePage() {
  const body = document.getElementById("board-body");
  const group = dateGroups[currentGroupIdx];

  body.classList.add("fade-out");
  body.classList.remove("fade-in");

  setTimeout(function() {
    const nextPage = currentPage + 1;
    if (nextPage >= group.totalPages) {
      // このグループの最後のページ → 次のグループへ or 周回
      groupCycleCount++;
      if (dateGroups.length > 1 && groupCycleCount >= GROUP_ROTATE_CYCLES) {
        // 次の日付グループへ切替
        groupCycleCount = 0;
        currentGroupIdx = (currentGroupIdx + 1) % dateGroups.length;
        currentPage = 0;
        const newGroup = dateGroups[currentGroupIdx];
        totalPages = newGroup.totalPages;
        updateHeaderForGroup(newGroup);
      } else {
        // 同じグループの先頭に戻る
        currentPage = 0;
      }
    } else {
      currentPage = nextPage;
    }

    renderCurrentGroupPage(false);
    body.classList.remove("fade-out");
    body.classList.add("fade-in");
  }, 300);
}

function updatePageIndicator() {
  const indicator = document.getElementById("page-indicator");
  const group = dateGroups[currentGroupIdx];
  if (!group) return;
  if (group.totalPages > 1) {
    indicator.textContent = `${currentPage + 1} / ${group.totalPages}`;
  } else {
    indicator.textContent = "";
  }
}

// ---------- フッター更新 ----------
function updateFooter(data, displayCount) {
  if (data.updated_at) {
    const dt = new Date(data.updated_at);
    const hh = String(dt.getHours()).padStart(2, "0");
    const mm = String(dt.getMinutes()).padStart(2, "0");
    const ss = String(dt.getSeconds()).padStart(2, "0");
    document.getElementById("update-time").textContent = `${hh}:${mm}:${ss}`;
  }

  document.getElementById("total-count").textContent = displayCount;

  const indicator = document.getElementById("page-indicator");
  if (totalPages > 1) {
    indicator.textContent = `${currentPage + 1} / ${totalPages}`;
  } else {
    indicator.textContent = "";
  }
}

// ---------- エラー表示 ----------
function renderError(message) {
  const body = document.getElementById("board-body");
  body.innerHTML = `
    <div class="board-message">
      <div class="board-message-icon">&#9888;</div>
      <div class="board-message-error">${escapeHtml(message)}</div>
    </div>`;
  stopPageRotation();
}

// ---------- HTMLエスケープ ----------
function escapeHtml(str) {
  if (str == null) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttr(str) {
  return escapeHtml(str);
}

// ---------- フルスクリーン ----------
function requestFullscreen() {
  const el = document.documentElement;
  if (el.requestFullscreen) el.requestFullscreen();
  else if (el.webkitRequestFullscreen) el.webkitRequestFullscreen();
  else if (el.msRequestFullscreen) el.msRequestFullscreen();
}

function toggleFullscreen() {
  if (document.fullscreenElement || document.webkitFullscreenElement) {
    if (document.exitFullscreen) document.exitFullscreen();
    else if (document.webkitExitFullscreen) document.webkitExitFullscreen();
  } else {
    requestFullscreen();
  }
}

// ---------- リサイズ対応 ----------
function onResize() {
  const newRows = calcRowsPerPage();
  if (newRows !== ROWS_PER_PAGE) {
    ROWS_PER_PAGE = newRows;
    stopAllTimers();
    if (dateGroups.length > 0) {
      dateGroups.forEach(function(g) {
        g.totalPages = Math.ceil(g.items.length / ROWS_PER_PAGE);
      });
      const group = dateGroups[currentGroupIdx];
      totalPages = group.totalPages;
      if (currentPage >= totalPages) currentPage = 0;
      renderCurrentGroupPage(false);
      if (totalPages > 1 || dateGroups.length > 1) startPageRotation();
    }
  }
}

// ---------- 初期化 ----------
async function initBoard() {
  startClock();
  await fetchConfig();
  await fetchAndRender();
  setInterval(fetchAndRender, REFRESH_INTERVAL);

  // リサイズ・全画面切替時に表示件数を再計算
  window.addEventListener("resize", onResize);

  // クリックで全画面トグル
  document.addEventListener("click", toggleFullscreen);
  // 初回は自動で全画面（ユーザー操作が必要なブラウザもある）
  requestFullscreen();
}

document.addEventListener("DOMContentLoaded", initBoard);
