/* =============================================
   工事予定表サイネージ - Warm Corporate Style
   ============================================= */

// ---------- グローバル変数 ----------
let ROWS_PER_PAGE = 12;
let PAGE_ROTATE_SEC = 15;
let currentPage = 0;
let totalPages = 1;
let previousDataJSON = "";
let previousItems = [];
let pageRotateTimer = null;

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
    document.getElementById("date").textContent = `${y}年${m}月${d}日(${w})`;

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
      ROWS_PER_PAGE = cfg.rows_per_page || 12;
      PAGE_ROTATE_SEC = cfg.page_rotate_sec || 15;
    }
  } catch (e) {
    // デフォルト値のまま
  }
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
    stopPageRotation();
    return;
  }

  // 「今日」のフィルタリング
  let displayItems = filterToday(data.items, data.today);

  // 空の場合
  if (displayItems.length === 0) {
    body.innerHTML = `
      <div class="board-message">
        <div class="board-message-icon">&#128736;</div>
        <div class="board-message-text">本日の工事予定はありません</div>
      </div>`;
    updateFooter(data, 0);
    stopPageRotation();
    return;
  }

  // ページ計算
  totalPages = Math.ceil(displayItems.length / ROWS_PER_PAGE);
  if (currentPage >= totalPages) {
    currentPage = 0;
  }

  // 現在ページのアイテムをスライス
  const start = currentPage * ROWS_PER_PAGE;
  const pageItems = displayItems.slice(start, start + ROWS_PER_PAGE);

  // 差分検出
  const newDataJSON = JSON.stringify(displayItems);
  const hasChanged = newDataJSON !== previousDataJSON;

  // 行のHTML生成
  const rowHeight = calcRowHeight(pageItems.length);
  const rows = pageItems.map((item, idx) => {
    const globalIdx = start + idx;
    const rowId = item.name + "|" + item.date + "|" + item.start_time;
    const prevItem = findPreviousItem(rowId);
    const changed = hasChanged && (!prevItem || JSON.stringify(prevItem) !== JSON.stringify(item));
    return renderRow(item, globalIdx, rowHeight, changed, rowId);
  });

  body.innerHTML = rows.join("");

  // フリップアニメーション完了後にクラス除去
  body.querySelectorAll(".flip-enter").forEach((el) => {
    el.addEventListener("animationend", () => {
      el.classList.remove("flip-enter");
    }, { once: true });
  });

  // データ保持
  previousDataJSON = newDataJSON;
  previousItems = displayItems;

  // フッター更新
  updateFooter(data, displayItems.length);

  // ページローテーション制御
  if (totalPages > 1) {
    startPageRotation();
  } else {
    stopPageRotation();
  }
}

// ---------- 「今日」フィルタリング ----------
function filterToday(items, todayStr) {
  if (!items || items.length === 0) return [];
  const todayItems = items.filter((item) => item.date === todayStr);
  if (todayItems.length > 0) return todayItems;
  return items;
}

// ---------- 行レンダリング ----------
function renderRow(item, index, rowHeight, changed, rowId) {
  const evenOdd = index % 2 === 0 ? "row-even" : "row-odd";
  const flipClass = changed ? " flip-enter" : "";
  const progressComplete = item.progress >= 100 ? " complete" : "";

  const timeDisplay = item.end_time
    ? `${escapeHtml(item.start_time)} <span class="time-dash">―</span> ${escapeHtml(item.end_time)}`
    : escapeHtml(item.start_time);

  return `<div class="board-row ${evenOdd}${flipClass}" data-id="${escapeAttr(rowId)}" style="height:${rowHeight}px;">
  <div class="col col-time">${timeDisplay}</div>
  <div class="col col-name">${escapeHtml(item.name)}</div>
  <div class="col col-location">${escapeHtml(item.location)}</div>
  <div class="col col-status"><span class="status-badge status-${item.status_code}"><span class="status-dot"></span>${escapeHtml(item.status)}</span></div>
  <div class="col col-person">${escapeHtml(item.person)}</div>
  <div class="col col-progress">
    <div class="progress-wrapper">
      <div class="progress-bar"><div class="progress-fill" style="width:${item.progress}%"></div></div>
    </div>
    <span class="progress-text${progressComplete}">${item.progress}%</span>
  </div>
</div>`;
}

// ---------- 行の高さ計算 ----------
function calcRowHeight(rowCount) {
  if (rowCount === 0) return 64;
  const available = window.innerHeight - 120 - 60 - 52; // header + table-header + footer
  return Math.floor(available / Math.max(rowCount, 1));
}

// ---------- 前回データ検索 ----------
function findPreviousItem(rowId) {
  for (const item of previousItems) {
    const id = item.name + "|" + item.date + "|" + item.start_time;
    if (id === rowId) return item;
  }
  return null;
}

// ---------- ページローテーション ----------
function startPageRotation() {
  if (pageRotateTimer) return;
  pageRotateTimer = setInterval(() => {
    rotatePage();
  }, PAGE_ROTATE_SEC * 1000);
}

function stopPageRotation() {
  if (pageRotateTimer) {
    clearInterval(pageRotateTimer);
    pageRotateTimer = null;
  }
}

function rotatePage() {
  const body = document.getElementById("board-body");
  body.classList.add("fade-out");
  body.classList.remove("fade-in");

  setTimeout(() => {
    currentPage = (currentPage + 1) % totalPages;
    rerenderCurrentPage();
    body.classList.remove("fade-out");
    body.classList.add("fade-in");
  }, 300);
}

function rerenderCurrentPage() {
  if (previousItems.length === 0) return;

  const body = document.getElementById("board-body");
  const start = currentPage * ROWS_PER_PAGE;
  const pageItems = previousItems.slice(start, start + ROWS_PER_PAGE);
  const rowHeight = calcRowHeight(pageItems.length);

  const rows = pageItems.map((item, idx) => {
    const globalIdx = start + idx;
    const rowId = item.name + "|" + item.date + "|" + item.start_time;
    return renderRow(item, globalIdx, rowHeight, false, rowId);
  });

  body.innerHTML = rows.join("");

  const indicator = document.getElementById("page-indicator");
  if (totalPages > 1) {
    indicator.textContent = `${currentPage + 1} / ${totalPages}`;
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

// ---------- 初期化 ----------
async function initBoard() {
  startClock();
  await fetchConfig();
  await fetchAndRender();
  setInterval(fetchAndRender, REFRESH_INTERVAL);
}

document.addEventListener("DOMContentLoaded", initBoard);
