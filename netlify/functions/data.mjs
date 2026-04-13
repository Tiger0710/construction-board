/**
 * 工事予定表 データ API
 *
 * GET  ?month=2604         → _2604.json を返却
 * GET  ?signage=true       → 前月+当月+翌月を統合し items[] 形式で返却
 * PUT  { month, data, sha } → _{month}.json を GitHub に保存
 */

const REPO = "Tiger0710/construction-board";
const BRANCH = "main";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
  "Content-Type": "application/json",
};

function getJSTNow() {
  return new Date(Date.now() + 9 * 60 * 60 * 1000);
}

function fmtDate(d) {
  return d.toISOString().split("T")[0];
}

function getTargetMonths() {
  const now = getJSTNow();
  const y = now.getUTCFullYear();
  const m = now.getUTCMonth();
  const fmt = (yr, mo) =>
    String(yr % 100).padStart(2, "0") + String(mo + 1).padStart(2, "0");
  return [
    m === 0 ? fmt(y - 1, 11) : fmt(y, m - 1),
    fmt(y, m),
    m === 11 ? fmt(y + 1, 0) : fmt(y, m + 1),
  ];
}

function ghHeaders(token) {
  return {
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github.v3+json",
    "User-Agent": "construction-board-data",
  };
}

async function fetchMonthData(token, month) {
  const path = `_${month}.json`;
  const url = `https://api.github.com/repos/${REPO}/contents/${path}?ref=${BRANCH}`;
  const res = await fetch(url, { headers: ghHeaders(token) });
  if (!res.ok) return null;
  const file = await res.json();
  const content = JSON.parse(
    Buffer.from(file.content, "base64").toString("utf-8")
  );
  content._sha = file.sha;
  return content;
}

function expandToItems(data) {
  const items = [];
  const projects = data.projects || [];
  const daily = data.daily || {};

  for (const proj of projects) {
    const start = new Date(proj.start_date + "T00:00:00Z");
    const end = new Date(proj.end_date + "T00:00:00Z");
    const d = new Date(start);

    while (d <= end) {
      const ds = fmtDate(d);
      const key = `${proj.id}/${ds}`;
      const dd = daily[key];

      const dayActive = dd ? dd.day !== false : true;
      const nightActive = dd ? dd.night === true : false;

      if (dayActive) {
        const pri = dd?.day_priority || "";
        items.push({
          date: ds,
          client: proj.client,
          title: proj.title,
          our_person: dd?.day_our_person || proj.our_person || "",
          safety_person: dd?.day_safety_person || proj.safety_person || "",
          partner: proj.partner || "",
          partner_person: proj.partner_person || "",
          work_content: dd?.day_work || "",
          work_time: "昼",
          priority: pri,
          is_priority: pri === "有",
          priority_detail: dd?.day_priority_detail || "",
        });
      }

      if (nightActive) {
        const pri = dd?.night_priority || "";
        items.push({
          date: ds,
          client: proj.client,
          title: proj.title,
          our_person: dd?.night_our_person || proj.our_person || "",
          safety_person: dd?.night_safety_person || proj.safety_person || "",
          partner: proj.partner || "",
          partner_person: proj.partner_person || "",
          work_content: dd?.night_work || "",
          work_time: "夜",
          priority: pri,
          is_priority: pri === "有",
          priority_detail: dd?.night_priority_detail || "",
        });
      }

      d.setUTCDate(d.getUTCDate() + 1);
    }
  }

  return items;
}

function sortItems(items) {
  return items.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    if (a.is_priority !== b.is_priority) return a.is_priority ? -1 : 1;
    return a.title < b.title ? -1 : a.title > b.title ? 1 : 0;
  });
}

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers: CORS_HEADERS, body: "" };
  }

  const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
  if (!GITHUB_TOKEN) {
    return {
      statusCode: 500,
      headers: CORS_HEADERS,
      body: JSON.stringify({ error: "サーバー設定エラー (GITHUB_TOKEN)" }),
    };
  }

  const params = event.queryStringParameters || {};

  try {
    // ===== GET ?signage=true =====
    if (event.httpMethod === "GET" && params.signage) {
      const months = getTargetMonths();
      let allItems = [];

      for (const m of months) {
        const d = await fetchMonthData(GITHUB_TOKEN, m);
        if (d) allItems.push(...expandToItems(d));
      }

      // 月またぎプロジェクトの重複排除
      const seen = new Set();
      allItems = allItems.filter((item) => {
        const key = `${item.date}|${item.client}|${item.title}|${item.work_time}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });

      sortItems(allItems);

      const jst = getJSTNow();
      const tmr = new Date(jst);
      tmr.setUTCDate(tmr.getUTCDate() + 1);

      return {
        statusCode: 200,
        headers: CORS_HEADERS,
        body: JSON.stringify({
          items: allItems,
          updated_at: new Date().toISOString(),
          today: fmtDate(jst),
          tomorrow: fmtDate(tmr),
        }),
      };
    }

    // ===== GET ?month=YYMM =====
    if (event.httpMethod === "GET" && params.month) {
      if (!/^\d{4}$/.test(params.month)) {
        return {
          statusCode: 400,
          headers: CORS_HEADERS,
          body: JSON.stringify({ error: "月はYYMM形式で指定してください" }),
        };
      }
      const d = await fetchMonthData(GITHUB_TOKEN, params.month);
      return {
        statusCode: 200,
        headers: CORS_HEADERS,
        body: JSON.stringify(d || { projects: [], daily: {} }),
      };
    }

    // ===== PUT =====
    if (event.httpMethod === "PUT") {
      let body;
      try {
        body = JSON.parse(event.body);
      } catch {
        return {
          statusCode: 400,
          headers: CORS_HEADERS,
          body: JSON.stringify({ error: "リクエスト形式が不正です" }),
        };
      }

      const { month, data, sha } = body;

      if (!month || !/^\d{4}$/.test(month)) {
        return {
          statusCode: 400,
          headers: CORS_HEADERS,
          body: JSON.stringify({ error: "月はYYMM形式で指定してください" }),
        };
      }
      if (!data) {
        return {
          statusCode: 400,
          headers: CORS_HEADERS,
          body: JSON.stringify({ error: "データが必要です" }),
        };
      }

      const saveData = { ...data };
      delete saveData._sha;

      const path = `_${month}.json`;
      const content = Buffer.from(
        JSON.stringify(saveData, null, 2),
        "utf-8"
      ).toString("base64");

      const headers = {
        ...ghHeaders(GITHUB_TOKEN),
        "Content-Type": "application/json",
      };

      // SHA取得（楽観的ロック）
      let currentSha = sha;
      if (!currentSha) {
        const url = `https://api.github.com/repos/${REPO}/contents/${path}?ref=${BRANCH}`;
        const getRes = await fetch(url, { headers });
        if (getRes.ok) {
          currentSha = (await getRes.json()).sha;
        }
      }

      const commitBody = {
        message: `Update ${month}.json`,
        content,
        branch: BRANCH,
      };
      if (currentSha) commitBody.sha = currentSha;

      const apiUrl = `https://api.github.com/repos/${REPO}/contents/${path}`;
      const putRes = await fetch(apiUrl, {
        method: "PUT",
        headers,
        body: JSON.stringify(commitBody),
      });

      if (putRes.ok) {
        const result = await putRes.json();
        return {
          statusCode: 200,
          headers: CORS_HEADERS,
          body: JSON.stringify({
            success: true,
            sha: result.content.sha,
            message: `${month}.json を保存しました`,
          }),
        };
      }

      if (putRes.status === 409) {
        return {
          statusCode: 409,
          headers: CORS_HEADERS,
          body: JSON.stringify({
            error:
              "データが他のユーザーによって更新されています。ページをリロードしてください。",
          }),
        };
      }

      const errText = await putRes.text();
      console.error("GitHub API error:", errText);
      return {
        statusCode: 500,
        headers: CORS_HEADERS,
        body: JSON.stringify({ error: `GitHub APIエラー: ${putRes.status}` }),
      };
    }

    return {
      statusCode: 405,
      headers: CORS_HEADERS,
      body: JSON.stringify({ error: "Method Not Allowed" }),
    };
  } catch (err) {
    console.error("Data function error:", err);
    return {
      statusCode: 500,
      headers: CORS_HEADERS,
      body: JSON.stringify({ error: `サーバーエラー: ${err.message}` }),
    };
  }
}
