/**
 * 工事予定表 ファイル一覧 + ダウンロード Netlify Function
 *
 * GET /.netlify/functions/files?month=2604
 *   → data/{month}/ 内の Excel ファイル一覧を JSON 返却
 *
 * GET /.netlify/functions/files?month=2604&download=入力_担当者A.xlsm
 *   → ファイルの base64 コンテンツを返却
 */

const REPO = "Tiger0710/construction-board";
const BRANCH = "main";

export async function handler(event) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  if (event.httpMethod !== "GET") {
    return { statusCode: 405, headers, body: JSON.stringify({ error: "Method Not Allowed" }) };
  }

  const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
  if (!GITHUB_TOKEN) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: "サーバー設定エラー (GITHUB_TOKEN)" }) };
  }

  const authHeaders = {
    Authorization: `Bearer ${GITHUB_TOKEN}`,
    Accept: "application/vnd.github.v3+json",
    "User-Agent": "construction-board-files",
  };

  const params = event.queryStringParameters || {};
  const month = params.month;
  const download = params.download;

  try {
    if (download) {
      // Download a specific file
      const dir = month ? `data/${month}` : "data";
      const path = `${dir}/${download}`;
      const url = `https://api.github.com/repos/${REPO}/contents/${path.split("/").map(encodeURIComponent).join("/")}?ref=${BRANCH}`;
      const res = await fetch(url, { headers: authHeaders });

      if (!res.ok) {
        return {
          statusCode: 404,
          headers,
          body: JSON.stringify({ error: "ファイルが見つかりません" }),
        };
      }

      const file = await res.json();
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          filename: file.name,
          content: file.content,  // base64 from GitHub API
          size: file.size,
        }),
      };
    }

    // List files in data/ or data/{month}/
    const dir = month ? `data/${month}` : "data";
    const url = `https://api.github.com/repos/${REPO}/contents/${dir.split("/").map(encodeURIComponent).join("/")}?ref=${BRANCH}`;
    const res = await fetch(url, { headers: authHeaders });

    if (!res.ok) {
      // Directory doesn't exist yet — return empty list
      if (res.status === 404) {
        return { statusCode: 200, headers, body: JSON.stringify({ files: [], month: month || null }) };
      }
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ error: `GitHub APIエラー: ${res.status}` }),
      };
    }

    const items = await res.json();
    const files = (Array.isArray(items) ? items : [])
      .filter((f) => f.type === "file" && /^入力_.+\.(xlsm|xlsx)$/.test(f.name))
      .map((f) => ({
        name: f.name,
        size: f.size,
        sha: f.sha,
      }));

    // Also list month subdirectories if no month specified
    let months = [];
    if (!month) {
      months = (Array.isArray(items) ? items : [])
        .filter((f) => f.type === "dir" && /^\d{4}$/.test(f.name))
        .map((f) => f.name)
        .sort()
        .reverse();
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ files, months, month: month || null }),
    };
  } catch (err) {
    console.error("Files error:", err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: `サーバーエラー: ${err.message}` }),
    };
  }
}
