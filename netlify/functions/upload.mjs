/**
 * 工事予定表 ファイルアップロード Netlify Function
 *
 * 受信した .xlsm ファイルを GitHub リポジトリの data/ に commit する。
 * commit により GitHub Actions が走り、自動で HTML 再生成 → Netlify デプロイ。
 */

const REPO = "Tiger0710/construction-board";
const BRANCH = "main";

export async function handler(event) {
  // CORS
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers, body: JSON.stringify({ error: "Method Not Allowed" }) };
  }

  const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
  if (!GITHUB_TOKEN) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: "サーバー設定エラー (GITHUB_TOKEN)" }) };
  }

  let filename, content, month;
  try {
    const body = JSON.parse(event.body);
    filename = body.filename;
    content = body.content; // base64
    month = body.month;     // YYMM (optional)
  } catch {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "リクエスト形式が不正です" }) };
  }

  // Validate month if provided
  if (month && !month.match(/^\d{4}$/)) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "月はYYMM形式で指定してください" }) };
  }

  // Validate filename
  if (!filename || !filename.match(/^入力_.+\.(xlsm|xlsx)$/)) {
    return {
      statusCode: 400,
      headers,
      body: JSON.stringify({ error: "ファイル名は「入力_担当者名.xlsm」の形式にしてください" }),
    };
  }

  // Validate size (base64 → ~75% of original, so 14MB base64 ≈ 10MB file)
  if (content && content.length > 14 * 1024 * 1024) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "ファイルサイズが大きすぎます" }) };
  }

  const path = month ? `data/${month}/${filename}` : `data/${filename}`;
  const encodedPath = path.split("/").map(encodeURIComponent).join("/");
  const apiBase = `https://api.github.com/repos/${REPO}/contents/${encodedPath}`;
  const authHeaders = {
    Authorization: `Bearer ${GITHUB_TOKEN}`,
    Accept: "application/vnd.github.v3+json",
    "User-Agent": "construction-board-uploader",
  };

  try {
    // Get existing file SHA (needed for update)
    let sha;
    const getRes = await fetch(`${apiBase}?ref=${BRANCH}`, { headers: authHeaders });
    if (getRes.ok) {
      const existing = await getRes.json();
      sha = existing.sha;
    }

    // Create or update file
    const commitBody = {
      message: `Update ${filename}`,
      content: content,
      branch: BRANCH,
    };
    if (sha) commitBody.sha = sha;

    const putRes = await fetch(apiBase, {
      method: "PUT",
      headers: { ...authHeaders, "Content-Type": "application/json" },
      body: JSON.stringify(commitBody),
    });

    if (putRes.ok) {
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({ success: true, message: `${filename} をアップロードしました` }),
      };
    } else {
      const err = await putRes.text();
      console.error("GitHub API error:", err);
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ error: `GitHub APIエラー: ${putRes.status}` }),
      };
    }
  } catch (err) {
    console.error("Upload error:", err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: `サーバーエラー: ${err.message}` }),
    };
  }
}
