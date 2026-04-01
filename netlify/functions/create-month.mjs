/**
 * 月別ファイル生成 Netlify Function
 * GitHub Actions workflow_dispatch をトリガーして .xlsx ファイルを生成する
 */

const REPO = "Tiger0710/construction-board";
const WORKFLOW = "create-month.yml";

export async function handler(event) {
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
    return { statusCode: 500, headers, body: JSON.stringify({ error: "GITHUB_TOKEN未設定" }) };
  }

  let month, persons;
  try {
    const body = JSON.parse(event.body);
    month = body.month;
    persons = body.persons;
  } catch {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "リクエスト形式が不正です" }) };
  }

  // Validate
  if (!month || !month.match(/^\d{4}$/)) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "月はYYMM形式で指定してください" }) };
  }
  if (!Array.isArray(persons) || persons.length === 0 || persons.length > 10) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "担当者名は1〜10名で指定してください" }) };
  }

  const personsStr = persons.join(",");

  try {
    const res = await fetch(
      `https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW}/dispatches`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${GITHUB_TOKEN}`,
          Accept: "application/vnd.github.v3+json",
          "User-Agent": "construction-board-admin",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          ref: "main",
          inputs: { month, persons: personsStr },
        }),
      }
    );

    if (res.status === 204 || res.ok) {
      return {
        statusCode: 200,
        headers,
        body: JSON.stringify({
          success: true,
          message: `${month}月のファイル生成を開始しました (${personsStr})`,
        }),
      };
    } else {
      const err = await res.text();
      console.error("GitHub API error:", res.status, err);
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ error: `GitHub APIエラー: ${res.status}` }),
      };
    }
  } catch (err) {
    console.error("Error:", err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: `サーバーエラー: ${err.message}` }),
    };
  }
}
