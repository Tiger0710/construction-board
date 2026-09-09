/** Month-independent project API; legacy sources remain migration backups. */
import { createHash } from "node:crypto";
const REPO = "Tiger0710/construction-board",
  BRANCH = process.env.DATA_BRANCH || "main";
const HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
  "Content-Type": "application/json",
  "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
};
class ApiError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}
const fail = (status, message) => {
  throw new ApiError(status, message);
};
// Netlify buffered payload limit is 6 MB; keep 2 MB margin for the response envelope.
const MAX_BYTES = 4 * 1024 * 1024,
  MAX_ITEMS = 12000;
const answer = (statusCode, data) => {
  const body = JSON.stringify(data);
  if (Buffer.byteLength(body) > MAX_BYTES)
    return {
      statusCode: 413,
      headers: HEADERS,
      body: JSON.stringify({
        error: "表示データが上限を超えています。管理者に連絡してください。",
      }),
    };
  return { statusCode, headers: HEADERS, body };
};
const object = (v) => v !== null && typeof v === "object" && !Array.isArray(v);
const validMonth = (m) =>
  typeof m === "string" && /^\d{2}(0[1-9]|1[0-2])$/.test(m);
const validUser = (u) =>
  typeof u === "string" &&
  u.length > 0 &&
  u.length <= 80 &&
  !/[\x00-\x1f\x7f/\\<>:"|?*]/.test(u) &&
  !u.includes("..") &&
  u.trim() === u;
const validDate = (s) =>
  typeof s === "string" &&
  /^20\d{2}-\d{2}-\d{2}$/.test(s) &&
  !isNaN(Date.parse(s)) &&
  new Date(s).toISOString().slice(0, 10) === s;
const dateString = (d) => d.toISOString().slice(0, 10);

export function validateData(data) {
  if (!object(data) || !Array.isArray(data.projects) || !object(data.daily))
    fail(400, "工事データの形式が不正です。");
  if (data.projects.length > 20000 || Object.keys(data.daily).length > 400000)
    fail(413, "データ件数が上限を超えています。");
  const ids = new Set();
  for (const p of data.projects) {
    if (
      !object(p) ||
      typeof p.id !== "string" ||
      !p.id ||
      p.id.length > 160 ||
      /[\/\x00-\x1f\x7f]/.test(p.id) ||
      ids.has(p.id)
    )
      fail(400, "工事IDが不正または重複しています。");
    ids.add(p.id);
    if (
      p.default_shift !== undefined &&
      p.default_shift !== "day" &&
      p.default_shift !== "night"
    )
      fail(400, "工事の基本時間帯は昼または夜を指定してください。");
    if (
      !validDate(p.start_date) ||
      !validDate(p.end_date) ||
      p.start_date > p.end_date
    )
      fail(400, "工事の開始日・終了日が不正です。");
    for (const value of Object.values(p))
      if (
        typeof value !== "string" &&
        value !== null &&
        typeof value !== "number" &&
        typeof value !== "boolean"
      )
        fail(400, "工事項目の形式が不正です。");
    for (const field of [
      "title",
      "client",
      "our_person",
      "safety_person",
      "partner",
      "partner_person",
    ])
      if (
        p[field] !== undefined &&
        p[field] !== null &&
        typeof p[field] !== "string"
      )
        fail(400, "工事名・担当者などの項目は文字列で指定してください。");
  }
  for (const [key, value] of Object.entries(data.daily)) {
    const slash = key.lastIndexOf("/");
    if (
      !ids.has(key.slice(0, slash)) ||
      !validDate(key.slice(slash + 1)) ||
      !object(value)
    )
      fail(400, "日別データの工事ID・日付が不正です。");
    for (const [field, v] of Object.entries(value)) {
      if (field === "day" || field === "night") {
        if (typeof v !== "boolean") fail(400, "昼夜の指定が不正です。");
      } else if (typeof v !== "string" && v !== null)
        fail(400, "日別データの値が不正です。");
    }
  }
  return { projects: data.projects, daily: data.daily };
}

export function mergeSources(sources) {
  const projects = new Map(),
    daily = Object.create(null),
    owners = new Map(),
    warnings = new Set();
  for (const source of [...sources].sort((a, b) =>
    (a.month || "").localeCompare(b.month || ""),
  )) {
    validateData(source.data);
    for (const p of source.data.projects) {
      const old = projects.get(p.id);
      if (old) {
        if (JSON.stringify(old) !== JSON.stringify(p))
          warnings.add(
            `工事「${p.title || p.id}」は月別登録に差があります。期間を統合し、基本情報は新しい月を採用しました。元ファイルは保持されます。`,
          );
        projects.set(p.id, {
          ...old,
          ...p,
          start_date:
            old.start_date < p.start_date ? old.start_date : p.start_date,
          end_date: old.end_date > p.end_date ? old.end_date : p.end_date,
        });
      } else projects.set(p.id, { ...p });
    }
    for (const [key, value] of Object.entries(source.data.daily)) {
      const date = key.slice(key.lastIndexOf("/") + 1),
        own = date.slice(2, 4) + date.slice(5, 7),
        previous = owners.get(key);
      if (previous && JSON.stringify(daily[key]) !== JSON.stringify(value))
        warnings.add(
          `日別設定 ${date}（工事ID: ${key.slice(0, key.lastIndexOf("/"))}）に月別の差があります。該当月の設定を優先し、元ファイルは保持されます。`,
        );
      if (!previous || previous !== own || source.month === own) {
        daily[key] = value;
        owners.set(key, source.month);
      }
    }
  }
  return { projects: [...projects.values()], daily, _warnings: [...warnings] };
}

export function expandToItems(
  data,
  user,
  windowStart,
  windowEnd,
  budget = { count: 0, bytes: 0 },
) {
  const items = [];
  for (const p of data.projects) {
    const start = p.start_date > windowStart ? p.start_date : windowStart,
      end = p.end_date < windowEnd ? p.end_date : windowEnd;
    for (
      let d = new Date(start + "T00:00:00Z");
      dateString(d) <= end;
      d.setUTCDate(d.getUTCDate() + 1)
    ) {
      const date = dateString(d),
        dd = data.daily[`${p.id}/${date}`];
      for (const [shift, label, active] of [
        ["day", "昼", dd ? dd.day !== false : p.default_shift !== "night"],
        ["night", "夜", dd ? dd.night === true : p.default_shift === "night"],
      ]) {
        if (!active) continue;
        const priority = dd?.[`${shift}_priority`] || "";
        const item = {
          id: p.id,
          project_id: p.id,
          user,
          date,
          client: p.client,
          title: p.title,
          our_person: dd?.[`${shift}_our_person`] || p.our_person || "",
          safety_person:
            dd?.[`${shift}_safety_person`] || p.safety_person || "",
          partner: p.partner || "",
          partner_person:
            dd?.[`${shift}_partner_person`] || p.partner_person || "",
          work_content: dd?.[`${shift}_work`] || "",
          work_time: label,
          priority,
          is_priority: priority === "有",
          priority_detail: dd?.[`${shift}_priority_detail`] || "",
        };
        budget.count++;
        budget.bytes += Buffer.byteLength(JSON.stringify(item)) + 1;
        if (budget.count > MAX_ITEMS || budget.bytes > MAX_BYTES - 65536)
          fail(
            413,
            "表示データが上限を超えています。管理者に連絡してください。",
          );
        items.push(item);
      }
    }
  }
  return items;
}

function api(token) {
  const deadline = Date.now() + 23000;
  return async function request(path, { method = "GET", body } = {}) {
    for (let attempt = 0; attempt < 3; attempt++) {
      const remaining = deadline - Date.now();
      if (remaining <= 0)
        fail(504, "データ取得がタイムアウトしました。再読み込みしてください。");
      let res;
      try {
        res = await fetch(`https://api.github.com/repos/${REPO}/${path}`, {
          method,
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/vnd.github+json",
            "User-Agent": "construction-board-data",
            "Content-Type": "application/json",
          },
          body: body ? JSON.stringify(body) : undefined,
          signal: AbortSignal.timeout(Math.min(7000, remaining)),
        });
      } catch {
        if (method === "GET" && attempt < 2) continue;
        fail(504, "データ通信がタイムアウトしました。再読み込みしてください。");
      }
      if (
        (res.status === 429 || res.status >= 500) &&
        method === "GET" &&
        attempt < 2
      ) {
        await new Promise((resolve) =>
          setTimeout(resolve, 150 * (attempt + 1)),
        );
        continue;
      }
      if (!res.ok)
        fail(
          res.status === 409 || res.status === 422 ? 409 : 502,
          res.status === 409 || res.status === 422
            ? "別の更新がありました。再読み込みして内容を確認してください。"
            : `データ保存先に接続できません（${res.status}）。再読み込みしてください。`,
        );
      try {
        return await res.json();
      } catch {
        fail(502, "データ保存先から不正な応答がありました。");
      }
    }
  };
}
async function snapshot(request) {
  const ref = await request(`git/ref/heads/${encodeURIComponent(BRANCH)}`),
    head = ref?.object?.sha;
  if (!head) fail(502, "保存先のリビジョンを取得できません。");
  const commit = await request(`git/commits/${head}`);
  const tree = await request(`git/trees/${commit.tree.sha}`);
  if (tree.truncated || !Array.isArray(tree.tree))
    fail(502, "データ一覧を完全に取得できません。");
  return {
    head,
    tree: commit.tree.sha,
    files: tree.tree.filter((f) => f.type === "blob"),
  };
}
function classify(file) {
  let m = /^_projects_(.+)\.json$/.exec(file.path);
  if (m) return { ...file, user: m[1], canonical: true };
  m = /^_(\d{4})(?:_(.+))?\.json$/.exec(file.path);
  return m && validMonth(m[1])
    ? { ...file, user: m[2] || "", month: m[1], canonical: false }
    : null;
}
function sourceFiles(snap, user) {
  const files = snap.files.map(classify).filter((f) => f && f.user === user),
    canonical = files.find((f) => f.canonical);
  return canonical ? [canonical] : files;
}
const revision = (files) =>
  "v2:" +
  createHash("sha256")
    .update(
      JSON.stringify(
        files
          .map((f) => [f.path, f.sha])
          .sort((a, b) => a[0].localeCompare(b[0])),
      ),
    )
    .digest("hex");
async function parallelMap(values, fn) {
  const results = new Array(values.length);
  let next = 0;
  await Promise.all(
    Array.from({ length: Math.min(6, values.length) }, async () => {
      while (next < values.length) {
        const i = next++;
        results[i] = await fn(values[i]);
      }
    }),
  );
  return results;
}
async function readSources(request, files) {
  if (files.length > 512)
    fail(413, "履歴ファイル数が上限を超えています。管理者に連絡してください。");
  let totalBytes = 0;
  return parallelMap(files, async (f) => {
    const blob = await request(`git/blobs/${f.sha}`);
    let data;
    totalBytes +=
      typeof blob.content === "string" ? Buffer.byteLength(blob.content) : 0;
    if (totalBytes > 32 * 1024 * 1024)
      fail(413, "履歴データが上限を超えています。管理者に連絡してください。");
    try {
      if (blob.encoding !== "base64" || typeof blob.content !== "string")
        throw new Error();
      data = JSON.parse(Buffer.from(blob.content, "base64").toString("utf8"));
      validateData(data);
    } catch {
      fail(
        502,
        `保存済みデータ ${f.path} を読み取れません。保存を中止しました。`,
      );
    }
    return { ...f, data };
  });
}
async function loadUser(request, snap, user) {
  const files = sourceFiles(snap, user),
    sources = await readSources(request, files);
  const data = files[0]?.canonical
    ? { ...sources[0].data, _warnings: [] }
    : mergeSources(sources);
  return {
    ...data,
    _revision: revision(files),
    _sha: files.length === 1 ? files[0].sha : undefined,
  };
}
function monthly(data, month) {
  const start = `20${month.slice(0, 2)}-${month.slice(2)}-01`,
    end = dateString(
      new Date(
        Date.UTC(Number("20" + month.slice(0, 2)), Number(month.slice(2)), 0),
      ),
    );
  const projects = data.projects.filter(
      (p) => p.start_date <= end && p.end_date >= start,
    ),
    ids = new Set(projects.map((p) => p.id));
  return {
    ...data,
    projects,
    daily: Object.fromEntries(
      Object.entries(data.daily).filter(
        ([k]) =>
          ids.has(k.slice(0, k.lastIndexOf("/"))) &&
          k.slice(-10) >= start &&
          k.slice(-10) <= end,
      ),
    ),
  };
}
async function writeAtomic(request, snap, path, data, message) {
  const blob = await request("git/blobs", {
    method: "POST",
    body: { content: JSON.stringify(data, null, 2), encoding: "utf-8" },
  });
  const tree = await request("git/trees", {
    method: "POST",
    body: {
      base_tree: snap.tree,
      tree: [{ path, mode: "100644", type: "blob", sha: blob.sha }],
    },
  });
  const commit = await request("git/commits", {
    method: "POST",
    body: { message, tree: tree.sha, parents: [snap.head] },
  });
  await request(`git/refs/heads/${encodeURIComponent(BRANCH)}`, {
    method: "PATCH",
    body: { sha: commit.sha, force: false },
  });
  return blob.sha;
}

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") return answer(200, {});
  if (!process.env.GITHUB_TOKEN)
    return answer(500, { error: "サーバーのデータ接続設定がありません。" });
  try {
    const params = event.queryStringParameters || {},
      request = api(process.env.GITHUB_TOKEN);
    if (event.httpMethod === "GET" && params.signage) {
      const snap = await snapshot(request),
        all = snap.files.map(classify).filter(Boolean),
        users = [...new Set(all.map((f) => f.user))];
      const selected = users
        .flatMap((user) => sourceFiles(snap, user))
        .filter(
          (f) =>
            f.user !== "" ||
            !all.some(
              (a) => a.user !== "" && !a.canonical && a.month === f.month,
            ),
        );
      const sources = await readSources(request, selected);
      const now = new Date(Date.now() + 9 * 3600000),
        year = now.getUTCFullYear(),
        month = now.getUTCMonth();
      const start = dateString(new Date(Date.UTC(year, month - 1, 1))),
        end = dateString(new Date(Date.UTC(year, month + 2, 0)));
      const items = [],
        warnings = [],
        budget = { count: 0, bytes: 0 };
      for (const user of users) {
        const group = sources.filter((f) => f.user === user);
        const merged = group[0]?.canonical
          ? group[0].data
          : mergeSources(group);
        warnings.push(...(merged._warnings || []));
        items.push(...expandToItems(merged, user, start, end, budget));
      }
      items.sort(
        (a, b) =>
          a.date.localeCompare(b.date) ||
          Number(b.is_priority) - Number(a.is_priority) ||
          (a.title || "").localeCompare(b.title || ""),
      );
      const tomorrow = new Date(now);
      tomorrow.setUTCDate(tomorrow.getUTCDate() + 1);
      return answer(200, {
        items,
        updated_at: new Date().toISOString(),
        today: dateString(now),
        tomorrow: dateString(tomorrow),
        _warnings: warnings,
      });
    }
    if (event.httpMethod === "GET" && params.month) {
      if (!validMonth(params.month))
        fail(400, "月は有効なYYMM形式で指定してください。");
      if (params.user !== undefined && !validUser(params.user))
        fail(400, "担当者名が不正です。");
      const snap = await snapshot(request);
      if (params.user) {
        if (
          params.scope !== "all" &&
          !sourceFiles(snap, params.user).some((f) => f.canonical)
        ) {
          const files = sourceFiles(snap, params.user).filter(
              (f) => f.month === params.month,
            ),
            sources = await readSources(request, files);
          return answer(
            200,
            sources.length
              ? { ...sources[0].data, _sha: files[0].sha }
              : { projects: [], daily: {} },
          );
        }
        const data = await loadUser(request, snap, params.user);
        return answer(
          200,
          params.scope === "all" ? data : monthly(data, params.month),
        );
      }
      return answer(200, {
        members: [
          ...new Set(
            snap.files
              .map(classify)
              .filter((f) => f && f.user)
              .map((f) => f.user),
          ),
        ],
      });
    }
    if (event.httpMethod === "PUT") {
      if (
        typeof event.body !== "string" ||
        Buffer.byteLength(event.body) > MAX_BYTES
      )
        fail(413, "保存データが大きすぎます。");
      let body;
      try {
        body = JSON.parse(event.body);
      } catch {
        fail(400, "リクエストの形式が不正です。");
      }
      if (!object(body) || !validMonth(body.month) || !validUser(body.user))
        fail(400, "月または担当者名が不正です。");
      const data = validateData(body.data),
        snap = await snapshot(request),
        files = sourceFiles(snap, body.user);
      let path;
      if (body.scope === "all") {
        if (body.revision !== revision(files))
          fail(
            409,
            "別の更新がありました。再読み込みして内容を確認してください。",
          );
        path = `_projects_${body.user}.json`;
      } else {
        if (files.some((f) => f.canonical))
          fail(
            409,
            "入力画面が更新されています。ページを再読み込みしてください。",
          );
        path = `_${body.month}_${body.user}.json`;
        const file = files.find((f) => f.path === path);
        if ((body.sha || null) !== (file?.sha || null))
          fail(409, "別の更新がありました。再読み込みしてください。");
      }
      const sha = await writeAtomic(
        request,
        snap,
        path,
        data,
        `Update projects ${body.user}`,
      );
      return answer(200, {
        success: true,
        sha,
        revision: revision([{ path, sha }]),
      });
    }
    return answer(405, { error: "Method Not Allowed" });
  } catch (error) {
    return answer(error.status || 500, {
      error: error.status
        ? error.message
        : "データ処理でエラーが発生しました。再読み込みしてください。",
    });
  }
}
