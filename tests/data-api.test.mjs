import test from "node:test";
import assert from "node:assert/strict";
import {
  handler,
  mergeSources,
  expandToItems,
  validateData,
} from "../netlify/functions/data.mjs";

const project = (id, start = "2026-05-29", end = "2026-09-02") => ({
  id,
  title: "同名工事",
  client: "取引先",
  start_date: start,
  end_date: end,
});
test("historical same ID reconciles once, distinct IDs remain; owning month daily wins", () => {
  const out = mergeSources([
    {
      month: "2605",
      data: {
        projects: [project("p")],
        daily: { "p/2026-06-01": { day_work: "old" } },
      },
    },
    {
      month: "2606",
      data: {
        projects: [project("p"), project("q")],
        daily: { "p/2026-06-01": { day_work: "june" } },
      },
    },
    {
      month: "2607",
      data: {
        projects: [project("p")],
        daily: { "p/2026-06-01": { day_work: "copied" } },
      },
    },
  ]);
  assert.equal(out.projects.length, 2);
  assert.equal(out.daily["p/2026-06-01"].day_work, "june");
  assert.equal(expandToItems(out, "森", "2026-06-01", "2026-06-01").length, 2);
});
test("long duration expansion bounded to requested dates", () => {
  const out = expandToItems(
    { projects: [project("p", "2000-01-01", "2099-12-31")], daily: {} },
    "森",
    "2026-06-01",
    "2026-06-30",
  );
  assert.equal(out.length, 30);
});
test("invalid dates, duplicate IDs and orphan daily are rejected", () => {
  for (const data of [
    { projects: [project("p", "2026-02-30")], daily: {} },
    { projects: [project("p"), project("p")], daily: {} },
    { projects: [project("p")], daily: { "missing/2026-06-01": {} } },
  ])
    assert.throws(() => validateData(data));
});
test("invalid user and month reject before network", async () => {
  process.env.GITHUB_TOKEN = "test";
  for (const queryStringParameters of [
    { month: "2613", user: "森" },
    { month: "2606", user: "../bad" },
  ]) {
    assert.equal(
      (await handler({ httpMethod: "GET", queryStringParameters })).statusCode,
      400,
    );
  }
});

function fakeGit(files, options = {}) {
  const calls = [];
  const fetch = async (url, init = {}) => {
    const path = url.split("/construction-board/")[1],
      method = init.method || "GET",
      body = init.body && JSON.parse(init.body);
    calls.push({ path, method, body });
    const response = (value, status = 200) => ({
      ok: status >= 200 && status < 300,
      status,
      json: async () => value,
    });
    if (options.failPath && path.includes(options.failPath))
      return response({}, options.failStatus || 500);
    if (path === "git/ref/heads/main")
      return response({ object: { sha: "head1" } });
    if (path === "git/commits/head1")
      return response({ tree: { sha: "tree1" } });
    if (path === "git/trees/tree1")
      return response({
        tree: Object.entries(files).map(([path, f]) => ({
          path,
          sha: f.sha,
          type: "blob",
        })),
      });
    if (path.startsWith("git/blobs/") && method === "GET") {
      const file = Object.values(files).find(
        (f) => f.sha === path.split("/").pop(),
      );
      return response({
        encoding: "base64",
        content: Buffer.from(file.raw || JSON.stringify(file.data)).toString(
          "base64",
        ),
      });
    }
    if (path === "git/blobs" && method === "POST")
      return response({ sha: "saved-blob" });
    if (path === "git/trees" && method === "POST")
      return response({ sha: "saved-tree" });
    if (path === "git/commits" && method === "POST")
      return response({ sha: "saved-commit" });
    if (path === "git/refs/heads/main" && method === "PATCH")
      return response({}, options.conflict ? 422 : 200);
    throw new Error("Unexpected mock request " + path);
  };
  return { fetch, calls };
}
const getAll = () =>
  handler({
    httpMethod: "GET",
    queryStringParameters: { month: "2609", user: "森", scope: "all" },
  });
const put = (revision, data, extra = {}) =>
  handler({
    httpMethod: "PUT",
    body: JSON.stringify({
      month: "2609",
      user: "森",
      scope: "all",
      revision,
      data,
      ...extra,
    }),
  });
async function withGit(files, options, fn) {
  const old = globalThis.fetch,
    mock = fakeGit(files, options);
  globalThis.fetch = mock.fetch;
  process.env.GITHUB_TOKEN = "test";
  try {
    await fn(mock);
  } finally {
    globalThis.fetch = old;
  }
}
const oldFile = () => ({
  "_2605_森.json": {
    sha: "may",
    data: { projects: [project("p")], daily: {} },
  },
});
test("all-history GET finds May long project in September and never writes", async () =>
  withGit(oldFile(), {}, async ({ calls }) => {
    const res = await getAll(),
      data = JSON.parse(res.body);
    assert.equal(res.statusCode, 200);
    assert.equal(data.projects.length, 1);
    assert.match(data._revision, /^v2:/);
    assert.ok(calls.every((c) => c.method === "GET"));
    assert.ok(calls.some((c) => c.path === "git/blobs/may"));
  }));
test("canonical wins over legacy, legacy monthly read filters dates and old PUT rejects", async () => {
  const files = {
    ...oldFile(),
    "_projects_森.json": {
      sha: "canonical",
      data: {
        projects: [project("new", "2026-09-01", "2026-10-02")],
        daily: { "new/2026-10-01": { night: true } },
      },
    },
  };
  await withGit(files, {}, async ({ calls }) => {
    const data = JSON.parse((await getAll()).body);
    assert.deepEqual(
      data.projects.map((p) => p.id),
      ["new"],
    );
    const monthly = await handler({
      httpMethod: "GET",
      queryStringParameters: { month: "2609", user: "森" },
    });
    assert.deepEqual(JSON.parse(monthly.body).daily, {});
    assert.equal(
      (
        await put(undefined, files["_projects_森.json"].data, {
          scope: undefined,
          sha: "canonical",
        })
      ).statusCode,
      409,
    );
    assert.ok(
      !calls.some((c) => c.path === "git/blobs/may" || c.method !== "GET"),
    );
  });
});
test("migration saves canonical with CAS and preserves every legacy source", async () =>
  withGit(oldFile(), {}, async ({ calls }) => {
    const data = JSON.parse((await getAll()).body),
      res = await put(data._revision, data);
    assert.equal(res.statusCode, 200);
    assert.match(JSON.parse(res.body).revision, /^v2:/);
    assert.deepEqual(calls.find((c) => c.path === "git/trees").body, {
      base_tree: "tree1",
      tree: [
        {
          path: "_projects_森.json",
          mode: "100644",
          type: "blob",
          sha: "saved-blob",
        },
      ],
    });
    assert.deepEqual(calls.find((c) => c.path === "git/commits").body.parents, [
      "head1",
    ]);
    assert.equal(calls.find((c) => c.method === "PATCH").body.force, false);
  }));
test("stale or missing revision cannot save", async () =>
  withGit(oldFile(), {}, async ({ calls }) => {
    assert.equal(
      (await put("stale", oldFile()["_2605_森.json"].data)).statusCode,
      409,
    );
    assert.equal(
      (await put(undefined, oldFile()["_2605_森.json"].data)).statusCode,
      409,
    );
    assert.ok(calls.every((c) => c.method === "GET"));
  }));
test("head change during write returns conflict", async () =>
  withGit(oldFile(), { conflict: true }, async () => {
    const data = JSON.parse((await getAll()).body);
    assert.equal((await put(data._revision, data)).statusCode, 409);
  }));
test("partial read failure never returns empty success, retries bounded", async () =>
  withGit(oldFile(), { failPath: "git/blobs/may" }, async ({ calls }) => {
    assert.equal((await getAll()).statusCode, 502);
    assert.equal(calls.filter((c) => c.path === "git/blobs/may").length, 3);
    assert.ok(calls.every((c) => c.method === "GET"));
  }));
test("corrupt source blocks read", async () =>
  withGit({ "_2605_森.json": { sha: "bad", raw: "{broken" } }, {}, async () => {
    assert.equal((await getAll()).statusCode, 502);
  }));
test("signage reads historical source and keeps distinct same-name IDs", async () => {
  const now = new Date(Date.now() + 9 * 3600000),
    date = now.toISOString().slice(0, 10);
  await withGit(
    {
      "_2501_森.json": {
        sha: "old",
        data: {
          projects: [
            project("p", "2025-01-01", "2099-12-31"),
            project("q", date, date),
          ],
          daily: {},
        },
      },
    },
    {},
    async () => {
      const res = await handler({
          httpMethod: "GET",
          queryStringParameters: { signage: "true" },
        }),
        data = JSON.parse(res.body);
      assert.equal(res.statusCode, 200);
      assert.equal(data.items.filter((i) => i.date === date).length, 2);
      assert.ok(data.items.length <= 93);
    },
  );
});
test("429 is retried then reported, never silently empty", async () =>
  withGit(
    oldFile(),
    { failPath: "git/blobs/may", failStatus: 429 },
    async ({ calls }) => {
      assert.equal((await getAll()).statusCode, 502);
      assert.equal(calls.filter((c) => c.path === "git/blobs/may").length, 3);
    },
  ));
test("network timeout is bounded and surfaces 504", async () => {
  const old = globalThis.fetch;
  let count = 0;
  globalThis.fetch = async () => {
    count++;
    throw new DOMException("timeout", "TimeoutError");
  };
  try {
    assert.equal((await getAll()).statusCode, 504);
    assert.equal(count, 3);
  } finally {
    globalThis.fetch = old;
  }
});
test("malformed GitHub JSON surfaces 502", async () => {
  const old = globalThis.fetch;
  globalThis.fetch = async () => ({
    ok: true,
    status: 200,
    json: async () => {
      throw new SyntaxError();
    },
  });
  try {
    assert.equal((await getAll()).statusCode, 502);
  } finally {
    globalThis.fetch = old;
  }
});
test("large history blob reads use at most six concurrent requests", async () => {
  const files = Object.fromEntries(
    Array.from({ length: 60 }, (_, i) => [
      `_${String(2000 + Math.floor(i / 12)).slice(-2)}${String((i % 12) + 1).padStart(2, "0")}_森.json`,
      { sha: "blob" + i, data: { projects: [project("p" + i)], daily: {} } },
    ]),
  );
  await withGit(files, {}, async (mock) => {
    let active = 0,
      peak = 0;
    const fetch = mock.fetch;
    globalThis.fetch = async (url, init) => {
      if (url.includes("/git/blobs/")) {
        active++;
        peak = Math.max(peak, active);
        await new Promise((r) => setTimeout(r, 1));
        const result = await fetch(url, init);
        active--;
        return result;
      }
      return fetch(url, init);
    };
    const res = await getAll();
    assert.equal(res.statusCode, 200);
    assert.equal(JSON.parse(res.body).projects.length, 60);
    assert.equal(peak, 6);
  });
});
test("obsolete combined fallback is excluded before fetching invalid content", async () =>
  withGit(
    { ...oldFile(), "_2605.json": { sha: "obsolete", raw: "bad JSON" } },
    {},
    async ({ calls }) => {
      assert.equal(
        (
          await handler({
            httpMethod: "GET",
            queryStringParameters: { signage: "true" },
          })
        ).statusCode,
        200,
      );
      assert.ok(!calls.some((c) => c.path === "git/blobs/obsolete"));
    },
  ));
test("output expansion has bounded size and fails explicitly", () => {
  const projects = Array.from({ length: 150 }, (_, i) =>
    project("p" + i, "2026-01-01", "2026-12-31"),
  );
  assert.throws(
    () =>
      expandToItems({ projects, daily: {} }, "森", "2026-06-01", "2026-08-31"),
    (e) => e.status === 413,
  );
});
test("request above safe buffered payload ceiling rejected before GitHub", async () =>
  withGit({}, {}, async ({ calls }) => {
    const res = await handler({
      httpMethod: "PUT",
      body: "x".repeat(4 * 1024 * 1024 + 1),
    });
    assert.equal(res.statusCode, 413);
    assert.equal(calls.length, 0);
  }));

test("legacy edit between load and migration invalidates revision", async () => {
  const files = oldFile();
  await withGit(files, {}, async ({ calls }) => {
    const data = JSON.parse((await getAll()).body);
    files["_2605_森.json"].sha = "changed-by-old-client";
    assert.equal((await put(data._revision, data)).statusCode, 409);
    assert.ok(calls.every((call) => call.method === "GET"));
  });
});

test("unrelated user change does not invalidate this user's revision", async () => {
  const files = oldFile();
  await withGit(files, {}, async () => {
    const data = JSON.parse((await getAll()).body);
    files["_2609_谷口.json"] = {
      sha: "unrelated",
      data: { projects: [], daily: {} },
    };
    assert.equal((await put(data._revision, data)).statusCode, 200);
  });
});

test("legacy monthly GET returns only original file and original SHA", async () => {
  const files = oldFile();
  files["_2606_森.json"] = {
    sha: "june",
    data: { projects: [project("q")], daily: {} },
  };
  await withGit(files, {}, async () => {
    const res = await handler({
      httpMethod: "GET",
      queryStringParameters: { month: "2605", user: "森" },
    });
    const data = JSON.parse(res.body);
    assert.equal(data._sha, "may");
    assert.deepEqual(
      data.projects.map((p) => p.id),
      ["p"],
    );
  });
});
