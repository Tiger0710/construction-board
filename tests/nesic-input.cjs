const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..');

(async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const page = await browser.newPage();
    const errors = [];
    const reads = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname !== 'board.test') return route.abort();
      if (url.pathname.startsWith('/.netlify/functions/')) {
        assert.equal(route.request().method(), 'GET');
        if (url.searchParams.has('signage')) return route.fulfill({ json: { items: [] } });
        if (url.searchParams.has('user')) {
          reads.push(url.searchParams.get('user'));
          return route.fulfill({ json: { projects: [], daily: {}, _revision: 'empty' } });
        }
        return route.fulfill({ json: { members: ['手島'] } });
      }
      const file = path.join(root, 'static', url.pathname.slice(1));
      return route.fulfill({ body: fs.readFileSync(file), contentType:
        file.endsWith('.html') ? 'text/html; charset=utf-8' : 'application/json' });
    });

    await page.goto('http://board.test/input.html');
    await page.getByRole('button', { name: /NESIC/ }).click();
    await page.waitForFunction(() => state.ready && !state.loading && state.user === 'NESIC');
    assert.equal(await page.evaluate(() => companyOfMember(state.user).id), 'hsj');
    assert.equal(await page.evaluate(() => state.data.projects.length), 0);
    assert.ok(reads.includes('NESIC'));
    assert.match(page.url(), /#NESIC\//);
    assert.deepEqual(errors, []);
    console.log('PASS NESIC appears under HSJ, opens its own empty input page, and reads its own data');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
