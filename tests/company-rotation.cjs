// Run: node tests/company-rotation.cjs (requires playwright and Chrome)
const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const org = JSON.parse(fs.readFileSync(path.join(root, 'static/org.json'), 'utf8'));

(async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  let passed = 0;
  const check = (name, actual, expected) => {
    assert.deepEqual(actual, expected, name);
    console.log('PASS ' + name); passed++;
  };
  try {
    const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.clock.install({ time: new Date('2026-09-09T12:00:00+09:00') });
    const today = '2026-09-09', tomorrow = '2026-09-10';
    const item = (user, date, n = 0) => ({ user, date, title: '通信設備更新工事 ' + n,
      client: 'テスト客先', our_person: user, safety_person: '安全担当', partner: '協力会社',
      partner_person: '担当者', work_content: '機器設置・動作確認', work_time: n % 2 ? '夜' : '昼' });
    const items = ['手島', 'コーワ東京', '森'].flatMap(user =>
      [today, tomorrow].flatMap(date => Array.from({ length: 12 }, (_, n) => item(user, date, n))));
    const data = { items, updated_at: '2026-09-09T03:00:00Z' };
    let dataUnavailable = false;
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname !== 'board.test') return route.abort();
      if (url.pathname.startsWith('/.netlify/functions/')) {
        assert.equal(route.request().method(), 'GET', 'No writes during tests');
        if (url.searchParams.has('signage') && dataUnavailable) return route.fulfill({ status: 503, json: { error: 'unavailable' } });
        return route.fulfill({ json: url.searchParams.has('signage') ? data : { members: [], projects: [], daily: {}, _revision: 'test' } });
      }
      const filename = path.join(root, 'static', url.pathname === '/' ? 'index.html' : url.pathname);
      const ext = path.extname(filename);
      return route.fulfill({ body: fs.readFileSync(filename), contentType:
        ext === '.html' ? 'text/html; charset=utf-8' : ext === '.css' ? 'text/css' : ext === '.js' ? 'application/javascript' : 'application/json' });
    });
    await page.goto('http://board.test/');
    await page.waitForFunction(() => dateGroups.length === 6);
    check('company-first order', await page.evaluate(() => dateGroups.map(g => g.company.id + '/' + g.label)),
      ['hsj/本日', 'hsj/明日', 'kowa/本日', 'kowa/明日', 'koyo/本日', 'koyo/明日']);
    check('pagination has no mixed companies', await page.evaluate(() => dateGroups.every(g =>
      g.items.every(i => companyOfMember(i.user).id === g.company.id) && g.totalPages === 2)), true);
    // More than 10 pages: the 5-minute refresh must not send the display back to HSJ.
    await page.clock.runFor(10 * 30000 + 1000);
    check('automatic rotation passes refresh and reaches Koyo', await page.locator('#header-company').textContent(), '光洋エンジニアリング');
    check('global totals retained', await page.locator('#total-count').textContent(), '72');
    check('new members route to Koyo; HSJ Mori unchanged', await page.evaluate(() =>
      ['森', '神邊', '猪股', '日向野', '森_吉村'].map(m => companyOfMember(m).id)), ['koyo', 'koyo', 'koyo', 'koyo', 'hsj']);
    check('empty company skipped and unknown retained', await page.evaluate(({ today, tomorrow }) =>
      buildDateGroups([{ user: '神邊', date: tomorrow }, { user: '不明', date: today }, { user: '', date: today }], today, tomorrow)
        .map(g => [g.company.id, g.items.length]), { today, tomorrow }), [['koyo', 1], ['other', 2]]);
    check('month boundary and fallback dates', await page.evaluate(() =>
      buildDateGroups([{ user: '森', date: '2026-10-01' }, { user: '森', date: '2026-09-30' }], '2026-09-09', '2026-09-10')
        .map(g => g.date)), ['2026-09-30', '2026-10-01']);
    check('background refresh preserves company/date/page', await page.evaluate(data => {
      stopAllTimers(); currentGroupIdx = 4; currentPage = 1;
      renderBoard({ ...data, items: data.items.filter(i => !(i.user === '手島' && i.date === '2026-09-09')) });
      return [dateGroups[currentGroupIdx].company.id, currentPage];
    }, data), ['koyo', 1]);
    await page.evaluate(data => { renderBoard(data); stopAllTimers(); }, data);
    fs.mkdirSync(path.join(root, '.netlify/qa'), { recursive: true });
    for (const width of [1920, 1366, 1024]) {
      await page.setViewportSize({ width, height: width === 1920 ? 1080 : 768 });
      await page.evaluate(() => { stopAllTimers(); currentPage = 0; renderCurrentGroupPage(); });
      check('header and rows fit ' + width, await page.evaluate(() => {
        const left = document.querySelector('.header-left').getBoundingClientRect();
        const center = document.querySelector('.header-center').getBoundingClientRect();
        const right = document.querySelector('.header-right').getBoundingClientRect();
        const rows = [...document.querySelectorAll('.board-row')];
        return left.right <= center.left && center.right <= right.left &&
          rows.every(r => r.getBoundingClientRect().bottom <= document.querySelector('.board-footer').getBoundingClientRect().top);
      }), true);
      await page.screenshot({ path: path.join(root, '.netlify/qa/company-' + width + '.png'), animations: 'disabled' });
    }
    await page.clock.setSystemTime(new Date('2026-09-09T23:59:59+09:00'));
    await page.clock.runFor(2000);
    check('midnight updates dates without waiting for refresh', await page.evaluate(() => ({
      date: dateGroups[currentGroupIdx].date, label: dateGroups[currentGroupIdx].label,
      header: document.getElementById('header-today').textContent,
      todayCount: document.getElementById('summary-today').textContent,
      tomorrowCount: document.getElementById('summary-tomorrow').textContent
    })), { date: tomorrow, label: '本日', header: '本日9/10 木曜日', todayCount: '36', tomorrowCount: '0' });
    dataUnavailable = true;
    await page.evaluate(() => forceReload());
    check('outage retains cached schedules with visible warning', await page.evaluate(() => ({
      rows: document.querySelectorAll('.board-row').length > 0,
      warning: document.getElementById('connection-status').textContent,
      visible: !document.getElementById('connection-status').hidden
    })), { rows: true, warning: '通信エラー：前回取得した予定を表示中', visible: true });
    check('offline footer fits 1024 without covering schedules', await page.evaluate(() => {
      const elements = [...document.querySelector('.board-footer').children].filter(e => !e.hidden);
      const rects = elements.map(e => e.getBoundingClientRect());
      return rects.every((r, i) => r.left >= 0 && r.right <= innerWidth &&
        (i === 0 || rects[i - 1].right <= r.left));
    }), true);
    await page.reload();
    await page.waitForFunction(() => !document.getElementById('connection-status').hidden);
    check('offline reload uses persisted cache', await page.locator('.board-row').count() > 0, true);
    dataUnavailable = false;
    await page.evaluate(() => forceReload());
    check('recovery clears cached warning', await page.locator('#connection-status').isHidden(), true);
    await page.evaluate(() => {
      window.savedFetch = window.fetch;
      window.fetch = function(url, options) {
        return new Promise((resolve, reject) => options.signal.addEventListener('abort',
          () => reject(new DOMException('Timed out', 'AbortError')), { once: true }));
      };
      window.pendingTimeoutCheck = fetchData().then(data => renderBoard(data));
    });
    await page.clock.runFor(30001);
    await page.evaluate(() => { window.fetch = window.savedFetch; return window.pendingTimeoutCheck; });
    check('timeout retains schedules and reports stale data', await page.evaluate(() =>
      !document.getElementById('connection-status').hidden && document.querySelectorAll('.board-row').length > 0), true);
    await page.evaluate(() => forceReload());
    check('recovery after timeout clears warning', await page.locator('#connection-status').isHidden(), true);

    check('empty state clears groups', await page.evaluate(() => { renderBoard({ items: [] }); return dateGroups.length; }), 0);
    await page.goto('http://board.test/input.html');
    await page.waitForFunction(() => typeof KNOWN_MEMBERS !== 'undefined' && KNOWN_MEMBERS.includes('神邊'));
    check('input fallback matches organization', await page.evaluate(() => DEFAULT_ORG), org);
    check('input Koyo roster', await page.evaluate(() => membersOfCompany('koyo')), ['光洋', '森', '神邊', '猪股', '日向野']);
    await page.screenshot({ path: path.join(root, '.netlify/qa/input.png') });
    check('browser runtime errors', errors, []);
    console.log(`${passed} PASS, 0 FAIL`);
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
