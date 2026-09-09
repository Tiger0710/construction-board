// Read-only diagnosis of the current month-boundary behavior. No network writes.
const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');

(async () => {
  const browser = await chromium.launch({ channel: 'chrome' });
  let count = 0;
  const check = (label, actual, expected) => {
    assert.deepEqual(actual, expected); console.log('CONFIRMED ' + label); count++;
  };
  try {
    const page = await browser.newPage();
    await page.clock.install({ time: new Date('2026-09-09T12:00:00+09:00') });
    const project = { id: 'original', client: '検証客先', title: '月またぎ検証工事',
      start_date: '2026-09-30', end_date: '2026-10-03', our_person: '森', partner: '検証協力会社' };
    let months = { '2609': { projects: [project], daily: {}, _sha: 'september' },
      '2610': { projects: [], daily: {} } };
    let requests = [], saves = [];
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname !== 'board.test') return route.abort();
      if (url.pathname.startsWith('/.netlify/functions/')) {
        if (route.request().method() === 'PUT') {
          saves.push(route.request().postDataJSON());
          return route.fulfill({ json: { sha: 'mock-save' } });
        }
        if (url.searchParams.has('signage')) return route.fulfill({ json: { items: [] } });
        if (!url.searchParams.has('user')) return route.fulfill({ json: { members: ['森'] } });
        const month = url.searchParams.get('month'); requests.push(month);
        return route.fulfill({ json: months[month] || { projects: [], daily: {} } });
      }
      const filename = path.join(root, 'static', url.pathname);
      return route.fulfill({ body: fs.readFileSync(filename), contentType:
        filename.endsWith('.html') ? 'text/html; charset=utf-8' : 'application/json' });
    });
    await page.goto('http://board.test/input.html#%E6%A3%AE/2610');
    await page.waitForFunction(() => !state.loading && state.adjProjects.length === 1);
    check('September current date: October shows September project as ghost',
      await page.evaluate(() => [state.data.projects.length, state.adjProjects.length]), [0, 1]);
    await page.evaluate(() => onAdjClick('2609'));
    await page.waitForFunction(() => !state.loading && state.month === '2609');
    check('clicking October ghost moves entire view back to September',
      await page.evaluate(() => [state.month, document.querySelectorAll('.gantt th.dh').length]), ['2609', 30]);
    // A month-spanning new project is saved only in the selected month file.
    await page.evaluate(() => {
      openModal();
      document.getElementById('m-client').value = '検証客先';
      document.getElementById('m-title').value = '新規月またぎ';
      document.getElementById('m-start').value = '2026-09-30';
      document.getElementById('m-end').value = '2026-10-03';
      saveProject();
    });
    await page.evaluate(() => saveData());
    check('save sends one September file, with October end date',
      saves.map(s => [s.month, s.data.projects.at(-1).end_date]), [['2609', '2026-10-03']]);
    months['2610'] = { projects: [{ ...project, id: 'recreated-in-october' }], daily: {} };
    await page.evaluate(async () => { state.month = '2610'; await loadData('2610'); });
    check('recreated project plus original produce two rows',
      await page.locator('.gantt tbody tr').count(), 2);
    months['2610'] = { projects: [{ ...project }], daily: {} };
    await page.evaluate(() => loadData('2610'));
    check('only identical IDs suppress ghost row',
      await page.locator('.gantt tbody tr').count(), 1);
    months['2610'] = { projects: [], daily: {} };
    await page.clock.setSystemTime(new Date('2026-10-01T12:00:00+09:00'));
    requests = [];
    await page.evaluate(() => loadData('2610'));
    check('October actual date: previous month is no longer fetched', requests, ['2610']);
    check('September project spanning October disappears from October input',
      await page.evaluate(() => [state.data.projects.length, state.adjProjects.length]), [0, 0]);
    console.log(`${count} behaviors reproduced; mock PUT only; no production writes`);
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
