// All HTTP calls, including writes, are intercepted. Never contacts a real backend.
const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const project = (id, title, start, end) => ({ id, client: '検証客先', title,
  our_person: '森', safety_person: '神邊', partner: '協力会社', partner_person: '猪股',
  start_date: start, end_date: end });
const fixture = () => ({ projects: [
  project('a', '月またぎ連続工事', '2026-08-30', '2026-08-31'),
  project('b', '月またぎ連続工事', '2026-09-01', '2026-09-03'),
  project('off', '重複期間の工事', '2026-08-31', '2026-09-02'),
  project('on', '重複期間の工事', '2026-09-01', '2026-09-03'),
  project('other', '無関係な工事', '2026-09-01', '2026-09-30')
], daily: {
  'a/2026-08-31': { day: true, night: true, day_work: '基礎作業', night_work: '夜間確認', night_our_person: '日向野' },
  'b/2026-09-02': { day: true, night: false, day_work: '配線', day_priority: '重点', day_priority_detail: '安全確認' },
  'off/2026-09-01': { day: false, night: false },
  'other/2026-09-03': { day: true, day_work: '無関係な入力' }
} });

(async () => {
  const browser = await chromium.launch({ channel: 'chrome' });
  let count = 0;
  const check = (name, actual, expected) => {
    assert.deepEqual(actual, expected, name); console.log('PASS ' + name); count++;
  };
  try {
    const page = await browser.newPage({ viewport: { width: 1366, height: 900 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.clock.install({ time: new Date('2026-09-09T12:00:00+09:00') });
    let backend = fixture(), revision = 'merge-r1', writes = [], reads = [];
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname !== 'board.test') return route.abort();
      if (url.pathname.startsWith('/.netlify/functions/')) {
        if (route.request().method() === 'PUT') {
          const body = route.request().postDataJSON();
          writes.push(body);
          assert.equal(body.scope, 'all');
          assert.equal(body.revision, revision);
          backend = structuredClone(body.data);
          revision += 'x';
          return route.fulfill({ json: { success: true, revision } });
        }
        if (url.searchParams.has('signage')) return route.fulfill({ json: { items: [] } });
        if (!url.searchParams.has('user')) return route.fulfill({ json: { members: ['森'] } });
        assert.equal(url.searchParams.get('scope'), 'all');
        reads.push(url.href);
        return route.fulfill({ json: { ...backend, _revision: revision } });
      }
      const file = path.join(root, 'static', url.pathname);
      return route.fulfill({ body: fs.readFileSync(file), contentType: file.endsWith('.html') ?
        'text/html; charset=utf-8' : file.endsWith('.js') ? 'application/javascript' : 'application/json' });
    });
    await page.goto('http://board.test/input.html#%E6%A3%AE/2609');
    await page.waitForFunction(() => state.ready && !state.loading);

    check('routine input has no merge button or modal', await page.locator('#merge-btn, #merge-modal').count(), 0);
    check('routine input has no merge script or callable merge UI', await page.evaluate(() => [
      document.querySelectorAll('script[src="project-merge.js"]').length,
      typeof ProjectMerge, typeof openMerge, typeof confirmMerge, typeof mergeState
    ]), [0, 'undefined', 'undefined', 'undefined', 'undefined']);
    check('loading does not automatically consolidate data', await page.evaluate(() => [state.data.projects.length, state.dirty]), [5, false]);
    check('loading does not write', writes.length, 0);

    await page.evaluate(() => {
      openModal();
      const source = state.data.projects.find(p => p.id === 'b');
      const fields = { client: 'm-client', title: 'm-title', our_person: 'm-our', safety_person: 'm-safety',
        partner: 'm-partner', partner_person: 'm-partner-person' };
      Object.entries(fields).forEach(([key, id]) => { document.getElementById(id).value = source[key]; });
      document.getElementById('m-start').value = '2026-09-04';
      document.getElementById('m-end').value = '2026-10-10';
    });
    let duplicateMessage = '';
    page.once('dialog', dialog => { duplicateMessage = dialog.message(); return dialog.dismiss(); });
    await page.evaluate(() => saveProject());
    check('continued registration warns and cancel opens existing project', await page.evaluate(() => [state.editingId,
      document.getElementById('m-start').value, document.getElementById('m-end').value, state.data.projects.length]),
      ['b', '2026-09-01', '2026-09-03', 5]);
    check('duplicate warning is shown', duplicateMessage.length > 0, true);
    await page.evaluate(() => {
      document.getElementById('m-end').value = '2026-10-10';
      document.getElementById('m-weekend-policy').value = 'work';
      saveProject();
      openDayEditor('b', '2026-10-01');
      onDmField('night_work', 'October continuation');
      closeDayEditor();
    });
    check('editing existing period preserves identity and entered shifts', await page.evaluate(() => [state.data.projects.length,
      state.data.projects.find(p => p.id === 'b').end_date, state.data.daily['a/2026-08-31'], state.data.daily['b/2026-09-02'], state.dirty]),
      [5, '2026-10-10', fixture().daily['a/2026-08-31'], fixture().daily['b/2026-09-02'], true]);
    const readCount = reads.length;
    await page.evaluate(() => changeMonth(1));
    check('month navigation retains unsaved period and daily input', await page.evaluate(() => [state.month, state.dirty,
      state.data.projects.find(p => p.id === 'b').end_date, state.data.daily['b/2026-10-01'].night_work]),
      ['2610', true, '2026-10-10', 'October continuation']);
    check('month navigation does not reload unsaved input', reads.length, readCount);
    check('same project stays editable in following month', await page.evaluate(() => {
      openModal('b'); return [state.editingId, document.getElementById('m-end').value];
    }), ['b', '2026-10-10']);
    await page.evaluate(() => { closeModal(); return saveData(); });
    check('explicit save uses all-month scope and revision', writes.map(w => [w.scope, w.revision, w.data.projects.length]), [['all', 'merge-r1', 5]]);
    await page.evaluate(() => loadData(state.month));
    check('save and reload preserve following-month details', await page.evaluate(() => [state.data.projects.length,
      state.data.projects.find(p => p.id === 'b').end_date, state.data.daily['b/2026-10-01'].night_work, state.dirty]),
      [5, '2026-10-10', 'October continuation', false]);
    await page.setViewportSize({ width: 390, height: 844 });
    await page.evaluate(() => openModal('b'));
    check('small viewport keeps normal edit modal in bounds', await page.locator('#modal .modal').evaluate(el => {
      const r = el.getBoundingClientRect(); return r.left >= 0 && r.right <= innerWidth;
    }), true);
    check('default shift and weekend inputs remain available', await page.evaluate(() => [
      document.getElementById('m-default-shift').value, document.getElementById('m-weekend-policy').value]), ['day', 'work']);
    check('no browser runtime errors', errors, []);
    console.log(`${count} PASS, 0 FAIL`);
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
