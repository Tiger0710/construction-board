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

    await page.locator('#merge-btn').click();
    check('candidate modal lists two related groups', await page.locator('#merge-group option').count(), 2);
    check('opening candidates never mutates or saves', await page.evaluate(() => [state.data.projects.length, state.dirty]), [5, false]);
    await page.locator('.merge-project[value="b"]').uncheck();
    check('one selected record cannot be merged', await page.locator('#merge-confirm').isDisabled(), true);
    await page.locator('.merge-project[value="b"]').check();
    check('nonconflicting periods enable explicit confirmation', await page.locator('#merge-confirm').isEnabled(), true);
    const original = await page.evaluate(() => JSON.stringify(state.data));
    page.once('dialog', dialog => dialog.dismiss());
    await page.locator('#merge-confirm').click();
    check('cancel confirmation preserves complete data', await page.evaluate(() => JSON.stringify(state.data)), original);
    check('confirmation cancellation keeps preview open', await page.locator('#merge-modal').isVisible(), true);

    await page.locator('#merge-group').selectOption('1');
    check('changing candidate resets preview and requires conflict choice', await page.evaluate(() => [mergeState.preview.primaryId,
      mergeState.preview.conflicts.length, document.getElementById('merge-confirm').disabled]), ['off', 1, true]);
    check('conflict explains implicit default and explicit off', await page.locator('#merge-details').innerText().then(t =>
      [t.includes('2026-09-01'), t.includes('その日は未入力'), t.includes('昼稼働: なし'), t.includes('昼稼働: あり')]), [true, true, true, true]);
    await page.evaluate(() => confirmMerge());
    check('direct confirm without radio is safely rejected', await page.evaluate(() => [state.data.projects.length,
      document.getElementById('merge-message').textContent.includes('選択')]), [5, true]);
    await page.locator('input[name="merge-choice-0"][value="off"]').check();
    check('conflict choice enables confirmation', await page.locator('#merge-confirm').isEnabled(), true);
    const qa = path.join(root, '.netlify', 'qa');
    fs.mkdirSync(qa, { recursive: true });
    await page.screenshot({ path: path.join(qa, 'merge.png') });
    await page.setViewportSize({ width: 390, height: 844 });
    check('small viewport keeps merge controls within screen', await page.locator('#merge-modal .modal').evaluate(el => {
      const r = el.getBoundingClientRect();
      return r.left >= 0 && r.right <= innerWidth && r.top >= 0 && r.bottom <= innerHeight;
    }), true);
    await page.locator('#merge-confirm').scrollIntoViewIfNeeded();
    check('small viewport confirm is reachable', await page.locator('#merge-confirm').isVisible(), true);
    await page.screenshot({ path: path.join(qa, 'merge-mobile.png') });
    page.once('dialog', dialog => dialog.accept());
    await page.locator('#merge-confirm').click();
    check('selected explicit off wins without changing unrelated daily records', await page.evaluate(() => [state.data.projects.length,
      state.data.daily['off/2026-09-01'], state.data.daily['other/2026-09-03'], state.dirty]),
    [4, { day: false, night: false }, fixture().daily['other/2026-09-03'], true]);
    check('local merge does not automatically write', writes.length, 0);

    const readCount = reads.length;
    await page.evaluate(() => changeMonth(1));
    check('month navigation retains unsaved merged project with no reload', await page.evaluate(() => [state.month, state.dirty,
      state.data.projects.some(p => p.id === 'on'), state.data.daily['off/2026-09-01'].day]), ['2610', true, false, false]);
    check('month navigation does not read replacement data', reads.length, readCount);
    await page.evaluate(() => changeMonth(-1));
    await page.evaluate(() => saveData());
    check('explicit save uses all-month scope and expected revision', writes.map(w => [w.scope, w.revision, w.data.projects.length]), [['all', 'merge-r1', 4]]);
    await page.evaluate(() => loadData(state.month));
    check('off selection survives save/reload', await page.evaluate(() => [state.data.projects.length, state.data.daily['off/2026-09-01'].day, state.dirty]), [4, false, false]);

    await page.setViewportSize({ width: 1366, height: 900 });
    await page.locator('#merge-btn').click();
    page.once('dialog', dialog => dialog.accept());
    await page.locator('#merge-confirm').click();
    check('nonconflicting merge preserves month endpoints and all day/night details', await page.evaluate(() => {
      const p = state.data.projects.find(p => p.id === 'a');
      return [p.start_date, p.end_date, state.data.projects.some(p => p.id === 'b'),
        state.data.daily['a/2026-08-31'], state.data.daily['a/2026-09-02']];
    }), ['2026-08-30', '2026-09-03', false, fixture().daily['a/2026-08-31'], fixture().daily['b/2026-09-02']]);

    await page.evaluate(() => {
      openModal();
      const vals = { 'm-client': '検証客先', 'm-title': '月またぎ連続工事', 'm-our': '森', 'm-safety': '神邊',
        'm-partner': '協力会社', 'm-partner-person': '猪股', 'm-start': '2026-09-04', 'm-end': '2026-09-10' };
      Object.entries(vals).forEach(([id, value]) => { document.getElementById(id).value = value; });
    });
    let duplicateMessage = '';
    page.once('dialog', dialog => { duplicateMessage = dialog.message(); return dialog.dismiss(); });
    await page.evaluate(() => saveProject());
    check('related next-period registration warns and cancel opens existing project', await page.evaluate(() => [state.editingId,
      document.getElementById('m-start').value, document.getElementById('m-end').value, state.data.projects.length]), ['a', '2026-08-30', '2026-09-03', 3]);
    check('duplicate warning explains continued-project editing', duplicateMessage.includes('継続工事') && duplicateMessage.includes('既存案件を編集'), true);
    await page.evaluate(() => closeModal());

    backend = fixture(); revision = 'reset-r1';
    await page.evaluate(() => { state.dirty = false; return loadData(state.month); });
    await page.locator('#merge-btn').click();
    await page.locator('#merge-group').selectOption('1');
    await page.locator('input[name="merge-choice-0"][value="on"]').check();
    page.once('dialog', dialog => dialog.accept());
    await page.locator('#merge-confirm').click();
    check('choosing implicit on is preserved explicitly after consolidation', await page.evaluate(() => state.data.daily['off/2026-09-01']), { day: true, night: false });

    await page.locator('#merge-btn').click();
    await page.evaluate(() => { state.data.daily['other/2026-09-04'] = { day: true, day_work: '新しい変更' }; markDirty(); });
    await page.evaluate(() => confirmMerge());
    check('stale preview refreshes instead of applying over newer edits', await page.evaluate(() => [state.data.projects.length,
      state.data.daily['other/2026-09-04'].day_work, mergeState.version === state.editVersion]), [4, '新しい変更', true]);
    check('no browser runtime errors', errors, []);
    console.log(`${count} PASS, 0 FAIL`);
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
