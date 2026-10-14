const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..');
const stored = {
  projects: [
    { id: 'changed', client: 'KNSI', title: '安品変更済み', our_person: '森', safety_person: '保科',
      partner: '', partner_person: '', start_date: '2026-10-01', end_date: '2026-10-31' },
    { id: 'blank', client: '', title: '空欄案件', our_person: '', safety_person: '',
      partner: '', partner_person: '', start_date: '2026-10-01', end_date: '2026-10-31' }
  ],
  daily: {}, _revision: 'r1'
};

(async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname !== 'board.test') return route.abort();
      if (url.pathname.startsWith('/.netlify/functions/')) {
        if (url.searchParams.has('signage')) return route.fulfill({ json: { items: [] } });
        if (url.searchParams.has('user')) return route.fulfill({ json: JSON.parse(JSON.stringify(stored)) });
        return route.fulfill({ json: { members: ['森_吉村'] } });
      }
      const file = path.join(root, 'static', url.pathname.slice(1));
      return route.fulfill({ body: fs.readFileSync(file), contentType:
        file.endsWith('.html') ? 'text/html; charset=utf-8' : 'application/json' });
    });

    await page.goto('http://board.test/input.html#森_吉村/2610');
    await page.waitForFunction(() => state.ready && !state.loading && state.user === '森_吉村');

    const val = id => page.locator(id).inputValue();
    const hints = () => page.locator('#modal .fixed-hint').count();

    // 編集: 保存済みの値を表示し、初期値で上書きしない
    await page.evaluate(() => openModal('changed'));
    assert.equal(await val('#m-safety'), '保科');
    assert.equal(await val('#m-client'), 'KNSI');
    assert.equal(await val('#m-our'), '森');
    assert.equal(await hints(), 0);
    console.log('PASS edit modal shows the saved safety person, not the default');

    // 編集して別項目だけ変更・保存しても安品は戻らない
    await page.locator('#m-partner').fill('協力A');
    await page.evaluate(() => saveProject());
    const p = await page.evaluate(() => state.data.projects.find(x => x.id === 'changed'));
    assert.equal(p.safety_person, '保科');
    assert.equal(p.partner, '協力A');
    console.log('PASS saving an edit keeps the changed safety person');

    // 空欄の既存案件も空欄のまま表示（勝手に補完しない）
    await page.evaluate(() => openModal('blank'));
    assert.equal(await val('#m-safety'), '');
    assert.equal(await val('#m-client'), '');
    await page.evaluate(() => closeModal());
    console.log('PASS blank existing project stays blank in edit modal');

    // 新規: 初期値が入り、バッジが出る
    await page.evaluate(() => openModal());
    assert.equal(await val('#m-client'), 'KNSI');
    assert.equal(await val('#m-our'), '森');
    assert.equal(await val('#m-safety'), '髙橋');
    assert.equal(await hints(), 3);
    await page.evaluate(() => closeModal());
    console.log('PASS new project modal still fills defaults with badges');

    // 新規→編集の順でもバッジが残らない
    await page.evaluate(() => openModal('changed'));
    assert.equal(await hints(), 0);
    assert.equal(await val('#m-safety'), '保科');
    console.log('PASS badges are cleared when switching from new to edit');

    assert.deepEqual(errors, []);
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
