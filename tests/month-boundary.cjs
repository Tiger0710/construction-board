// Browser regressions. API writes are mocked; no production data is changed.
const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
(async () => {
  const browser = await chromium.launch({ channel: 'chrome' });
  let count = 0;
  const check = (name, actual, expected) => { assert.deepEqual(actual, expected, name); console.log('PASS ' + name); count++; };
  try {
    const page = await browser.newPage({ viewport: { width: 1366, height: 768 } });
    const errors = []; page.on('pageerror', e => errors.push(e.message));
    await page.clock.install({ time: new Date('2026-10-01T12:00:00+09:00') });
    let data = { projects: [{id:'stable',client:'検証',title:'長期工事',start_date:'2026-07-01',end_date:'2027-02-02',our_person:'森'}], daily:{'stable/2026-09-30':{day:true,day_work:'前月の作業',night:true,night_work:'夜間確認'}} };
    let revision = 'r1', failLoad = false, failSave = false, writes = [];
    await page.route('**/*', async route => {
      const u = new URL(route.request().url());
      if (u.hostname !== 'board.test') return route.abort();
      if (u.pathname.startsWith('/.netlify/functions/')) {
        if (route.request().method() === 'PUT') {
          const body = route.request().postDataJSON(); writes.push(body);
          if (failSave || body.revision !== revision) return route.fulfill({ status:409,json:{error:'競合しています'} });
          assert.equal(body.scope,'all'); data = body.data; revision += 'x';
          return route.fulfill({ json:{success:true,revision} });
        }
        if (u.searchParams.has('signage')) return route.fulfill({json:{items:[]}});
        if (!u.searchParams.has('user')) return route.fulfill({json:{members:['森']}});
        if(failLoad) return route.fulfill({status:503,json:{error:'取得失敗'}});
        assert.equal(u.searchParams.get('scope'),'all');
        return route.fulfill({json:{...data,_revision:revision}});
      }
      const f=path.join(root,'static',u.pathname);
      return route.fulfill({body:fs.readFileSync(f),contentType:f.endsWith('.html')?'text/html; charset=utf-8':f.endsWith('.js')?'application/javascript':'application/json'});
    });
    await page.goto('http://board.test/input.html#%E6%A3%AE/2610');
    await page.waitForFunction(()=>!state.loading && state.data.projects.length===1);
    check('long project appears in October after real month change',await page.locator('.gantt tbody tr').count(),1);
    await page.evaluate(()=>openDayEditor('stable','2026-10-01'));
    check('October cell edits in October',await page.evaluate(()=>[state.month,dmState.dateStr]),['2610','2026-10-01']);
    await page.evaluate(()=>copyFromPrevDay());
    check('copy previous day across month preserves both shifts',await page.evaluate(()=>getDayData('stable','2026-10-01').night_work),'夜間確認');
    await page.evaluate(()=>closeDayEditor());
    await page.evaluate(()=>changeMonth(1));
    check('month navigation retains unsaved changes and identity',await page.evaluate(()=>[state.month,state.dirty,state.data.projects[0].id]),['2611',true,'stable']);
    await page.evaluate(()=>saveData());
    check('save uses canonical all-month scope',writes.map(w=>[w.scope,w.data.projects.length,w.data.projects[0].id]),[['all',1,'stable']]);
    await page.evaluate(()=>loadData(state.month));
    check('saved October edit survives reload in November',await page.evaluate(()=>getDayData('stable','2026-10-01').day_work),'前月の作業');
    await page.evaluate(()=>{state.month='2701';renderGantt();});
    check('year boundary keeps a single editable project',await page.locator('.gantt tbody tr').count(),1);
    await page.evaluate(()=>{state.month='2703';renderGantt();});
    check('out-of-range projects excluded from displayed month',await page.locator('.gantt tbody tr').count(),0);
    await page.evaluate(()=>{state.month='2610';renderGantt();onDmField;openDayEditor('stable','2026-10-02');onDmField('day_work','保存前の作業');closeDayEditor();});
    failSave=true; await page.evaluate(()=>saveData());
    check('conflict retains unsaved changes',await page.evaluate(()=>[state.dirty,getDayData('stable','2026-10-02').day_work]),[true,'保存前の作業']);
    failSave=false; await page.evaluate(()=>saveData());
    failLoad=true; await page.evaluate(()=>loadData(state.month));
    check('failed load cannot add or save empty replacement',await page.evaluate(()=>{openModal();return [state.ready,document.getElementById('save-btn').disabled,document.getElementById('modal').classList.contains('hidden')];}),[false,true,true]);
    failLoad=false;await page.evaluate(()=>loadData(state.month));
    check('retry restores usable data',await page.evaluate(()=>[state.ready,state.data.projects.length]),[true,1]);
    check('no browser runtime errors',errors,[]);
    console.log(`${count} PASS, 0 FAIL`);
  } finally { await browser.close(); }
})().catch(e=>{console.error(e);process.exitCode=1;});
