const { chromium } = require('playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
(async () => {
  const browser = await chromium.launch({channel:'chrome'});
  let count=0;
  const check=(name,actual,expected)=>{assert.deepEqual(actual,expected,name);console.log('PASS '+name);count++;};
  try {
    const page=await browser.newPage({viewport:{width:1366,height:900}});
    const errors=[];page.on('pageerror',e=>errors.push(e.message));
    await page.clock.install({time:new Date('2026-09-28T12:00:00+09:00')});
    let data={projects:[{id:'legacy',client:'客先',title:'既存工事',start_date:'2026-09-28',end_date:'2026-10-05'}],daily:{}};
    let revision='r1',savedUser;
    await page.route('**/*',async route=>{
      const u=new URL(route.request().url());
      if(u.hostname!=='board.test')return route.abort();
      if(u.pathname.startsWith('/.netlify/functions/')){
        if(route.request().method()==='PUT'){
          const body=route.request().postDataJSON();savedUser=body.user;
          assert.equal(body.revision,revision);data=body.data;revision+='x';
          return route.fulfill({json:{success:true,revision}});
        }
        if(u.searchParams.has('signage'))return route.fulfill({json:{items:[]}});
        return route.fulfill({json:u.searchParams.has('user')?{...data,_revision:revision}:{members:['光洋']}});
      }
      const f=path.join(root,'static',u.pathname);
      return route.fulfill({body:fs.readFileSync(f),contentType:f.endsWith('.html')?'text/html; charset=utf-8':f.endsWith('.js')?'application/javascript':'application/json'});
    });
    await page.goto('http://board.test/input.html');
    await page.getByRole('button',{name:'佐藤 データあり'}).waitFor();
    check('entry uses Sato label',await page.getByRole('button',{name:'佐藤 データあり'}).count(),1);
    await page.getByRole('button',{name:'佐藤 データあり'}).click();
    await page.waitForFunction(()=>state.ready);
    check('Sato retains existing member identity',await page.evaluate(()=>[state.user,memberLabel(state.user)]),['光洋','佐藤']);
    await page.evaluate(()=>openModal('legacy'));
    check('legacy project defaults to day',await page.locator('#m-default-shift').inputValue(),'day');
    await page.evaluate(()=>{closeModal();openModal();});
    await page.locator('#m-client').fill('テスト客先');await page.locator('#m-title').fill('夜間の月またぎ工事');
    await page.locator('#m-start').fill('2026-09-28');await page.locator('#m-end').fill('2026-10-05');
    await page.locator('#m-default-shift').selectOption('night');
    fs.mkdirSync(path.join(root,'.netlify/qa'),{recursive:true});
    await page.screenshot({path:path.join(root,'.netlify/qa/default-shift.png')});
    await page.evaluate(()=>saveProject());
    const id=await page.evaluate(()=>state.data.projects.find(p=>p.title==='夜間の月またぎ工事').id);
    check('project stores selected night default',await page.evaluate(id=>state.data.projects.find(p=>p.id===id).default_shift,id),'night');
    await page.evaluate(id=>openDayEditor(id,'2026-09-28'),id);
    check('unentered editor defaults to night only',await page.evaluate(()=>[document.getElementById('dm-day-on').checked,document.getElementById('dm-night-on').checked]),[false,true]);
    await page.evaluate(()=>{onDmField('night_work','夜間設置');closeDayEditor();});
    check('first work edit materializes night, not day',await page.evaluate(id=>getDayData(id,'2026-09-28'),id).then(d=>[d.day,d.night,d.night_work]),[false,true,'夜間設置']);
    await page.evaluate(id=>{openDayEditor(id,'2026-09-29');onDmToggle('day',true);closeDayEditor();},id);
    check('daily override can activate both shifts',await page.evaluate(id=>getDayData(id,'2026-09-29'),id).then(d=>[d.day,d.night]),[true,true]);
    await page.evaluate(id=>openModal(id),id);await page.locator('#m-default-shift').selectOption('day');await page.evaluate(()=>saveProject());
    check('default change keeps entered night and explicit shifts',await page.evaluate(id=>[getDayData(id,'2026-09-28').night,getDayData(id,'2026-09-29').night],id),[true,true]);
    check('default change affects only unentered day',await page.evaluate(id=>projectShifts(state.data.projects.find(p=>p.id===id),getDayData(id,'2026-10-02')),id),{day:true,night:false});
    check('automatic weekend remains off',await page.evaluate(id=>projectShifts(state.data.projects.find(p=>p.id===id),getDayData(id,'2026-10-03')),id),{day:false,night:false});
    await page.evaluate(id=>openModal(id),id);await page.locator('#m-default-shift').selectOption('night');await page.evaluate(()=>{saveProject();changeMonth(1);});
    check('October chart uses night default',await page.evaluate(id=>{const p=state.data.projects.find(p=>p.id===id);return projectShifts(p,getDayData(id,'2026-10-02'));},id),{day:false,night:true});
    await page.evaluate(()=>saveData());await page.evaluate(()=>loadData(state.month));
    check('saved selector survives reload',await page.evaluate(id=>state.data.projects.find(p=>p.id===id).default_shift,id),'night');
    check('Sato saves into original member data',savedUser,'光洋');
    check('persisted explicit work retained',data.daily[id+'/2026-09-28'].night_work,'夜間設置');
    check('new automatic weekend off is identifiable',await page.evaluate(id=>getDayData(id,'2026-10-03')._weekend_auto,id),'off');
    await page.evaluate(id=>{openDayEditor(id,'2026-10-04');onDmToggle('day',false);closeDayEditor();openModal(id);},id);
    await page.locator('#m-weekend-policy').selectOption('work');await page.evaluate(()=>saveProject());
    check('work policy restores automatic Saturday to base night',await page.evaluate(id=>projectShifts(state.data.projects.find(p=>p.id===id),getDayData(id,'2026-10-03')),id),{day:false,night:true});
    check('explicit Sunday off remains off',await page.evaluate(id=>projectShifts(state.data.projects.find(p=>p.id===id),getDayData(id,'2026-10-04')),id),{day:false,night:false});
    await page.evaluate(id=>openModal(id),id);await page.locator('#m-weekend-policy').selectOption('off');await page.evaluate(()=>saveProject());
    check('switching back restores automatic weekend off',await page.evaluate(id=>getDayData(id,'2026-10-03').day,id),false);
    await page.evaluate(id=>{openDayEditor(id,'2026-10-03');onDmToggle('night',true);onDmField('night_work','土曜夜の手入力');closeDayEditor();openModal(id);},id);
    await page.locator('#m-weekend-policy').selectOption('work');await page.locator('#m-end').fill('2026-10-12');
    await page.screenshot({path:path.join(root,'.netlify/qa/weekend-policy.png')});await page.evaluate(()=>saveProject());
    check('entered Saturday work survives policy switch',await page.evaluate(id=>getDayData(id,'2026-10-03').night_work,id),'土曜夜の手入力');
    check('work policy extends into next weekend without new off',await page.evaluate(id=>getDayData(id,'2026-10-10'),id),null);
    await page.evaluate(()=>saveData());await page.evaluate(()=>loadData(state.month));await page.evaluate(id=>openModal(id),id);
    check('weekend choice survives save and reload',await page.locator('#m-weekend-policy').inputValue(),'work');
    await page.locator('#m-weekend-policy').selectOption('off');await page.evaluate(()=>saveProject());
    check('off policy fills extended weekend',await page.evaluate(id=>getDayData(id,'2026-10-10')._weekend_auto,id),'off');
    check('off policy does not erase entered Saturday night',await page.evaluate(id=>getDayData(id,'2026-10-03').night,id),true);
    await page.evaluate(()=>{state.data.daily['legacy/2026-10-03']={day:false,night:false};openModal('legacy');});
    await page.locator('#m-weekend-policy').selectOption('work');await page.evaluate(()=>saveProject());
    check('legacy untagged weekend off remains protected',await page.evaluate(()=>getDayData('legacy','2026-10-03')),{day:false,night:false});
    check('no browser errors',errors,[]);
    console.log(`${count} PASS, 0 FAIL`);
  }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
