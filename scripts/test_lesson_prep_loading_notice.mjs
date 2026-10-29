import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import http from 'node:http';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {chromium} from 'playwright';

const root=fileURLToPath(new URL('../',import.meta.url));
const month='2026-11';
const row=ym=>({date:ym+'-02',time:'18:00～19:00',campus:'hon',groupKey:'hon_j1_A_math',grade:'j1',class:'A',subject:'math',room:'1',teacher:'工藤'});
const payload=ym=>({month:ym,count:1,schedule:[row(ym)],teacherSchedule:{items:[row(ym)]}});
const server=http.createServer(async(req,res)=>{
 try{const file=path.resolve(root,'.'+new URL(req.url,'http://local').pathname);if(!file.startsWith(root))throw Error('outside root');res.setHeader('Content-Type',path.extname(file)==='.js'?'text/javascript':path.extname(file)==='.html'?'text/html':'application/json');res.end(await fs.readFile(file));}
 catch{res.writeHead(404);res.end();}
});
await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
const base=`http://127.0.0.1:${server.address().port}`;
const browser=await chromium.launch({headless:true,...(process.env.CALENDAR_TEST_CHANNEL?{channel:process.env.CALENDAR_TEST_CHANNEL}:{})});
try{
 for(const width of [390,1280])for(const scenario of ['fast','delayed','failed','timeout','unavailable','changed-month']){
  const page=await browser.newPage({viewport:{width,height:900}}),errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  await page.clock.install({time:new Date('2026-11-02T12:00:00+09:00')});
  await page.addInitScript(()=>{
   window.visibleScheduleWarnings=[];
   new MutationObserver(()=>{const element=document.getElementById('schedule-source-status');if(element&&!element.hidden)window.visibleScheduleWarnings.push(element.textContent);}).observe(document,{subtree:true,childList:true,attributes:true,characterData:true});
  });
  let finish;
  const gate=new Promise(resolve=>{finish=resolve;});
  await page.route('**/*',async route=>{
   const url=new URL(route.request().url()),file=path.basename(url.pathname);
   if(url.pathname==='/api/schedule/calendar'){
    const ym=url.searchParams.get('month');
    if(ym!==month)return route.fulfill({json:payload(ym)});
    if(['delayed','timeout','changed-month'].includes(scenario))await gate;
    if(['failed','unavailable','changed-month'].includes(scenario))return route.fulfill({status:503,json:{error:'offline'}}).catch(()=>{});
    return route.fulfill({json:payload(ym)}).catch(()=>{});
   }
   if(url.pathname==='/api/recordings/publication')return route.fulfill({json:{rules:[]}});
   if(file.startsWith('schedule_')&&file.endsWith('.json')){
    if(scenario==='unavailable'&&file===`schedule_${month}.json`)return route.fulfill({status:503,body:''});
    const ym=file.slice(9,16);return route.fulfill({json:[row(ym)]});
   }
   if(file.endsWith('.json'))return route.fulfill({json:{month,entries:{},slots:{},rules:[]}});
   return url.origin===base?route.continue():route.abort();
  });
  await page.goto(base+'/lesson_prep.html',{waitUntil:'domcontentloaded'});
  await page.waitForFunction(()=>Boolean(window.CalendarScheduleSource));
  const notice=page.locator('#schedule-source-status');
  if(['delayed','timeout','changed-month'].includes(scenario)){
   await page.clock.runFor(1300);
   await page.waitForFunction(()=>document.querySelectorAll('#teacherSelect option').length>1);
   assert.equal(await notice.count(),0,'pending loads must not create hidden warning markup either');
   assert.equal(await notice.isVisible(),false,`${scenario}: a pending live response must never show a warning`);
   assert.deepEqual(await page.evaluate(()=>window.visibleScheduleWarnings),[]);
   await page.waitForFunction(()=>document.querySelectorAll('#teacherSelect option').length>1);
   if(scenario==='delayed'){
    finish();await page.waitForFunction(()=>Boolean(CalendarScheduleSource.snapshot('2026-11')));
    assert.equal(await notice.count(),0,'healthy delayed data never creates a warning element');
    assert.equal(await notice.isVisible(),false);
    assert.deepEqual(await page.evaluate(()=>window.visibleScheduleWarnings),[],'a healthy delayed response must not flash a warning');
   }else if(scenario==='changed-month'){
    await page.locator('#datePicker').fill('2026-12-02');await page.locator('#datePicker').dispatchEvent('change');
    await page.waitForFunction(()=>Boolean(CalendarScheduleSource.snapshot('2026-12')));
    const failed=page.waitForResponse(response=>response.url().includes('month=2026-11')&&response.status()===503);
    finish();await (await failed).finished();
    assert.equal(await notice.isVisible(),false);assert.deepEqual(await page.evaluate(()=>window.visibleScheduleWarnings),[]);
    assert.equal(await notice.count(),0,'an old failed month cannot create a warning in the current month');
   }else{
    await page.clock.runFor(10_001);
    await page.waitForFunction(()=>{const element=document.getElementById('schedule-source-status');return element&&!element.hidden;});
    assert.match(await notice.innerText(),/接続できないため、保存済み/);
    finish();
   }
  }else if(scenario==='fast'){
   await page.waitForFunction(()=>document.querySelectorAll('#teacherSelect option').length>1);
   assert.equal(await notice.count(),0,'healthy loads must not create warning markup');
   assert.equal(await notice.isVisible(),false);assert.deepEqual(await page.evaluate(()=>window.visibleScheduleWarnings),[]);
  }else{
   await page.waitForFunction(()=>{const element=document.getElementById('schedule-source-status');return element&&!element.hidden;});
   assert.match(await notice.innerText(),scenario==='unavailable'?/授業データを取得できません/:/接続できないため、保存済み/);
  }
  assert.deepEqual(errors,[]);assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);
  if(process.env.CALENDAR_SCREENSHOT_DIR&&['delayed','failed'].includes(scenario)){await fs.mkdir(process.env.CALENDAR_SCREENSHOT_DIR,{recursive:true});await page.screenshot({path:path.join(process.env.CALENDAR_SCREENSHOT_DIR,`lesson-prep-${scenario}-${width}.png`),fullPage:true});}
  await page.close();console.log(`${width}px ${scenario}: initial loading/failure notice verified`);
 }
}finally{await browser.close();server.closeAllConnections();await new Promise(resolve=>server.close(resolve));}
