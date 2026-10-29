import assert from 'node:assert/strict';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {chromium} from 'playwright';
const root=fileURLToPath(new URL('../',import.meta.url));
// Use the real child origin with local files, avoiding local-network permission
// differences while exercising the same cross-origin message contract as production.
const childBase='https://benkyokurabu.github.io/student-calendar';
const browser=await chromium.launch({headless:true,...(process.env.CALENDAR_TEST_CHANNEL?{channel:process.env.CALENDAR_TEST_CHANNEL}:{})});
const event=(teacher,date='2026-11-02')=>({teacher,date,time:'18:00～19:00',campus:'hon',groupKey:'hon_j1_A_math',grade:'j1',class:'A',subject:'math',room:'1'});
try{
 for(const scenario of [
  {label:'today',teacher:'工藤',rows:[event('工藤'),event('金城')],expected:'工藤'},
  {label:'alias',teacher:'佐藤（治）',rows:[event('治彦')],expected:'治彦'},
  {label:'no-lessons',teacher:'金城',rows:[event('工藤')],expected:''},
  {label:'unknown',teacher:'未確認',rows:[event('工藤')],expected:''},
  {label:'ambiguous',teacher:'髙山',rows:[event('髙山'),event('高山')],expected:''},
  {label:'invalid-date',teacher:'工藤',date:'2026-11-00',rows:[event('工藤')],expected:''},
  {label:'wrong-parent',teacher:'工藤',rows:[event('工藤')],expected:'',untrusted:true},
  {label:'standalone',teacher:'工藤',rows:[event('工藤')],expected:'',standalone:true},
 ]){
  const page=await browser.newPage({viewport:{width:390,height:844}}),errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  await page.clock.install({time:new Date('2026-11-01T12:00:00+09:00')});
  await page.addInitScript(()=>{window.printCalls=0;window.print=()=>{window.printCalls++;};});
  await page.route('**/*',async route=>{
   const url=new URL(route.request().url()),file=path.basename(url.pathname);
   if(url.pathname==='/embed-test'){
    const child=new URL(childBase+'/lesson_prep.html');child.search=new URLSearchParams({embedded:'1',teacher:scenario.teacher,date:scenario.date||'2026-11-02'}).toString();
    return route.fulfill({contentType:'text/html; charset=utf-8',body:`<!doctype html><meta name="viewport" content="width=device-width"><script>window.messages=[];addEventListener('message',e=>messages.push(e.data));</script><iframe title="授業準備" src="${child.href.replaceAll('&','&amp;')}" style="width:100%;height:650px;border:0"></iframe>`});
   }
   if(url.pathname==='/api/schedule/calendar'){
    const month=url.searchParams.get('month'),rows=scenario.rows.map(row=>({...row,date:month+'-02'}));
    return route.fulfill({json:{month,count:rows.length,schedule:rows,teacherSchedule:{items:rows}}});
   }
   if(url.pathname==='/api/recordings/publication')return route.fulfill({json:{rules:[]}});
   if(file.endsWith('.json'))return route.fulfill({json:{month:'2026-11',entries:{},slots:{},rules:[]}});
   if(url.href.startsWith(childBase+'/'))return route.fulfill({path:path.join(root,url.pathname.slice('/student-calendar/'.length))});
   return route.abort();
  });
  await page.goto(scenario.standalone?childBase+'/lesson_prep.html?embedded=1&teacher=工藤&date=2026-11-02':scenario.untrusted?'https://untrusted.example/embed-test':'http://localhost:31888/embed-test',{waitUntil:'domcontentloaded'});
  const frame=scenario.standalone?page.mainFrame():page.frames().find(frame=>frame.url().startsWith(childBase+'/lesson_prep.html'))||await page.waitForEvent('framenavigated',{predicate:frame=>frame.url().startsWith(childBase+'/lesson_prep.html')});
  assert(frame);
  await frame.waitForFunction(()=>document.querySelectorAll('#teacherSelect option').length>1);
  assert.equal(await frame.locator('#datePicker').inputValue(),scenario.standalone||scenario.untrusted||scenario.date?'2026-11-01':'2026-11-02');
  if(!scenario.standalone&&!scenario.untrusted){
   await page.waitForFunction(()=>window.messages.some(message=>message.type==='lesson-prep:ready'));
   const ready=await page.evaluate(()=>window.messages.find(message=>message.type==='lesson-prep:ready'));
   assert.equal(ready.hasLessons,!!scenario.expected);assert(Number.isFinite(ready.height)&&ready.height>=100);
   assert.equal(ready.matchedTeacher,scenario.expected||undefined);
  }else if(scenario.untrusted)assert.deepEqual(await page.evaluate(()=>window.messages),[]);
  assert.equal(await frame.locator('#teacherSelect').inputValue(),scenario.expected);
  assert.equal(await frame.locator('#output .lesson-card').count(),scenario.expected?1:0);
  if(scenario.label==='today'){
   await frame.evaluate(()=>{
    window.dispatchEvent(new MessageEvent('message',{origin:'https://untrusted.example',source:parent,data:{type:'lesson-prep:print'}}));
    window.dispatchEvent(new MessageEvent('message',{origin:new URL(document.referrer).origin,source:window,data:{type:'lesson-prep:print'}}));
   });
   assert.equal(await frame.evaluate(()=>window.printCalls),0,'untrusted origin or wrong source must not print');
   await page.evaluate(origin=>document.querySelector('iframe').contentWindow.postMessage({type:'lesson-prep:print'},origin),childBase);
   await frame.waitForFunction(()=>window.printCalls===1);
   await frame.locator('#teacherSelect').selectOption('金城');
   await page.waitForFunction(()=>window.messages.filter(message=>message.type==='lesson-prep:availability').at(-1)?.hasLessons===false);
   await page.evaluate(origin=>document.querySelector('iframe').contentWindow.postMessage({type:'lesson-prep:print'},origin),childBase);
   assert.equal(await frame.evaluate(()=>window.printCalls),1,'a changed selection must not print the previous teacher sheet');
   await frame.getByRole('button',{name:'表示',exact:true}).click();
   await page.waitForFunction(()=>window.messages.filter(message=>message.type==='lesson-prep:availability').at(-1)?.matchedTeacher==='金城');
   assert.equal(await frame.locator('#teacherSelect').inputValue(),'金城','the initial query must never override a later user choice');
   await frame.getByRole('button',{name:'印刷',exact:true}).click();
   assert.equal(await frame.evaluate(()=>window.printCalls),2,'existing in-sheet print button is preserved');
  }
  assert.deepEqual(errors,[]);
  await page.close();console.log(`Lesson prep embed ${scenario.label}: selection, availability and origin isolation passed`);
 }
}finally{await browser.close();}
