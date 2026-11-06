import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import http from 'node:http';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const {chromium}=await import(process.env.CALENDAR_PLAYWRIGHT_MODULE||'playwright');
const root=fileURLToPath(new URL('../',import.meta.url)),month='2026-11';
const event=(ym=month)=>({date:ym+'-04',time:'4:55～6:15',campus:'minami',groupKey:'minami_e6_A_arith',grade:'e6',class:'A',subject:'arith',room:'1',teacher:'確認用職員'});
const other=(ym=month)=>({...event(ym),time:'6:35～8:05',groupKey:'minami_j1_A_math',grade:'j1',subject:'math'});
const key=ev=>[ev.date,ev.time,ev.campus,ev.groupKey,ev.room].join('|');
const url='https://example.test/completed-video',corrected='https://example.test/corrected-video';
const server=http.createServer(async(req,res)=>{
 try{const file=path.resolve(root,'.'+new URL(req.url,'http://local').pathname);if(!file.startsWith(root))throw Error('outside root');res.setHeader('Content-Type',path.extname(file)==='.js'?'text/javascript':'text/html');res.end(await fs.readFile(file));}
 catch{res.writeHead(404);res.end();}
});
await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
const base=`http://127.0.0.1:${server.address().port}`,browser=await chromium.launch({headless:true});
try{
 for(const width of [390,1280])for(const scenario of ['completed','missing','wrong-date','wrong-time','wrong-class','legacy','processing','hidden','override-hidden','corrected','late']){
  const page=await browser.newPage({viewport:{width,height:950}}),errors=[];page.on('pageerror',error=>errors.push(error.message));
  await page.context().route('https://example.test/**',route=>route.fulfill({contentType:'text/html',body:'<title>Recording destination</title>'}));
  await page.clock.install({time:new Date('2026-11-04T12:00:00+09:00')});
  let available=scenario!=='late';
  await page.route('**/*',async route=>{
   const requestUrl=new URL(route.request().url()),file=path.basename(requestUrl.pathname);
   if(requestUrl.pathname==='/api/schedule/calendar'){
    const ym=requestUrl.searchParams.get('month'),rows=[event(ym),other(ym)];
    return route.fulfill({json:{month:ym,count:rows.length,schedule:rows,teacherSchedule:{items:rows}}});
   }
   if(requestUrl.pathname==='/api/recordings/publication')return route.fulfill({json:{rules:scenario==='hidden'?[{eventKeys:[key(event())],status:'hidden'}]:scenario==='corrected'?[{eventKeys:[key(event())],status:'public',recordingCorrection:true,url:corrected}]:[]}});
   if(file.startsWith('zoom_recording_overrides_'))return route.fulfill({json:{month,entries:scenario==='override-hidden'?{[key(event())]:{hidden:true,url:''}}:{}}});
   if(file.startsWith('zoom_recording_urls_')){
    let matched=event();
    if(scenario==='wrong-date')matched={...matched,date:'2026-11-03'};
    if(scenario==='wrong-time')matched={...matched,time:'8:25～9:55'};
    if(scenario==='wrong-class')matched={...matched,class:'B',groupKey:'minami_e6_B_arith'};
    const record={...matched,url,...(scenario==='legacy'?{}:{recordingCompleted:scenario!=='processing'})};
    return route.fulfill({json:{month,generatedAt:'2026-11-04T09:30:00+09:00',entries:available&&scenario!=='missing'?{[key(matched)]:record}:{}}});
   }
   if(file.endsWith('.json'))return route.fulfill({json:{month,entries:{},slots:{},rules:[]}});
   return requestUrl.origin===base?route.continue():route.abort();
  });
  await page.goto(base+'/lesson_prep.html',{waitUntil:'networkidle'});
  await page.locator('#teacherSelect').selectOption('確認用職員');await page.getByRole('button',{name:'表示',exact:true}).click();
  const cards=page.locator('#output .lesson-card');assert.equal(await cards.count(),2);
  const first=cards.nth(0),second=cards.nth(1),badge=first.locator('[data-recording-available]');
  const expected=['completed','corrected'].includes(scenario);
  assert.equal(await badge.count(),expected?1:0,scenario);
  assert.equal(await second.locator('[data-recording-available]').count(),0,'a different class must not borrow the status');
  if(expected){
   assert.equal((await badge.innerText()).trim(),'動画を見る');
   assert.equal(await badge.getAttribute('href'),scenario==='corrected'?corrected:url);
   assert.equal(await badge.getAttribute('target'),'_blank');
   assert.match(await badge.getAttribute('rel'),/noopener/);
   const [popup]=await Promise.all([page.waitForEvent('popup'),badge.click()]);
   await popup.waitForLoadState();
   assert.equal(popup.url(),scenario==='corrected'?corrected:url);await popup.close();
  }
  if(scenario==='corrected')assert.equal(await first.getByRole('button',{name:'録画URLをコピー',exact:true}).getAttribute('data-url'),corrected);
  if(['legacy','processing'].includes(scenario))assert.equal(await first.getByRole('button',{name:'録画URLをコピー',exact:true}).count(),1,'existing published links remain usable without claiming completed video');
  if(['hidden','override-hidden','missing','wrong-date','wrong-time','wrong-class'].includes(scenario))assert.equal(await first.getByRole('button',{name:'録画URLをコピー',exact:true}).count(),0);
  if(scenario==='late'){
   const controls=await page.locator('.controls').boundingBox();available=true;await page.clock.runFor(31000);
   await page.waitForFunction(()=>document.querySelectorAll('[data-recording-available]').length===1);
   assert.equal(await first.getByRole('link',{name:'動画を見る',exact:true}).getAttribute('href'),url);
   assert.deepEqual(await page.locator('.controls').boundingBox(),controls,'auto refresh preserves the controls');
  }
  assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);assert.deepEqual(errors,[]);
  if(process.env.CALENDAR_SCREENSHOT_DIR&&['completed','late'].includes(scenario)){
   await fs.mkdir(process.env.CALENDAR_SCREENSHOT_DIR,{recursive:true});await page.screenshot({path:path.join(process.env.CALENDAR_SCREENSHOT_DIR,`recording-available-${scenario}-${width}.png`),fullPage:true});
  }
  await page.close();console.log(`${width}px ${scenario}: completed-video badge and existing policy verified`);
 }
}finally{await browser.close();server.closeAllConnections();await new Promise(resolve=>server.close(resolve));}
