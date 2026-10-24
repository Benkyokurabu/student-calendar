import assert from 'node:assert/strict';
import fs from 'node:fs/promises';import http from 'node:http';import path from 'node:path';import {fileURLToPath} from 'node:url';
import {chromium} from 'playwright';
const root=fileURLToPath(new URL('../',import.meta.url));
const server=http.createServer(async(req,res)=>{try{const f=path.resolve(root,'.'+new URL(req.url,'http://local').pathname);if(!f.startsWith(root))throw Error();res.setHeader('Content-Type',({'.html':'text/html','.js':'text/javascript','.css':'text/css'})[path.extname(f)]||'application/json');res.end(await fs.readFile(f));}catch{res.writeHead(404);res.end();}});
await new Promise(r=>server.listen(0,'127.0.0.1',r));const base=`http://127.0.0.1:${server.address().port}`;
const browser=await chromium.launch({headless:true,...(process.env.CALENDAR_TEST_CHANNEL?{channel:process.env.CALENDAR_TEST_CHANNEL}:{})});
const rows=['hon','minami'].map(c=>({date:'2026-10-23',time:'5:00～6:00',grade:'j3',class:'S',subject:'math',campus:c,room:'1',groupKey:`${c}_j3_S_math`,label:'中3S数学'}));
const key=r=>`${r.date}|${r.time}|${r.campus}|${r.groupKey}|${r.room}`;
try{for(const width of [390,1280]){
 const page=await browser.newPage({viewport:{width,height:900}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));await page.clock.install({time:new Date('2026-10-23T12:00:00+09:00')});
 let unavailable=true,scheduleRequests=0,publicationRequests=0;
 await page.route('**/*',async route=>{
  const u=new URL(route.request().url());
  if(u.pathname==='/api/schedule/calendar'){
   scheduleRequests++;if(unavailable)return new Promise(()=>{});
   return route.fulfill({json:{month:'2026-10',count:2,schedule:rows,teacherSchedule:{items:rows}}});
  }
  if(u.pathname==='/api/recordings/publication'){
   publicationRequests++;if(unavailable)return new Promise(()=>{});return route.fulfill({json:{rules:[]}});
  }
  const file=path.basename(u.pathname);
  if(file.endsWith('.json')){
   const data=file.startsWith('schedule_')?rows:file.startsWith('journal_')?{entries:Object.fromEntries(rows.map(r=>[key(r),{content:'テスト用の指導内容',homework:['テスト用の宿題'],recordingUrl:'https://example.test/journal'}]))}
    :file.startsWith('zoom_recording_urls_')?{entries:Object.fromEntries(rows.map(r=>[key(r),{url:'https://example.test/recording'}]))}:file==='recording_release.json'?{rules:[]}:{};
   return route.fulfill({json:data});
  }
  if(u.origin===base)return route.continue();return route.abort();
 });
 await page.goto(base+'/calendar.html',{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>Boolean(window.CalendarScheduleSource&&window.CalendarRequest));
 assert.equal(scheduleRequests,1);await page.clock.runFor(10_001);
 await page.locator('#subjectArea input').first().waitFor();
 assert((await page.locator('#hintArea').innerText()).includes('授業スケジュールを表示しました'));
 await page.selectOption('#gradeSelect','j3');await page.selectOption('#campusSelect','hon');await page.locator('#subjectArea input').first().check();
 await page.locator('#monthGrid .cell.has-lessons').first().click();
 assert((await page.locator('#dayDetailArea').innerText()).includes('2026-10-23'));
 assert.equal(await page.locator('#dayDetailArea a.j-link').count(),0);
 assert.equal(await page.locator('#schedule-source-status').count(),0,'student calendar omits the source connection notice');
 assert(publicationRequests>=1);await page.clock.runFor(8_001);
 await page.waitForFunction(()=>document.getElementById('hintArea').textContent.includes('取得できていません'));
 assert((await page.locator('#dayDetailArea').innerText()).includes('公開状況を確認できません'));
 assert.equal(await page.locator('#dayDetailArea a.j-link').count(),0);
  assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
  if(process.env.CALENDAR_SCREENSHOT_DIR){await fs.mkdir(process.env.CALENDAR_SCREENSHOT_DIR,{recursive:true});await page.screenshot({path:path.join(process.env.CALENDAR_SCREENSHOT_DIR,`calendar-unavailable-${width}.png`),fullPage:true});}
 unavailable=false;await page.clock.runFor(60_001);
 await page.locator('#dayDetailArea a.j-link').first().waitFor();
 assert.equal(await page.locator('#dayDetailArea a.j-link').first().getAttribute('href'),'https://example.test/recording');
 assert.equal(await page.locator('#subjectArea input').first().isChecked(),true);
 assert((await page.locator('#dayDetailArea').innerText()).includes('2026-10-23'));
 assert.equal(await page.locator('#hintArea').innerText(),'');
 assert.equal(await page.locator('#schedule-source-status').count(),0,'automatic schedule refresh must also omit the notice');
 assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));assert.deepEqual(errors,[]);
 if(process.env.CALENDAR_SCREENSHOT_DIR){await fs.mkdir(process.env.CALENDAR_SCREENSHOT_DIR,{recursive:true});await page.screenshot({path:path.join(process.env.CALENDAR_SCREENSHOT_DIR,`calendar-recovered-${width}.png`),fullPage:true});}
 await page.close();console.log(`${width}px: stalled common APIs, archive display, fail-closed recordings, recovery and selection retention passed`);
}}finally{await browser.close();server.closeAllConnections();await new Promise(r=>server.close(r));}
