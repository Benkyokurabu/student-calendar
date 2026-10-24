import assert from 'node:assert/strict';
import fs from 'node:fs/promises';import http from 'node:http';import path from 'node:path';import {fileURLToPath} from 'node:url';
import {chromium} from 'playwright';
const root=fileURLToPath(new URL('../',import.meta.url));
const approved=JSON.parse(await fs.readFile(path.join(root,'recording_emergency_2026-10.json'),'utf8'));
const key=Object.keys(approved.entries).at(-1),[date,time,campus,groupKey,room]=key.split('|');
const [,grade,klass,subject]=groupKey.split('_');const row={date,time,campus,groupKey,room,grade,class:klass,subject,label:'動作確認用'};
const server=http.createServer(async(req,res)=>{try{const f=path.resolve(root,'.'+new URL(req.url,'http://local').pathname);if(!f.startsWith(root))throw Error();res.setHeader('Content-Type',({'.html':'text/html','.js':'text/javascript','.css':'text/css'})[path.extname(f)]||'application/json');res.end(await fs.readFile(f));}catch{res.writeHead(404);res.end();}});
await new Promise(r=>server.listen(0,'127.0.0.1',r));const base=`http://127.0.0.1:${server.address().port}`;
const browser=await chromium.launch({headless:true});
try{for(const width of [390,1280]){
 const page=await browser.newPage({viewport:{width,height:900}});const errors=[];page.on('pageerror',e=>errors.push(e.message));let live=false;
 await page.clock.install({time:new Date('2026-10-24T14:00:00+09:00')});
 await page.route('**/*',route=>{
  const u=new URL(route.request().url());
  if(u.pathname==='/api/schedule/calendar')return route.fulfill({status:503,json:{error:'offline'}});
  if(u.pathname==='/api/recordings/publication')return route.fulfill(live?{json:{rules:[{eventKeys:[key],status:'hidden',url:''}]}}:{status:503,json:{error:'offline'}});
  if(path.basename(u.pathname).startsWith('schedule_')&&u.pathname.endsWith('.json'))return route.fulfill({json:[row]});
  if(u.hostname==='raw.githubusercontent.com')return route.abort();
  if(u.origin===base)return route.continue();return route.abort();
 });
 await page.goto(base+'/calendar.html',{waitUntil:'networkidle'});
 await page.selectOption('#gradeSelect',grade);await page.selectOption('#campusSelect',campus);
 await page.locator('#subjectArea input').first().check();
 const classes=page.locator('#classArea input');if(await classes.count())for(const el of await classes.all())await el.check();
 await page.locator('#monthGrid .cell.has-lessons').first().click();
 const link=page.locator('#dayDetailArea a.j-link');await link.first().waitFor();assert.equal(await link.first().getAttribute('href'),approved.entries[key]);
 assert((await page.locator('#hintArea').innerText()).includes('保存済みのURL'));assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));assert.deepEqual(errors,[]);
 await fs.mkdir('analysis_outputs',{recursive:true});await page.screenshot({path:`analysis_outputs/zoom-emergency-${width}.png`,fullPage:true});
 live=true;await page.clock.runFor(30_001);await page.waitForFunction(()=>document.querySelectorAll('#dayDetailArea a.j-link').length===0);
 assert((await page.locator('#dayDetailArea').innerText()).includes('公開待ち'));assert.equal(await page.locator('#subjectArea input').first().isChecked(),true);
 await page.close();console.log(`${width}px: existing published URL visible during API outage; restored live restriction takes priority`);
}}finally{await browser.close();server.closeAllConnections();await new Promise(r=>server.close(r));}
