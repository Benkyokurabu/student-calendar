import assert from 'node:assert/strict';
import fs from 'node:fs/promises';import http from 'node:http';import path from 'node:path';import {fileURLToPath} from 'node:url';
import {chromium} from 'playwright';
const root=fileURLToPath(new URL('../',import.meta.url));
const approved=JSON.parse(await fs.readFile(path.join(root,'recording_emergency_2026-10.json'),'utf8'));
const key=Object.keys(approved.entries).at(-1),[date,time,campus,groupKey,room]=key.split('|');
const [,grade,klass,subject]=groupKey.split('_');const row={date,time,campus,groupKey,room,grade,class:klass,subject,label:'速度確認用'};
const server=http.createServer(async(req,res)=>{try{const f=path.resolve(root,'.'+new URL(req.url,'http://local').pathname);if(!f.startsWith(root))throw Error();res.setHeader('Content-Type',({'.html':'text/html','.js':'text/javascript','.css':'text/css'})[path.extname(f)]||'application/json');res.end(await fs.readFile(f));}catch{res.writeHead(404);res.end();}});
await new Promise(r=>server.listen(0,'127.0.0.1',r));const base=`http://127.0.0.1:${server.address().port}`;
const browser=await chromium.launch({headless:true});
try{for(const width of [390,1280]){
 const page=await browser.newPage({viewport:{width,height:900}});const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.clock.install({time:new Date('2026-10-24T14:00:00+09:00')});
 await page.route('**/*',async route=>{
  const u=new URL(route.request().url());
  if(u.pathname==='/api/schedule/calendar'){await new Promise(r=>setTimeout(r,6000));return route.fulfill({status:503,json:{error:'offline'}}).catch(()=>{});}
  if(u.pathname==='/api/recordings/publication'){await new Promise(r=>setTimeout(r,4500));return route.fulfill({json:{rules:[{eventKeys:[key],status:'hidden',url:''}]}}).catch(()=>{});}
  if(path.basename(u.pathname).startsWith('schedule_')&&u.pathname.endsWith('.json'))return route.fulfill({json:[row]});
  if(u.hostname==='raw.githubusercontent.com')return route.abort();
  if(u.origin===base)return route.continue();return route.abort();
 });
 const start=Date.now();await page.goto(base+'/calendar.html',{waitUntil:'domcontentloaded'});
 await page.waitForFunction(()=>document.querySelector('#monthGrid')?.children.length>0);
 const scheduleMs=Date.now()-start;assert(scheduleMs<3500,`schedule waited ${scheduleMs}ms for the failed API`);
 await page.selectOption('#gradeSelect',grade);await page.selectOption('#campusSelect',campus);
 await page.locator('#subjectArea input').first().check();
 for(const el of await page.locator('#classArea input').all())await el.check();
 await page.locator('#monthGrid .cell.has-lessons').first().click();
 const link=page.locator('#dayDetailArea a.j-link');await link.first().waitFor();
 const recordingMs=Date.now()-start;assert(recordingMs<5000,`recording waited ${recordingMs}ms for publication API`);
 assert.equal(await link.first().getAttribute('href'),approved.entries[key]);
 await page.waitForFunction(()=>document.querySelectorAll('#dayDetailArea a.j-link').length===0,{},{timeout:10000});
 assert((await page.locator('#dayDetailArea').innerText()).includes('公開待ち'),'late live restrictions must override the outage snapshot');
 assert.deepEqual(errors,[]);assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
 console.log(`${width}px: schedule ${scheduleMs}ms, recording ${recordingMs}ms; late hidden rule applied`);await page.close();
}}finally{await browser.close();server.closeAllConnections();await new Promise(r=>server.close(r));}
