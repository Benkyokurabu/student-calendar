import assert from 'node:assert/strict';import fs from 'node:fs';import vm from 'node:vm';
const code=fs.readFileSync(new URL('../calendar-request.js',import.meta.url),'utf8');
function setup(fetch){let expire,signal;const context={AbortController,Promise,Error,fetch:async(u,o)=>{signal=o.signal;return fetch(u,o);},setTimeout(fn){expire=fn;return 1;},clearTimeout(){}};vm.createContext(context);vm.runInContext(code,context);return {json:context.CalendarRequest.json,expire:()=>expire(),signal:()=>signal};}
{
const s=setup(async()=>({ok:true,json:async()=>({ok:true})}));assert.equal((await s.json('example')).ok,true);assert.equal(s.signal().aborted,false);
}
{
const s=setup(()=>new Promise(()=>{}));const pending=s.json('example');s.expire();await assert.rejects(pending,/時間内/);assert.equal(s.signal().aborted,true);
}
{
const s=setup(async()=>({ok:true,json:()=>new Promise(()=>{})}));const pending=s.json('example');await Promise.resolve();await Promise.resolve();s.expire();await assert.rejects(pending,/時間内/);assert.equal(s.signal().aborted,true);
}
{
const s=setup(async()=>({ok:false,status:503}));await assert.rejects(s.json('example'),/503/);
}
console.log('Calendar request: normal JSON, stalled response, stalled body and HTTP failure passed.');
