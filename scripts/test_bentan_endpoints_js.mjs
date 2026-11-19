import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import vm from 'node:vm';
const root=new URL('../',import.meta.url),source=await fs.readFile(new URL('bentan-endpoints.js',root),'utf8');
function fixture(manifest,{ok=true,reject=false}={}){
 const window={},requests=[];
 const fetch=async(url,options)=>{requests.push({url:String(url),options});if(reject)throw Error('offline');return {ok,json:async()=>manifest};};
 vm.runInNewContext(source,{window,document:{currentScript:{src:'https://pages.example.invalid/student-calendar/bentan-endpoints.js'}},URL,fetch});return {window,requests};
}
test('Pages policy and private staff navigation use the same source manifest',async()=>{
 const {window,requests}=fixture({version:1,apiBaseUrl:'https://g10.example.invalid',staffBaseUrl:'https://g10.vpn.example.invalid'});
 assert.equal(await window.bentanApiUrl('/api/recordings/publication'),'https://g10.example.invalid/api/recordings/publication');
 assert.equal(await window.bentanStaffUrl('/staff/recordings?month=2026-11'),'https://g10.vpn.example.invalid/staff/recordings?month=2026-11');
 assert.equal(requests.length,1);assert.equal(requests[0].options.cache,'no-store');
 await assert.rejects(window.bentanApiUrl('/rest/v1/students'));
});
test('Unavailable or invalid manifest remains unavailable; no old-origin fallback',async()=>{
 for(const setup of [fixture({}, {reject:true}),fixture({version:1,apiBaseUrl:'https://user:secret@example.invalid',staffBaseUrl:'https://staff.example.invalid'})]){
  let called=false;assert.equal(await setup.window.bentanPublication(async()=>{called=true;return {rules:[]};}),null);assert.equal(called,false);
 }
});
test('Staff navigation can use this PC loopback; public API cannot use HTTP',async()=>{
 const {window}=fixture({version:1,apiBaseUrl:'https://g10.example.invalid',staffBaseUrl:'http://127.0.0.1:3000'});
 assert.equal(await window.bentanStaffUrl('/staff/recordings?month=2026-11'),'http://127.0.0.1:3000/staff/recordings?month=2026-11');
 const rejected=fixture({version:1,apiBaseUrl:'http://127.0.0.1:3000',staffBaseUrl:'http://127.0.0.1:3000'});
 await assert.rejects(rejected.window.bentanApiUrl('/api/recordings/publication'));
});
test('Only the configured approved zrok public GET origin receives the API header',async()=>{
 const origin='https://bentang10ops20261119.shares.zrok.io';
 const z=fixture({version:1,apiBaseUrl:origin,staffBaseUrl:'http://127.0.0.1:3000'}).window;
 await z.bentanApiUrl('/api/recordings/publication');
 const options=z.bentanApiFetchOptions(origin+'/api/recordings/publication?_=123',{cache:'no-store'});
 assert.equal(options.headers.skip_zrok_interstitial,'true');assert.equal(options.credentials,'omit');
 for(const url of ['https://other.shares.zrok.io/api/recordings/publication',origin+'/api/recordings/capture','http://bentang10ops20261119.shares.zrok.io/api/recordings/publication'])assert.equal(z.bentanApiFetchOptions(url,{cache:'no-store'}).headers,undefined);
 const legacy=fixture({version:1,apiBaseUrl:'https://line-check-system.vercel.app',staffBaseUrl:'https://line-check-system.vercel.app'}).window;
 await legacy.bentanApiUrl('/api/recordings/publication');
 assert.equal(legacy.bentanApiFetchOptions('https://line-check-system.vercel.app/api/recordings/publication',{}).headers,undefined);
 assert.equal(legacy.bentanApiFetchOptions(origin+'/api/recordings/publication',{}).headers,undefined);
});
test('Interstitial HTML, invalid JSON and missing rules never become an authoritative policy',async()=>{
 const {window}=fixture({version:1,apiBaseUrl:'https://bentang10ops20261119.shares.zrok.io',staffBaseUrl:'http://127.0.0.1:3000'});
 for(const payload of ['<html>zrok</html>',{nonce:'fixture'}])assert.equal(await window.bentanPublication(async()=>payload),null);
 assert.equal(await window.bentanPublication(async()=>{throw SyntaxError('invalid JSON');}),null);
 assert.deepEqual(await window.bentanPublication(async()=>({rules:[]})),{rules:[]});
});
test('Both existing pages retain parseable inline code and use shared publication loader',async()=>{
 for(const name of ['calendar.html','lesson_prep.html']){
  const html=await fs.readFile(new URL(name,root),'utf8');
  assert.match(html,/bentan-endpoints\.js/);assert.match(html,/bentanPublication\(/);
  assert.doesNotMatch(html,/https:\/\/line-check-system\.vercel\.app\/api\/recordings\/publication/);
  for(const match of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g))if(!/\bsrc=/.test(match[1]))new vm.Script(match[2],{filename:name});
 }
});
