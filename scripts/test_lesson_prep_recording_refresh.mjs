import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../lesson_prep.html',import.meta.url),'utf8');
function extract(name){const start=html.indexOf(`function ${name}(`),asyncStart=html.lastIndexOf('async ',start);let depth=0;assert.ok(start>=0,name);for(let i=html.indexOf('{',start);i<html.length;i++){if(html[i]==='{')depth++;if(html[i]==='}'&&--depth===0)return html.slice(asyncStart===start-6?asyncStart:start,i+1);}throw Error(name);}
const month='2026-11',key='2026-11-19|4:55～6:15|hon|hon_e4_A_arith|3';
const payload=(entries={[key]:{date:'2026-11-19',url:'https://example.test/video',recordingCompleted:true}},generatedAt='2026-11-19T09:25:00Z')=>({month,generatedAt,entries});
function context(){
 let current=month,rendered=0;
 const ctx={Map,Date,JSON,rawJsonUrl:file=>'https://example.test/'+file,zoomRecordingData:{},zoomRecordingMonths:new Map(),responses:[],
  loadNewestJSON:async()=>ctx.responses.shift()??null,document:{getElementById:()=>({get value(){return current+'-19'}})},render:()=>rendered++};
 ctx.bentanPublication = loader => loader(["https://g10.example.invalid/api/recordings/publication"]);
 vm.createContext(ctx);
 const names=['inferMonthFromRecordingEntries','mergeRecordingDataForMonth','loadZoomRecordings','refreshZoomRecordingsIfChanged'];
 if(html.includes('function recordingSnapshotForMonth('))names.unshift('recordingSnapshotForMonth');
 vm.runInContext(names.map(extract).join('\n'),ctx);
 return {ctx,select:value=>{current=value},renders:()=>rendered};
}
test('a failed automatic refresh preserves the fetched URL and does not redraw the sheet',async()=>{
 const {ctx,renders}=context();ctx.responses=[payload(),null,{entries:{}},{rules:[]}];ctx.zoomRecordingData=await ctx.loadZoomRecordings(month);
 ctx.responses=[null,null,null,null];assert.equal(await ctx.refreshZoomRecordingsIfChanged(),false);
 assert.equal(ctx.zoomRecordingData.entries[key].url,'https://example.test/video');assert.equal(renders(),0);
});
test('visibility failures keep existing restrictions; new restrictions apply even when URL fetch fails',async()=>{
 const {ctx}=context(),rule={eventKeys:[key],status:'hidden'};
 ctx.responses=[payload(),null,{entries:{[key]:{hidden:true}}},{rules:[rule]}];await ctx.loadZoomRecordings(month);
 ctx.responses=[null,null,null,null];const retained=await ctx.loadZoomRecordings(month);
 assert.equal(retained.hiddenEntries[key].hidden,true);assert.equal(retained.publicationRules[0].status,'hidden');
 ctx.responses=[null,null,{entries:{}},{rules:[]}];const cleared=await ctx.loadZoomRecordings(month);assert.equal(cleared.publicationRules.length,0);assert.equal(Object.keys(cleared.hiddenEntries).length,0);
 ctx.responses=[null,null,null,{rules:[rule]}];assert.equal((await ctx.loadZoomRecordings(month)).publicationRules[0].status,'hidden');
});
test('other months and malformed payloads never replace the accepted month',async()=>{
 const {ctx}=context();ctx.responses=[payload(),null,null,null];await ctx.loadZoomRecordings(month);
 ctx.responses=[{month,entries:[]},{month:'2026-12',entries:{}},null,null];const retained=await ctx.loadZoomRecordings(month);assert.ok(retained.entries[key]);
 ctx.responses=[null,payload(),null,null];const next=await ctx.loadZoomRecordings('2026-12');assert.equal(Object.keys(next.entries).length,0);assert.equal(next.month,'2026-12');
});
test('a successful newer empty dataset clears URLs; a delayed older dataset cannot roll them back',async()=>{
 const {ctx}=context();ctx.responses=[payload(),null,null,null];await ctx.loadZoomRecordings(month);
 ctx.responses=[payload({},'2026-11-19T10:00:00Z'),null,null,null];assert.equal(Object.keys((await ctx.loadZoomRecordings(month)).entries).length,0);
 ctx.responses=[payload(),null,null,null];assert.equal(Object.keys((await ctx.loadZoomRecordings(month)).entries).length,0);
});
test('a response arriving after the teacher changes month cannot replace the displayed month',async()=>{
 const {ctx,select,renders}=context();let finish;ctx.loadZoomRecordings=()=>new Promise(resolve=>{finish=resolve});ctx.zoomRecordingData={month:'2026-12',entries:{}};
 const pending=ctx.refreshZoomRecordingsIfChanged();select('2026-12');finish(payload());assert.equal(await pending,false);assert.equal(ctx.zoomRecordingData.month,'2026-12');assert.equal(renders(),0);
});
