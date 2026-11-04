import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../calendar.html',import.meta.url),'utf8');
const start=html.indexOf('  let recordingRefreshPending = false;');
const end=html.indexOf('\n  gradeSelect.addEventListener',start);
const normalizeStart=html.indexOf('  function normalizeRecordingEntries(');
const normalizeEnd=html.indexOf('\n  async function loadRecordingsSafe(',normalizeStart);
assert(start>=0&&end>start&&normalizeStart>=0&&normalizeEnd>normalizeStart);
const key='2026-11-02|6:35～8:05|minami|minami_j1_S_math|1';
let tick,rendered=0,monthly=null,latest=null;
const ctx={document:{hidden:false},currentYM:'2026-11',recordingsLoadedYM:'2026-11',publicationRules:[],publicationAvailable:false,journalAvailable:true,journalEntries:{},emergencyRecordings:null,zoomRecordingEntries:{},
 rawJsonUrl:name=>'raw/'+name,setInterval:fn=>{tick=fn;},renderAll:()=>{rendered++;},updateDetailNotice:()=>{},loadJournalSafe:async()=>{},
 fetchNewestJsonAny:async names=>names[0].includes('/publication')?{rules:[{status:'hidden',eventKeys:[key]}]}:names[0].startsWith('recording_emergency_')?null:names[0]==='zoom_recording_urls_latest.json'?latest:monthly};
vm.createContext(ctx);vm.runInContext(html.slice(normalizeStart,normalizeEnd)+html.slice(start,end),ctx);
await tick();assert.equal(rendered,1);assert.equal(Object.keys(ctx.zoomRecordingEntries).length,0);
monthly={month:'2026-11',entries:{[key]:{url:'https://example.test/new-recording'}}};
await tick();assert.equal(ctx.zoomRecordingEntries[key].url,'https://example.test/new-recording');assert.equal(rendered,2);
assert.equal(ctx.publicationRules[0].status,'hidden','new links must not relax a publication restriction');
await tick();assert.equal(rendered,2,'unchanged links must not rebuild the view');
monthly=null;latest={month:'2026-12',entries:{future:{url:'https://example.test/future'}}};
await tick();assert.equal(ctx.zoomRecordingEntries[key].url,'https://example.test/new-recording');assert(!ctx.zoomRecordingEntries.future);
let release;const gate=new Promise(resolve=>{release=resolve;});
ctx.fetchNewestJsonAny=async()=>{await gate;return {month:'2026-11',entries:{late:{url:'https://example.test/late'}}};};
const pending=tick();ctx.currentYM='2026-12';release();await pending;assert(!ctx.zoomRecordingEntries.late);
console.log('Open-calendar recording refresh: newly published links, unchanged view, failure retention, hidden rules and month race passed');
