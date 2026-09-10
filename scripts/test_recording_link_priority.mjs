import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
function extract(source,name){const start=source.indexOf(`function ${name}(`);assert(start>=0);let n=0;for(let i=source.indexOf('{',start);i<source.length;i++){if(source[i]==='{')n++;if(source[i]==='}'&&--n===0)return source.slice(start,i+1);}throw Error(name);}
const key='2026-09-10|6:35～8:05|hon|hon_j3_B_eng|1';
for(const page of ['calendar.html','calendar_journal.html']){
 const html=fs.readFileSync(new URL('../'+page,import.meta.url),'utf8');
 const ctx={zoomRecordingEntries:{[key]:{url:'https://zoom.us/today-hon'}},recordingOverrides:{},override:null,buildEventKeyFromLesson:it=>it.key,findRecordingRecordForLesson:()=>ctx.override,recordingUrlFromRecord:r=>r?.url||''};vm.createContext(ctx);vm.runInContext(extract(html,'findRecordingUrlForLesson'),ctx);
 const resolve=(key,old='https://zoom.us/stale-july')=>ctx.findRecordingUrlForLesson({key},{recordingUrl:old});
 assert.equal(resolve(key),'https://zoom.us/today-hon',page+' stale journal must not win');
 assert.equal(resolve(key.replaceAll('hon','minami'),'https://example.com/manual'),'https://example.com/manual',page+' no cross-campus borrowing');
 assert.equal(resolve(key.replace('2026-09-10','2026-09-11'),''),'',page+' no borrowing another date');
 assert.equal(resolve(key.replace(/\|1$/,'|2'),''),'',page+' no borrowing another room');
 if(page==='calendar.html'){ctx.override={url:''};assert.equal(resolve(key),'','explicit suppression preserved');ctx.override={url:'https://example.com/override'};assert.equal(resolve(key),'https://example.com/override');}
 // Check the actual incident's selected link with repository data.
 ctx.override=null;ctx.zoomRecordingEntries=JSON.parse(fs.readFileSync(new URL('../zoom_recording_urls_2026-09.json',import.meta.url),'utf8')).entries;
 const entry=JSON.parse(fs.readFileSync(new URL('../journal_2026-09.json',import.meta.url),'utf8')).entries[key];
 assert.equal(ctx.findRecordingUrlForLesson({key},entry),ctx.zoomRecordingEntries[key].url);
 console.log(page+': exact lesson recording priority verified');
}
