import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const root=new URL('../',import.meta.url),html=fs.readFileSync(new URL('calendar.html',root),'utf8');
function extract(name){const start=html.indexOf(`function ${name}(`);assert(start>=0);let depth=0;for(let i=html.indexOf('{',start);i<html.length;i++){if(html[i]==='{')depth++;if(html[i]==='}'&&--depth===0)return html.slice(start,i+1);}throw Error(name);}
const approved=JSON.parse(fs.readFileSync(new URL('recording_emergency_2026-10.json',root),'utf8'));
const source=JSON.parse(fs.readFileSync(new URL('zoom_recording_urls_2026-10.json',root),'utf8')).entries;
class TestDate extends Date {static now(){return Date.parse(approved.approvedAt)+1000;}}
const ctx={URL,Date:TestDate,publicationAvailable:false,publicationRules:[],emergencyRecordings:approved,recordingReleaseRules:[],zoomRecordingEntries:source,recordingOverrides:{},buildEventKeyFromLesson:it=>it.key};
vm.createContext(ctx);vm.runInContext(['emergencyRecordingUrl','isRecordingBlocked','findRecordingUrlForLesson'].map(extract).join('\n'),ctx);
let count=0;
for(const [key,row] of Object.entries(source)){
 const result=ctx.findRecordingUrlForLesson({key},{recordingUrl:'https://example.test/unsafe-journal'});
 if(row.hidden||!approved.entries[key])assert.equal(result,'');else{assert.equal(result,approved.entries[key]);count++;}
}
assert(Object.keys(approved.entries).length>0);assert.equal(count,Object.keys(approved.entries).length);
const key=Object.keys(approved.entries)[0],saved=approved.entries[key];
for(const other of [key.replace('2026-10','2026-09'),key.replace(/\|\d+$/,'|99'),key.replace('|hon|','|minami|')])if(other!==key&&!approved.entries[other])assert.equal(ctx.findRecordingUrlForLesson({key:other},{}),'');
ctx.recordingReleaseRules=[{eventKeys:[key],status:'blocked'}];assert.equal(ctx.findRecordingUrlForLesson({key},{}),'');ctx.recordingReleaseRules=[];
approved.entries[key]='https://zoom.us.evil.test/rec/play/a';assert.equal(ctx.findRecordingUrlForLesson({key},{}),'');approved.entries[key]=saved;
const expiry=approved.expiresAt;approved.expiresAt='2000-01-01T00:00:00Z';assert.equal(ctx.findRecordingUrlForLesson({key},{}),'');approved.expiresAt=expiry;
ctx.publicationAvailable=true;ctx.publicationRules=[{eventKeys:[key],status:'hidden'}];assert.equal(ctx.findRecordingUrlForLesson({key},{}),'');
console.log(`${count} already-public recording URLs approved; hidden, expired, mismatched and live-blocked recordings remain unavailable.`);
