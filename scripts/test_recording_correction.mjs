import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
function extract(html,name){const start=html.indexOf(`function ${name}(`);assert(start>=0);let depth=0;for(let i=html.indexOf('{',start);i<html.length;i++){if(html[i]==='{')depth++;if(html[i]==='}'&&--depth===0)return html.slice(start,i+1);}throw Error(name);}
const key='2026-10-29|18:35|hon|hon_j3_S_math|1',original='https://zoom.us/rec/play/original',url='https://zoom.us/rec/share/corrected';
const rule={eventKeys:[key],status:'public',recordingCorrection:true,url};
const ctx={publicationAvailable:true,publicationRules:[rule],recordingReleaseRules:[],recordingOverrides:{},zoomRecordingEntries:{[key]:{url:original}},buildEventKeyFromLesson:row=>row.key,recordingUrlFromRecord:row=>row?.url || '',findRecordingRecordForLesson:()=>null};
vm.createContext(ctx);const student=fs.readFileSync(new URL('../calendar.html',import.meta.url),'utf8');
vm.runInContext(['isRecordingBlocked','findRecordingUrlForLesson'].map(name=>extract(student,name)).join('\n'),ctx);
assert.equal(ctx.findRecordingUrlForLesson({key},{}),url);
ctx.recordingReleaseRules=[{eventKeys:[key],status:'blocked'}];assert.equal(ctx.findRecordingUrlForLesson({key},{}),'');ctx.recordingReleaseRules=[];
ctx.recordingOverrides[key]={hidden:true,url:''};assert.equal(ctx.findRecordingUrlForLesson({key},{}),'');ctx.recordingOverrides={};
rule.status='hidden';assert.equal(ctx.findRecordingUrlForLesson({key},{}),'');rule.status='public';
const teacher={zoomRecordingData:{entries:{[key]:{url:original}},publicationRules:[rule],hiddenEntries:{}},makeKey:row=>row.key};
vm.createContext(teacher);vm.runInContext(extract(fs.readFileSync(new URL('../lesson_prep.html',import.meta.url),'utf8'),'findRecordingForLesson'),teacher);
assert.equal(teacher.findRecordingForLesson({key}).url,url);
rule.status='hidden';assert.equal(teacher.findRecordingForLesson({key}),null);rule.status='public';
teacher.zoomRecordingData.hiddenEntries[key]={hidden:true,url:''};assert.equal(teacher.findRecordingForLesson({key}),null);
console.log('Recording corrections: student and teacher links agree; automatic URL, private policy and local suppression verified');
