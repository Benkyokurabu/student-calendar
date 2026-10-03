import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html = fs.readFileSync(new URL('../calendar.html', import.meta.url), 'utf8');
function extract(name) {
  const start = html.indexOf(`function ${name}(`);
  let depth = 0;
  for (let i = html.indexOf('{', start); i < html.length; i++) {
    if (html[i] === '{') depth++;
    if (html[i] === '}' && --depth === 0) return html.slice(start, i + 1);
  }
  throw Error(name);
}
const key = '2026-10-02|6:35～8:05|hon|hon_j3_B_eng|1';
const ctx = {
  publicationRules: [],
  recordingReleaseRules: [{eventKeys: [key], status: 'blocked', releaseAt: '2000-01-01T00:00'}],
  buildEventKeyFromLesson: it => it.key,
  recordingOverrides: {},
  findRecordingRecordForLesson: () => ({url: 'https://example.test/override'}),
  recordingUrlFromRecord: rec => rec?.url || '',
  zoomRecordingEntries: {[key]: {url: 'https://example.test/recording'}},
};
vm.createContext(ctx);
vm.runInContext(extract('isRecordingBlocked') + extract('findRecordingUrlForLesson'), ctx);
assert.equal(ctx.findRecordingUrlForLesson({key}, {recordingUrl: 'https://example.test/journal'}), '');
assert.equal(ctx.isRecordingBlocked({key: key + '2'}), false);
ctx.recordingReleaseRules[0].status = 'released';
assert.equal(ctx.findRecordingUrlForLesson({key}, {}), 'https://example.test/override');
ctx.recordingReleaseRules = [];
assert.equal(ctx.findRecordingUrlForLesson({key}, {}), 'https://example.test/override');
ctx.publicationRules=[{eventKeys:[key],status:'hidden',url:''}];
assert.equal(ctx.findRecordingUrlForLesson({key}, {recordingUrl:'https://example.test/old'}),'');
ctx.publicationRules[0]={eventKeys:[key],status:'public',url:'https://example.test/restored'};
ctx.recordingReleaseRules=[{eventKeys:[key],status:'blocked'}];
ctx.zoomRecordingEntries={};
assert.equal(ctx.findRecordingUrlForLesson({key}, {}),'https://example.test/restored');
// Automatic test matches use each campus's own date and cover every recording
// segment in that lesson. The server has already combined the campus checks.
const hon='2026-10-01|18:00～19:00|hon|hon_j1_S_math|2';
const minami='2026-10-05|18:00～19:00|minami|minami_j1_S_math|1';
ctx.recordingReleaseRules=[];
ctx.publicationRules=[{eventKeys:[],status:'hidden',match:{date:'2026-10-01',campus:'hon',group:'hon_j1_S_math'}},{eventKeys:[],status:'hidden',match:{date:'2026-10-05',campus:'minami',group:'minami_j1_S_math'}}];
for(const lesson of [hon,minami])assert.equal(ctx.findRecordingUrlForLesson({key:lesson},{recordingUrl:'https://example.test/original'}),'');
assert.equal(ctx.isRecordingBlocked({key:hon.replace('10-01','10-08')}),false);
ctx.publicationRules.forEach(rule=>{rule.status='public';rule.url='https://example.test/released';});
for(const lesson of [hon,minami])assert.equal(ctx.findRecordingUrlForLesson({key:lesson},{}),'https://example.test/released');
// All inline JavaScript must parse, including the teacher settings handler.
for (const page of ['calendar.html', 'lesson_prep.html']) {
  const source = fs.readFileSync(new URL('../' + page, import.meta.url), 'utf8');
  for (const match of source.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) new vm.Script(match[1]);
}
console.log('Recording release gate: blocked overrides/journal links, confirmed release, unchanged normal links, and page syntax passed.');
