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
// All inline JavaScript must parse, including the teacher settings handler.
for (const page of ['calendar.html', 'lesson_prep.html']) {
  const source = fs.readFileSync(new URL('../' + page, import.meta.url), 'utf8');
  for (const match of source.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) new vm.Script(match[1]);
}
console.log('Recording release gate: blocked overrides/journal links, confirmed release, unchanged normal links, and page syntax passed.');
