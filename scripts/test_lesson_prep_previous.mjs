import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';

const html = readFileSync(new URL('../lesson_prep.html', import.meta.url), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)?.[1];
assert.ok(script, '授業準備シートのスクリプトが見つかりません');

const elements = {
  datePicker: { value: '2026-09-21', addEventListener() {} },
  teacherSelect: { value: '担当' },
  output: { innerHTML: '', className: '' },
};
const document = {
  getElementById: id => elements[id],
  addEventListener() {},
  createElement() {
    return {
      set textContent(value) { this.text = String(value); },
      get innerHTML() { return this.text ?? ''; },
    };
  },
  body: { classList: { add() {} } },
};
const context = {
  document,
  location: { search: '' },
  localStorage: { getItem: () => null },
  URLSearchParams,
  fixture: null,
};
runInNewContext(script.replace(/\/\* ====== 起動 ====== \*\/[\s\S]*$/, ''), context);

const group = 'hon_j1_A_math';
const pairGroup = 'minami_j1_A_math';
const key = (date, gk = group) => `${date}|16:00～17:20|hon|${gk}|1`;
const event = (date, extra = {}) => ({
  date, time: '16:00～17:20', campus: 'hon', groupKey: group, room: '1',
  teacher: '担当', grade: 'j1', class: 'A', subject: 'math', ...extra,
});
const recorded = { content: '直前の授業', teacher: '前回担当' };

function show({ today, currentSlots = [], currentEntries = {}, previousSlots = [], previousEntries = {}, extra = {} }) {
  elements.datePicker.value = today;
  const ev = event(today, extra);
  context.fixture = {
    schedule: [ev],
    journal: { month: today.slice(0, 7), entries: currentEntries },
    map: { slots: currentSlots },
    previousJournal: { entries: previousEntries },
    previousMap: { slots: previousSlots },
  };
  runInNewContext(`
    scheduleData = fixture.schedule;
    journalData = fixture.journal;
    journalMap = fixture.map;
    prevMonthJournal = fixture.previousJournal;
    prevMonthJournalMap = fixture.previousMap;
    zoomRecordingData = { entries: {} };
    render();
  `, context);
  return elements.output.innerHTML;
}

let result = show({
  today: '2026-09-21',
  currentSlots: { [group]: [key('2026-09-14'), key('2026-09-18')] },
  currentEntries: {
    [key('2026-09-14')]: { content: '前々回の授業' },
    [key('2026-09-18')]: { teacher: '前回担当' },
    [key('2026-09-21')]: { prevEntry: { date: '2026-09-14', content: '前々回の授業' } },
  },
});
assert.match(result, /前回の記録がありません/);
assert.doesNotMatch(result, /前々回の授業/);

result = show({
  today: '2026-09-21',
  currentSlots: { [group]: [key('2026-09-18')] },
  currentEntries: { [key('2026-09-18')]: recorded },
});
assert.match(result, /直前の授業/);
assert.doesNotMatch(result, /前回の記録がありません/);

result = show({
  today: '2026-10-03',
  previousSlots: { [group]: [key('2026-09-20'), key('2026-09-27')] },
  previousEntries: {
    [key('2026-09-20')]: { content: '前々回の授業' },
    [key('2026-09-27')]: {},
  },
});
assert.match(result, /前回の記録がありません/);
assert.doesNotMatch(result, /前々回の授業/);

result = show({
  today: '2026-09-21',
  currentSlots: {
    [group]: [key('2026-09-18')],
    [pairGroup]: [key('2026-09-18', pairGroup)],
  },
  currentEntries: {
    [key('2026-09-18')]: {},
    [key('2026-09-18', pairGroup)]: { content: 'ペア校舎の記録' },
  },
  extra: { _pairGroupKeys: [group, pairGroup] },
});
assert.match(result, /ペア校舎の記録/);

result = show({
  today: '2026-09-21',
  currentSlots: {
    [group]: [key('2026-09-18')],
    [pairGroup]: [key('2026-09-19', pairGroup)],
  },
  currentEntries: {
    [key('2026-09-18')]: { content: '前々回の授業' },
    [key('2026-09-19', pairGroup)]: {},
  },
  extra: { _pairGroupKeys: [group, pairGroup] },
});
assert.match(result, /前回の記録がありません/);
assert.doesNotMatch(result, /前々回の授業/);

console.log('授業準備シートの前回記録テスト: 5件成功');
