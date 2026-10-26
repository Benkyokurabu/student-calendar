import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const script = fs.readFileSync(new URL("../calendar-schedule-source.js", import.meta.url), "utf8");
const month = "2026-11";
const row = { date: month + "-02", time: "5:00～6:00", campus: "hon", groupKey: "hon_j1_S_math" };
function setup(responder, overrides = {}) {
  const calls = [];
  const context = { Intl, Date, AbortSignal, Map, setInterval, setTimeout, clearTimeout, fetch: async (url, options) => {
    calls.push({ url, options });
    return responder(url);
  }, ...overrides };
  vm.createContext(context); vm.runInContext(script, context);
  return { source: context.CalendarScheduleSource, calls };
}
const ok = (data) => ({ ok: true, json: async () => data });
const payload = { month, count: 1, schedule: [row], teacherSchedule: { year: 2026, month: 11, teachers: ["工藤"], items: [{ ...row, teacher: "工藤" }] } };
{
  const { source, calls } = setup(() => ok(payload));
  assert.equal((await source.load(month)).length, 1);
  assert.equal((await source.load(month, "teacher")).items.length, 1);
  assert.equal(calls.length, 1, "both calendars share one accepted snapshot");
  assert.equal(calls[0].options.credentials, "omit");
}
{
  const { source, calls } = setup((url) => url.startsWith("https:") ? Promise.reject(new Error("offline")) : ok([row]));
  assert.equal((await source.load(month)).length, 1);
  assert(calls[1].url.startsWith("schedule_2026-11.json"));
  assert(!calls.some((call) => call.url.includes("latest")), "never substitute a different month's latest data");
}
{
  const { source, calls } = setup((url) => url.startsWith("https:") ? ok({ ...payload, count: 2 }) : ok([row]));
  await source.load(month);
  assert.equal(calls.length, 2, "count mismatch must use the checked archive");
}
{
  const { source } = setup(() => ok([{ ...row, date: "2026-10-01" }]));
  await assert.rejects(source.load(month), /確認できません/);
  await assert.rejects(source.load("2026-13"), /不正/);
}
{
  let clock = Date.parse("2026-10-19T09:00:00Z");
  class ClockDate extends Date { constructor(...args) { super(...(args.length ? args : [clock])); } static now() { return clock; } }
  let interval;
  let shown;
  const element = { dataset: {}, style: {}, setAttribute() {} };
  const document = { hidden: false, getElementById: () => element, addEventListener() {} };
  const { source } = setup(() => ok({ ...payload, schedule: [{ ...row, time: String(clock) }] }), {
    Date: ClockDate, document, setInterval(callback, ms) { assert.equal(ms, 60000); interval = callback; return 1; },
  });
  await source.load(month);
  source.watch(() => month, "student", (value) => { shown = value; });
  clock += 60000;
  await interval();
  assert.equal(shown[0].time, String(clock), "an open calendar picks up a new accepted month without another publish");
  assert.equal(element.dataset.source, "live");
}
{
  let finishLive;
  const stalled = new Promise(resolve => { finishLive = resolve; });
  const { source, calls } = setup(url => url.startsWith("https:") ? stalled : ok([row]), {
    setTimeout(callback, ms) { assert.equal(ms, 1200); queueMicrotask(callback); return 1; }, clearTimeout() {},
  });
  assert.equal((await source.load(month))[0].time, row.time, "a stalled API cannot hold the archive for ten seconds");
  assert.equal(calls.length, 2);
  finishLive(ok({ ...payload, schedule: [{ ...row, time: "updated" }] }));
  await new Promise(resolve => setImmediate(resolve));
  assert.equal((await source.load(month))[0].time, "updated", "late accepted data replaces the archive on the next read");
}
{
  let finishLive;
  const stalled = new Promise(resolve => { finishLive = resolve; });
  const { source } = setup(url => url.startsWith("https:") ? stalled : Promise.reject(new Error("no archive")), {
    setTimeout(callback) { queueMicrotask(callback); return 1; }, clearTimeout() {},
  });
  const loaded = source.load(month, "teacher");
  await new Promise(resolve => setImmediate(resolve));
  finishLive(ok(payload));
  assert.equal((await loaded).items.length, 1, "a missing archive must still allow a slower live response");
}
console.log("Shared calendar source: accepted data, cache, bounded fallback, count/month validation and automatic refresh passed.");
{
  let finishLive;
  const stalled = new Promise(resolve => { finishLive = resolve; });
  let current = month;
  const shown = [];
  const element = { dataset: {}, style: {}, setAttribute() {} };
  const document = { hidden: false, getElementById: () => element, addEventListener() {} };
  const { source } = setup(url => url.startsWith("https:") ? stalled : ok([row]), {
    document, setInterval() { return 1; }, setTimeout(callback) { queueMicrotask(callback); return 1; }, clearTimeout() {},
  });
  source.watch(() => current, "student", rows => shown.push(rows));
  const archived = await source.load(month);
  assert.equal(archived[0].time, row.time);
  assert.equal(element.dataset.source, "archive");
  const updated = { ...payload, schedule: [{ ...row, time: "latest" }] };
  finishLive(ok(updated));
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(shown.length, 1, "a delayed live response updates the open page without waiting a minute");
  assert.equal(shown[0][0].time, "latest");
  assert.equal(element.hidden, true);
  assert.equal(element.dataset.source, "live");
  assert.equal(source.snapshot(month)[0].time, "latest", "other slower resources cannot force an older archive back into the page");
  assert.equal(source.snapshot(month, "teacher").items.length, 1);
  assert.equal(source.snapshot("2026-09"), null);
  current = "2026-12";
}
{
  let finishLive;
  const stalled = new Promise(resolve => { finishLive = resolve; });
  let current = month;
  const shown = [];
  const document = { hidden: false, getElementById: () => null, addEventListener() {} };
  const { source } = setup(url => url.startsWith("https:") ? stalled : ok([row]), {
    document, setInterval() { return 1; }, setTimeout(callback) { queueMicrotask(callback); return 1; }, clearTimeout() {},
  });
  source.watch(() => current, "student", rows => shown.push(rows), { silent: true });
  await source.load(month, "student", { silent: true });
  current = "2026-12";
  finishLive(ok(payload));
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(shown.length, 0, "a late response for a previous month never replaces the newly selected month");
}
console.log("Delayed live responses: immediate update, notice recovery, accepted snapshot and month-change isolation passed.");
