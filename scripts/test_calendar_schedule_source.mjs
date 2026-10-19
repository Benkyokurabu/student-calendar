import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const script = fs.readFileSync(new URL("../calendar-schedule-source.js", import.meta.url), "utf8");
const month = "2026-11";
const row = { date: month + "-02", time: "5:00～6:00", campus: "hon", groupKey: "hon_j1_S_math" };
function setup(responder) {
  const calls = [];
  const context = { Intl, Date, AbortSignal, Map, setInterval, fetch: async (url, options) => {
    calls.push({ url, options });
    return responder(url);
  } };
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
console.log("Shared calendar source: accepted data, cache, offline fallback, count and month validation passed.");
