import assert from "node:assert/strict";
import fs from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";

const root = fileURLToPath(new URL("../", import.meta.url));
const month = "2026-11";
const archived = JSON.parse(await fs.readFile(path.join(root, `schedule_${month}.json`), "utf8"));
const rows = archived.map((row, index) => index === 0 ? { ...row, time: "19:01～20:01" } : row);
const teacher = { year: 2026, month: 11, teachers: [], items: rows };
const server = http.createServer(async (request, response) => {
  try {
    const file = path.resolve(root, "." + new URL(request.url, "http://local").pathname);
    if (!file.startsWith(root)) throw new Error("outside root");
    response.setHeader("Content-Type", ({ ".html": "text/html", ".js": "text/javascript", ".json": "application/json", ".css": "text/css" })[path.extname(file)] || "text/plain");
    response.end(await fs.readFile(file));
  } catch { response.writeHead(404); response.end(); }
});
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
const browser = await chromium.launch({ headless: true, ...(process.env.CALENDAR_TEST_CHANNEL ? { channel: process.env.CALENDAR_TEST_CHANNEL } : {}) });
try {
  const page = await browser.newPage({ viewport: { width: 390, height: 844 } });
  const errors = [];
  page.on("pageerror", error => errors.push(error.message));
  await page.clock.install({ time: new Date("2026-11-02T12:00:00+09:00") });
  await page.route("https://raw.githubusercontent.com/**", route => route.fulfill({ status: 404, body: "" }));
  await page.route("https://line-check-system.vercel.app/api/schedule/calendar*", async route => {
    if (new URL(route.request().url()).searchParams.get("month") !== month) return route.fulfill({ status: 404, body: "" });
    await new Promise(resolve => setTimeout(resolve, 2200));
    await route.fulfill({ json: { month, count: rows.length, schedule: rows, teacherSchedule: teacher } });
  });
  // The live schedule arrives after the fast archive, but before this slower
  // companion request. The initial Promise.all must preserve the live result.
  await page.route("https://line-check-system.vercel.app/api/recordings/publication*", async route => {
    await new Promise(resolve => setTimeout(resolve, 4000));
    await route.fulfill({ json: { rules: [] } });
  });
  await page.goto(`http://127.0.0.1:${server.address().port}/lesson_prep.html`, { waitUntil: "domcontentloaded" });
  await page.waitForResponse(response => response.url().includes("schedule_2026-11.json") && response.status() === 200);
  assert.equal(await page.locator("#schedule-source-status").count(), 0, "loading must not create a warning element, even hidden");
  await page.waitForFunction(() => Boolean(CalendarScheduleSource.snapshot("2026-11")));
  await page.waitForFunction(() => document.querySelectorAll("#teacherSelect option").length > 1);
  assert.equal(await page.evaluate(() => scheduleData.find(row => row.date.startsWith("2026-11-")).time), "19:01～20:01");
  assert.equal(await page.locator("#schedule-source-status").isVisible(), false);
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), 390);
  assert.deepEqual(errors, []);
  if (process.env.CALENDAR_SCREENSHOT_DIR) {
    await fs.mkdir(process.env.CALENDAR_SCREENSHOT_DIR, { recursive: true });
    await page.screenshot({ path: path.join(process.env.CALENDAR_SCREENSHOT_DIR, "lesson-prep-late-live-390.png"), fullPage: true });
  }
  console.log("Lesson preparation: delayed live schedule replaces archive, survives slower recordings, clears warning and fits mobile.");
} finally {
  await browser.close();
  await new Promise(resolve => server.close(resolve));
}
