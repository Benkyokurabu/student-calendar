import assert from "node:assert/strict";
import fs from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";

const root = fileURLToPath(new URL("../", import.meta.url));
const rows = JSON.parse(await fs.readFile(path.join(root, "schedule_2026-11.json"), "utf8"));
const teacher = { year: 2026, month: 11, teachers: [...new Set(rows.map((r) => r.teacher).filter(Boolean))], items: rows.map((r) => ({ ...r, campusLabel: r.campus === "hon" ? "本校" : "南教室", teacherLabel: r.teacher, title: r.displayTitle || r.label })) };
const server = http.createServer(async (request, response) => {
  try {
    const file = path.resolve(root, "." + new URL(request.url, "http://local").pathname);
    if (!file.startsWith(root)) throw new Error("outside root");
    response.setHeader("Content-Type", ({ ".html": "text/html", ".js": "text/javascript", ".json": "application/json", ".css": "text/css" })[path.extname(file)] || "text/plain");
    response.end(await fs.readFile(file));
  } catch { response.writeHead(404); response.end(); }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const base = process.env.CALENDAR_TEST_BASE || `http://127.0.0.1:${server.address().port}/`;
const browser = await chromium.launch({ headless: true, ...(process.env.CALENDAR_TEST_CHANNEL ? { channel: process.env.CALENDAR_TEST_CHANNEL } : {}) });
try {
  for (const name of ["calendar.html", "teacher_calendar.html", "lesson_prep.html"]) {
    const page = await browser.newPage({ viewport: { width: 390, height: 844 } });
    const errors = [];
    page.on("pageerror", (error) => errors.push(error.message));
    if (!process.env.CALENDAR_TEST_BASE) {
      await page.clock.install({ time: new Date("2026-10-19T12:00:00+09:00") });
      await page.route("https://line-check-system.vercel.app/api/schedule/calendar*", async (route) => {
        const month = new URL(route.request().url()).searchParams.get("month");
        if (month !== "2026-11") return route.fulfill({ status: 404, json: { error: "not ready" } });
        return route.fulfill({ json: { month, count: rows.length, schedule: rows, teacherSchedule: teacher } });
      });
      await page.route("https://line-check-system.vercel.app/api/recordings/publication*", route => route.fulfill({ json: { rules: [] } }));
      // Prove an absent archive cannot hide a disconnected live calendar.
      await page.route("**/schedule_2026-11.json*", (route) => route.fulfill({ status: 404, body: "" }));
      await page.route("**/teacher_schedule_2026-11.json*", (route) => route.fulfill({ status: 404, body: "" }));
    }
    await page.goto(base + name, { waitUntil: "networkidle" });
    if (name === "lesson_prep.html") {
      await page.locator("#datePicker").fill(rows[0].date);
      await page.locator("#datePicker").dispatchEvent("change");
      await page.waitForFunction(() => scheduleData.some((row) => row.date.startsWith("2026-11-")));
      assert.equal(await page.evaluate(() => scheduleData.filter((row) => row.date.startsWith("2026-11-")).length), rows.length);
    } else {
      await page.getByRole("button", { name: "翌月", exact: false }).click();
      await page.waitForFunction(() => document.getElementById("ym-pill").textContent.includes("11月"));
      await page.waitForFunction(() => document.querySelector("#schedule-source-status")?.dataset.month === "2026-11");
      await page.waitForLoadState("networkidle");
      if (name === "calendar.html") {
        await page.selectOption("#gradeSelect", "j3");
        await page.selectOption("#campusSelect", "hon");
        await page.locator("#subjectArea input").first().check();
      } else {
        for (const input of await page.locator("#teacherChecks input").all()) await input.uncheck();
        const unassigned = rows.filter((row) => !row.teacher).length;
        if (unassigned) {
          await page.getByLabel("未判定", { exact: true }).check();
          assert((await page.locator("#statusArea").innerText()).includes(`${unassigned}件`), "unassigned source events remain selectable");
          await page.getByLabel("未判定", { exact: true }).uncheck();
        }
        await page.locator("#teacherChecks input").first().check();
      }
      await page.locator("#monthGrid .cell.has-lessons").first().waitFor();
      await page.locator("#monthGrid .cell.has-lessons").first().click();
      assert((await page.locator("#dayDetailArea").innerText()).length > 20);
      assert.equal(await page.locator("#schedule-source-status").getAttribute("data-source"), "live");
    }
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), 390);
    assert.deepEqual(errors, []);
    if (process.env.CALENDAR_SCREENSHOT_DIR) {
      await fs.mkdir(process.env.CALENDAR_SCREENSHOT_DIR, { recursive: true });
      await page.screenshot({ path: path.join(process.env.CALENDAR_SCREENSHOT_DIR, name + "-november-390.png"), fullPage: true });
    }
    await page.close();
    console.log(`${name}: November ${rows.length}-event live source, selection and 390px layout passed`);
  }
} finally { await browser.close(); await new Promise((resolve) => server.close(resolve)); }
