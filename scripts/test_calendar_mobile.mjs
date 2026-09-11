import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import http from 'node:http';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {chromium} from 'playwright';

const root = fileURLToPath(new URL('../', import.meta.url));
const key = 'benkyo_calendar_mobile_view_v1';
const pages = ['calendar.html', 'teacher_calendar.html'];
const latest = JSON.parse(await fs.readFile(path.join(root, 'schedule_latest.json'), 'utf8'));
const currentDate = latest.find(e => e.date)?.date;
assert(currentDate);
const server = http.createServer(async (request, response) => {
  try {
    const file = path.resolve(root, '.' + decodeURIComponent(new URL(request.url, 'http://localhost').pathname));
    if (!file.startsWith(root)) throw new Error('Outside test root');
    const types = {'.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json'};
    response.setHeader('Content-Type', (types[path.extname(file)] || 'application/octet-stream') + '; charset=utf-8');
    response.end(await fs.readFile(file));
  } catch { response.writeHead(404); response.end(); }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const base = process.env.CALENDAR_TEST_BASE || `http://127.0.0.1:${server.address().port}/`;
const browser = await chromium.launch({headless: true, ...(process.env.CALENDAR_TEST_CHANNEL ? {channel: process.env.CALENDAR_TEST_CHANNEL} : {})});
try {
  for (const name of pages) for (const width of [320, 390]) {
    const context = await browser.newContext({viewport: {width, height: 844}});
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.clock.install({time: new Date(currentDate + 'T12:00:00+09:00')});
    await page.goto(base + name, {waitUntil: 'networkidle'});
    const viewport = page.locator('.calendar-scroll');
    const compact = page.getByRole('button', {name: 'コンパクト表示', exact: true});
    const scroll = page.getByRole('button', {name: '横スクロール表示', exact: true});
    assert.equal(await compact.getAttribute('aria-pressed'), 'true');
    assert.equal(await viewport.getAttribute('data-view'), 'compact');
    if (name === 'teacher_calendar.html') {
      await page.locator('#teacherChecks input').first().check();
    } else {
      await page.selectOption('#gradeSelect', 'j3');
      await page.selectOption('#campusSelect', 'hon');
      await page.locator('#subjectArea input').first().check();
    }
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), width);
    const day = page.locator('#monthGrid .cell.has-lessons').first();
    await day.click();
    const detail = await page.locator('#dayDetailArea').innerText();
    assert(detail.length > 20);
    if (process.env.CALENDAR_SCREENSHOT_DIR) {
      await fs.mkdir(process.env.CALENDAR_SCREENSHOT_DIR, {recursive: true});
      await page.screenshot({path: path.join(process.env.CALENDAR_SCREENSHOT_DIR, `${name}-${width}-compact.png`), fullPage: true});
    }
    await scroll.click();
    assert.equal(await scroll.getAttribute('aria-pressed'), 'true');
    assert.equal(await page.evaluate(k => localStorage.getItem(k), key), 'scroll');
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), width);
    assert(await viewport.evaluate(e => e.scrollWidth > e.clientWidth));
    assert(await page.locator('#monthGrid .has-lessons .lesson-list').first().isVisible());
    if (process.env.CALENDAR_SCREENSHOT_DIR) await page.screenshot({path: path.join(process.env.CALENDAR_SCREENSHOT_DIR, `${name}-${width}-scroll.png`), fullPage: true});
    await viewport.evaluate(e => { e.scrollLeft = 250; });
    assert(await viewport.evaluate(e => e.scrollLeft > 0));
    assert.equal(await page.locator('#dayDetailArea').innerText(), detail);
    await page.reload({waitUntil: 'networkidle'});
    assert.equal(await scroll.getAttribute('aria-pressed'), 'true');
    await page.click('#prevBtn');
    await page.waitForLoadState('networkidle');
    assert.equal(await viewport.getAttribute('data-view'), 'scroll');
    await page.click('#nextBtn');
    await page.waitForLoadState('networkidle');
    await compact.click();
    assert.equal(await viewport.evaluate(e => e.scrollLeft), 0);
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), width);
    await scroll.click();
    await page.setViewportSize({width: 1280, height: 900});
    assert.equal(await compact.isVisible(), false);
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), 1280);
    assert(await viewport.evaluate(e => e.scrollWidth <= e.clientWidth + 1));
    await page.setViewportSize({width, height: 844});
    assert.equal(await scroll.getAttribute('aria-pressed'), 'true');
    const other = await context.newPage();
    await other.clock.install({time: new Date(currentDate + 'T12:00:00+09:00')});
    await other.goto(base + pages[(pages.indexOf(name) + 1) % pages.length], {waitUntil: 'networkidle'});
    assert.equal(await other.locator('.calendar-scroll').getAttribute('data-view'), 'scroll');
    await other.getByRole('button', {name: 'コンパクト表示', exact: true}).click();
    await page.waitForFunction(() => document.querySelector('.calendar-scroll').dataset.view === 'compact');
    assert.deepEqual(errors, []);
    await context.close();
    console.log(`${name} ${width}px: compact, scroll, persistence, details, month navigation and desktop passed`);
  }
} finally {
  await browser.close();
  await new Promise(resolve => server.close(resolve));
}
