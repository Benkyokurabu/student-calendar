/* Optional bridge for the authenticated 勉たん lesson-preparation screen. */
(function (root) {
  "use strict";
  const params = new URLSearchParams(root.location.search);
  function allowedOrigin(value) {
    try {
      const url = new URL(value);
      return url.origin === "https://line-check-system.vercel.app" ||
        (url.protocol === "http:" && ["127.0.0.1", "localhost"].includes(url.hostname));
    } catch { return false; }
  }
  let parentOrigin = "";
  try { parentOrigin = new URL(document.referrer).origin; } catch { /* Standalone page. */ }
  const active = params.get("embedded") === "1" && root.parent !== root && allowedOrigin(parentOrigin);
  const key = value => String(value || "").normalize("NFKC").replace(/\s/gu, "").replace(/(?:先生|さん)$/u, "");
  // Existing login-directory and instructor-document names; never match a surname prefix.
  const aliases = [["髙山", "高山"], ["佐藤(治)", "治彦", "佐藤治彦"], ["佐藤(俊)", "俊介", "佐藤俊介"], ["東", "東泰平"], ["内田", "内田航太"]];
  const equivalents = value => aliases.find(group => group.includes(key(value))) || [key(value)];
  let initialized = false, interacted = false, readySent = false, rendered = null;
  let lastAvailability = "", lastHeight = 0, heightQueued = false;
  const datePicker = () => document.getElementById("datePicker");
  const teacherSelect = () => document.getElementById("teacherSelect");
  function state() {
    const hasLessons = !!rendered?.hasLessons && rendered.date === datePicker().value && rendered.teacher === teacherSelect().value;
    return {hasLessons, ...(hasLessons ? {matchedTeacher: rendered.teacher} : {})};
  }
  function send(type, value) { if (active) root.parent.postMessage({type, ...value}, parentOrigin); }
  function height() { return Math.max(100, Math.ceil(document.body.getBoundingClientRect().height), document.body.scrollHeight); }
  function publishAvailability() {
    if (!active || !readySent) return;
    const value = state(), serialized = JSON.stringify(value);
    if (serialized !== lastAvailability) { lastAvailability = serialized; send("lesson-prep:availability", value); }
  }
  function publishHeight() {
    if (!active || !readySent || heightQueued) return;
    heightQueued = true;
    root.requestAnimationFrame(() => {
      heightQueued = false;
      const value = height();
      if (value !== lastHeight) { lastHeight = value; send("lesson-prep:height", {height: value}); }
    });
  }
  function initialDate(fallback) {
    const value = params.get("date") || "";
    return active && /^20\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(value) &&
      new Date(value + "T00:00:00Z").toISOString().slice(0, 10) === value ? value : fallback;
  }
  function ready(schedule, render) {
    if (!active || initialized) return;
    initialized = true;
    const wanted = params.get("teacher");
    if (!interacted && wanted) {
      const names = new Set(schedule.filter(event => event.date === datePicker().value).map(event => String(event.teacher || "").trim()).filter(Boolean));
      const matches = [...names].filter(name => equivalents(wanted).includes(key(name)));
      if (matches.length === 1 && [...teacherSelect().options].some(option => option.value === matches[0])) {
        teacherSelect().value = matches[0];
        render();
      }
    }
    readySent = true;
    lastHeight = height();
    send("lesson-prep:ready", {...state(), height: lastHeight});
    publishAvailability();
    if (root.ResizeObserver) new ResizeObserver(publishHeight).observe(document.body);
  }
  function onRendered(date, teacher, hasLessons) {
    if (!active) return;
    rendered = {date, teacher, hasLessons};
    publishAvailability();
    publishHeight();
  }
  if (active) {
    document.addEventListener("change", event => {
      if (![datePicker(), teacherSelect()].includes(event.target)) return;
      interacted = true;
      publishAvailability();
    });
    root.addEventListener("message", event => {
      if (event.source !== root.parent || event.origin !== parentOrigin || !allowedOrigin(event.origin)) return;
      if (event.data?.type === "lesson-prep:print" && readySent && state().hasLessons) root.print();
    });
  }
  root.LessonPrepEmbed = {initialDate, ready, onRendered};
})(globalThis);
