/* All active calendars read the same accepted schedule as the classroom screen. */
(function (root) {
  "use strict";
  const endpoint = "https://line-check-system.vercel.app/api/schedule/calendar";
  const cache = new Map();
  const pending = new Map();
  const listeners = new Set();
  let noticeVersion = 0;
  const snapshot = (month, kind = "student") => {
    const data = cache.get(month)?.value;
    return data ? (kind === "teacher" ? data.teacherSchedule : data.schedule) : null;
  };
  const currentMonth = () => new Intl.DateTimeFormat("sv-SE", { timeZone: "Asia/Tokyo", year: "numeric", month: "2-digit" }).format(new Date());
  function notice(month, fallback, silent, loading = false) {
    if (silent || !root.document) return;
    let element = document.getElementById("schedule-source-status");
    if (!element) {
      element = document.createElement("div"); element.id = "schedule-source-status";
      element.setAttribute("role", "status");
      element.style.cssText = "margin:12px;padding:12px;border:1px solid #d49c27;border-radius:8px;background:#fff8dc;color:#5d460e;font:14px/1.6 sans-serif";
      document.body.prepend(element);
    }
    element.hidden = loading || !fallback || month < currentMonth();
    element.textContent = fallback === "unavailable"
      ? "授業データを取得できません。通信状態を確認して、再度読み込んでください。約1分ごとに再確認します。"
      : "最新の授業データに接続できないため、保存済みのスケジュールを表示しています。約1分ごとに再確認します。";
    element.dataset.source = fallback === "unavailable" ? "unavailable" : fallback ? "archive" : "live";
    element.dataset.month = month;
  }
  function validRows(rows, month) {
    return Array.isArray(rows) && rows.length > 0 && rows.every((event) => event && event.date?.startsWith(month + "-") && event.time && event.groupKey && ["hon", "minami"].includes(event.campus));
  }
  async function request(url) {
    const response = await fetch(url, { cache: "no-store", credentials: "omit", signal: AbortSignal.timeout(10000) });
    if (!response.ok) throw new Error(`授業データを取得できません（${response.status}）`);
    return response.json();
  }
  async function live(month) {
    const previous = cache.get(month);
    if (previous && previous.until > Date.now()) return previous.value;
    if (!pending.has(month)) pending.set(month, (async () => {
      const data = await request(`${endpoint}?month=${month}`);
      if (data.month !== month || !validRows(data.schedule, month) || data.count !== data.schedule.length ||
          data.teacherSchedule?.items?.length !== data.count || !validRows(data.teacherSchedule.items, month)) throw new Error("授業データの月・件数が一致しません");
      cache.set(month, { value: data, until: Date.now() + 15000 });
      // A healthy response can arrive after the archive was rendered. Apply it
      // now rather than starting another cold request at the next minute tick.
      for (const listener of listeners) listener(month, data);
      return data;
    })().finally(() => pending.delete(month)));
    return pending.get(month);
  }
  async function load(month, kind = "student", options = {}) {
    if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(month) || !["student", "teacher"].includes(kind)) throw new Error("対象月が不正です");
    const version = options.silent ? null : ++noticeVersion;
    const showNotice = (fallback, loading = false) => {
      if (version === noticeVersion) notice(month, fallback, options.silent, loading);
    };
    showNotice(false, true);
    // Keep checking live data, but do not hold the calendar behind an unhealthy API.
    const liveResult = live(month).then(data => ({ data }), () => ({ data: null }));
    let timer;
    const quick = await Promise.race([
      liveResult,
      new Promise(resolve => { timer = setTimeout(() => resolve(null), 1200); }),
    ]);
    clearTimeout(timer);
    if (quick?.data) {
      showNotice(false);
      return kind === "teacher" ? quick.data.teacherSchedule : quick.data.schedule;
    }
    try {
      const prefix = kind === "teacher" ? "teacher_schedule_" : "schedule_";
      const archived = await request(`${prefix}${month}.json?ts=${Date.now()}`);
      const rows = kind === "teacher" ? archived?.items : archived;
      if (!validRows(rows, month)) throw new Error(`${month}の授業データを確認できません`);
      // A live response may have arrived while reading the archive. It stays authoritative.
      const snapshot = cache.get(month);
      const accepted = snapshot?.until > Date.now() ? snapshot.value : null;
      if (accepted) {
        showNotice(false);
        return kind === "teacher" ? accepted.teacherSchedule : accepted.schedule;
      }
      // The fast archive deadline is not a connection failure. Keep the page
      // usable while live data is pending, and warn only after it actually fails.
      showNotice(true, !quick);
      if (!quick) void liveResult.then(eventual => showNotice(!eventual.data));
      return archived;
    } catch (error) {
      const eventual = await liveResult;
      if (!eventual.data) { showNotice("unavailable"); throw error; }
      showNotice(false);
      return kind === "teacher" ? eventual.data.teacherSchedule : eventual.data.schedule;
    }
  }
  function watch(getMonth, kind, apply, options = {}) {
    let running = false;
    listeners.add((month, data) => {
      if (document.hidden || getMonth() !== month) return;
      notice(month, false, options.silent);
      Promise.resolve().then(() => {
        if (getMonth() === month) return apply(kind === "teacher" ? data.teacherSchedule : data.schedule);
      }).catch(error => console.error("授業データの表示更新に失敗しました", error));
    });
    const refresh = async () => {
      if (running || document.hidden) return;
      running = true;
      const month = getMonth();
      try {
        const value = await load(month, kind, options);
        if (getMonth() === month) apply(value);
      } catch {
        if (options.silent) return;
        notice(month, true, false);
        const element = document.getElementById("schedule-source-status");
        if (element) { element.hidden = false; element.textContent = "授業データの更新を確認できません。現在の表示を保持し、約1分ごとに再確認します。"; }
      } finally { running = false; }
    };
    const timer = setInterval(refresh, 60000);
    document.addEventListener("visibilitychange", () => { if (!document.hidden) void refresh(); });
    return timer;
  }
  root.CalendarScheduleSource = { load, watch, snapshot, currentMonth };
})(globalThis);
