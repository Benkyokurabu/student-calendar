/* A response includes its JSON body: stalled bodies also have a deadline. */
(function (root) {
  "use strict";
  async function json(url, timeoutMs = 8000) {
    const controller = new AbortController();
    let timer;
    try {
      return await Promise.race([
        (async () => {
          const response = await fetch(url, { cache: "no-store", signal: controller.signal });
          if (!response.ok) throw new Error(`データを取得できません（${response.status}）`);
          return await response.json();
        })(),
        new Promise((_, reject) => {
          timer = setTimeout(() => { reject(new Error("接続先の応答を時間内に確認できませんでした。")); controller.abort(); }, timeoutMs);
        }),
      ]);
    } finally { clearTimeout(timer); }
  }
  root.CalendarRequest = { json };
})(globalThis);
