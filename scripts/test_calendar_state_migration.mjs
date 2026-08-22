import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

function extractFunction(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} is missing`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;
  for (let i = bodyStart; i < source.length; i += 1) {
    if (source[i] === "{") depth += 1;
    if (source[i] === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  throw new Error(`${name} has an unterminated body`);
}

for (const page of ["calendar.html", "calendar_journal.html"]) {
  const html = fs.readFileSync(new URL(`../${page}`, import.meta.url), "utf8");
  const migration = extractFunction(html, "migrateLegacySelectionsForLoadedMonth");
  const context = { saveCount: 0 };
  vm.createContext(context);
  vm.runInContext(`
    state = { grade: "j2", campus: "all", selected: { eng: new Set(["B"]) } };
    allLessons = [
      { grade: "j2", subject: "eng", class: "B", campus: "hon" },
      { grade: "j2", subject: "eng", class: "B", campus: "minami" },
    ];
    makeClassToken = (_mode, campus, cls) => \`${"${campus}:${cls}"}\`;
    saveState = () => { saveCount += 1; };
    ${migration};
    migrated = migrateLegacySelectionsForLoadedMonth();
  `, context);

  assert.equal(context.migrated, true, `${page}: legacy selection was not migrated`);
  assert.deepEqual(
    [...context.state.selected.eng].sort(),
    ["hon:B", "minami:B"],
    `${page}: class B must remain visible for both campuses`,
  );
  assert.equal(context.saveCount, 1, `${page}: migrated state must be persisted`);
  assert.equal(
    (html.match(/migrateLegacySelectionsForLoadedMonth\(\);/g) || []).length,
    3,
    `${page}: every schedule-loading path must run the migration`,
  );
}

const august = JSON.parse(fs.readFileSync(new URL("../schedule_2026-08.json", import.meta.url), "utf8"));
const incidentLesson = august.filter(item =>
  item.date === "2026-08-18"
  && item.grade === "j2"
  && item.class === "B"
  && item.subject === "eng"
  && item.campus === "hon"
);
assert.equal(incidentLesson.length, 1, "8/18 本校2B英 must exist exactly once");
assert.equal(incidentLesson[0].groupKey, "hon_j2_B_eng");

console.log("Calendar state migration and 8/18 本校2B英 regression checks passed.");
