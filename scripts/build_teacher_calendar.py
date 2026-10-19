"""Build teacher calendar archives from the canonical student events, never a second parser."""
import argparse
import json
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path


def teacher_payload(rows, month):
    if not rows or any(not event.get("date", "").startswith(month + "-") for event in rows):
        raise ValueError("Schedule is empty or contains another month")
    if len({"|".join(str(event.get(k, "")) for k in ("date", "time", "campus", "groupKey", "room")) for event in rows}) != len(rows):
        raise ValueError("Duplicate schedule entries")
    year, number = map(int, month.split("-"))
    items = [{**event, "campusLabel": {"hon": "本校", "minami": "南教室"}[event["campus"]],
              "teacherLabel": event.get("teacher") or "未判定", "title": event.get("displayTitle") or event.get("label", "")} for event in rows]
    return {"year": year, "month": number, "teachers": sorted({event["teacher"] for event in rows if event.get("teacher")}), "items": items}


def build(root, today=None):
    current = (today or datetime.now(timezone(timedelta(hours=9)))).strftime("%Y-%m")
    payloads = {}
    for path in sorted(root.glob("schedule_*.json")):
        match = re.fullmatch(r"schedule_(20\d{2}-\d{2})\.json", path.name)
        if not match:
            continue
        month = match[1]
        # Historical teacher files are retained. Only active months need rebuilding.
        if month < current:
            continue
        payload = teacher_payload(json.loads(path.read_text("utf-8")), month)
        payloads[month] = payload
        (root / f"teacher_schedule_{month}.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", "utf-8")
    if current in payloads:
        (root / "teacher_schedule_latest.json").write_text(json.dumps(payloads[current], ensure_ascii=False, indent=2) + "\n", "utf-8")
    return payloads


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent.parent)
    args = parser.parse_args()
    for month, payload in build(args.root).items():
        print(f"{month}: teacher calendar {len(payload['items'])} events")
