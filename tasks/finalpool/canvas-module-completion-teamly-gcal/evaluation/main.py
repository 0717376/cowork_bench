"""Evaluation for canvas-module-completion-teamly-gcal.

The agent must:
  1. Read course 7 ("Creative Computing and Culture", Fall 2014) modules from
     Canvas and build a Teamly page titled "CCC Fall 2014 - Module Tracker" with
     a table over columns Module_Name / Item_Count / Type_Summary / Status, one
     row per module (Introduction, Week 1, Week 3, Week 5, Week 8), Status set to
     "Not Started".
  2. Create 4 "CCC Module Review Meeting" calendar events on the first four
     consecutive Saturdays after launch_time, each starting 10:00 and lasting 1h.
  3. Email teaching-team@ccc.example.com from ta@university.example.com with
     subject containing "Module Review Setup", referencing the tracker, the 4
     meeting dates and a summary of the 5 modules.

Item_Count is recomputed LIVE from canvas.module_items (nothing hardcoded) and
the expected Saturday dates are computed from the passed --launch_time, so the
eval stays honest and deterministic.

CRITICAL_CHECKS (semantic): any failure => overall FAIL regardless of accuracy.
Otherwise pass threshold: accuracy >= 70%.
"""
import argparse
import datetime as _dt
import json
import os
import re
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

CRITICAL_CHECKS = {
    "Teamly tracker page exists with 5 modules and Status 'Not Started'",
    "Item_Count values match live Canvas module item counts",
    "4 'CCC Module Review Meeting' events on consecutive Saturdays at 10:00 for 1h",
    "Email to teaching-team@ccc.example.com references tracker, 4 dates and 5 modules",
}

# RU+EN aliases for the 5 module names (lowercased; agent writes English names,
# but tolerate Russian renderings of the generic "Week N" / "Introduction").
MODULE_ALIASES = {
    "Introduction": ["introduction", "введение", "вводный"],
    "Week 1": ["week 1", "неделя 1"],
    "Week 3": ["week 3", "неделя 3"],
    "Week 5": ["week 5", "неделя 5"],
    "Week 8": ["week 8", "неделя 8"],
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        print(f"  [FAIL] {name}: {str(detail)[:300]}")


def get_expected_modules():
    """Live per-module item counts for course 7 from Canvas."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT m.name, COUNT(mi.id)
        FROM canvas.modules m
        LEFT JOIN canvas.module_items mi ON mi.module_id = m.id
        WHERE m.course_id = 7
        GROUP BY m.id, m.name
        ORDER BY m.position, m.id
    """)
    rows = cur.fetchall()
    conn.close()
    return {name: cnt for name, cnt in rows}


def expected_saturdays(launch_time):
    """Compute the first 4 consecutive Saturdays strictly after launch_time."""
    # launch_time looks like "2026-05-25 05:33:35 Monday"; take the date part.
    d = None
    if launch_time:
        try:
            d = _dt.date.fromisoformat(str(launch_time).strip()[:10])
        except ValueError:
            d = None
    if d is None:
        d = _dt.datetime.now().date()
    # Saturday == weekday() 5; "first Saturday after launch" => strictly later.
    days_ahead = (5 - d.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7
    first = d + _dt.timedelta(days=days_ahead)
    return [first + _dt.timedelta(weeks=i) for i in range(4)]


def check_teamly(expected):
    print("\n=== Checking Teamly Module Tracker ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title, COALESCE(body, '')
        FROM teamly.pages
        WHERE title ILIKE '%module tracker%'
           OR title ILIKE '%ccc fall 2014%'
           OR (title ILIKE '%module%' AND title ILIKE '%tracker%')
           OR title ILIKE '%трекер%модул%'
    """)
    pages = cur.fetchall()

    if not pages:
        cur.execute("SELECT COUNT(*) FROM teamly.pages")
        total = cur.fetchone()[0]
        record("Teamly tracker page exists with 5 modules and Status 'Not Started'",
               False, f"Found {total} pages, none matching 'CCC Fall 2014 - Module Tracker'")
        record("Item_Count values match live Canvas module item counts", False,
               "No tracker page found")
        conn.close()
        return
    conn.close()

    body = "\n".join(str(b) for _, _, b in pages)
    body_lower = body.lower()

    # Structural (non-critical): column headers present (preserved English names).
    for col in ["Module_Name", "Item_Count", "Type_Summary", "Status"]:
        record(f"Tracker mentions column {col}", col.lower() in body_lower,
               f"'{col}' not found in page body")

    # Module-name coverage (RU+EN tolerant).
    present = {m: any(a in body_lower for a in al) for m, al in MODULE_ALIASES.items()}
    record("Tracker lists all 5 module names",
           all(present.values()),
           f"present={present}")

    # Status 'Not Started' for the 5 entries. The page table should contain
    # 'Not Started' (or RU 'не начат') at least 5 times.
    not_started_en = body_lower.count("not started")
    not_started_ru = len(re.findall(r"не\s+начат", body_lower))
    record("Status 'Not Started' present for all 5 modules",
           (not_started_en + not_started_ru) >= 5,
           f"'Not Started' x{not_started_en}, 'не начат' x{not_started_ru}")

    # CRITICAL: tracker page exists, lists all 5 modules, each marked Not Started.
    record("Teamly tracker page exists with 5 modules and Status 'Not Started'",
           all(present.values()) and (not_started_en + not_started_ru) >= 5,
           f"modules_present={present}, not_started={not_started_en + not_started_ru}")

    # CRITICAL: Item_Count correctness. For each module present in the body, its
    # correct live item count must appear near its name on the same line.
    lines = [ln.lower() for ln in body.splitlines()]
    correct = 0
    misses = []
    for mod, cnt in expected.items():
        aliases = MODULE_ALIASES.get(mod, [mod.lower()])
        cnt_pat = r"(?<!\d)" + re.escape(str(cnt)) + r"(?!\d)"
        ok = False
        for ln in lines:
            if any(a in ln for a in aliases) and re.search(cnt_pat, ln):
                ok = True
                break
        if ok:
            correct += 1
        else:
            misses.append(f"{mod}={cnt}")
    record("Item_Count values match live Canvas module item counts",
           correct == len(expected),
           f"{correct}/{len(expected)} correct; missing/wrong: {misses}")


def check_gcal(saturdays):
    print("\n=== Checking Google Calendar Events ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT summary, start_datetime, end_datetime FROM gcal.events ORDER BY start_datetime")
    events = cur.fetchall()
    conn.close()

    review = [e for e in events
              if "module review" in (e[0] or "").lower()
              or "ccc" in (e[0] or "").lower()
              or "обзор модул" in (e[0] or "").lower()]
    record("At least 4 'CCC Module Review Meeting' events created",
           len(review) >= 4,
           f"Found {len(review)} review events out of {len(events)} total")

    # CRITICAL: each expected Saturday has a matching event starting 10:00 for 1h.
    matched = 0
    details = []
    for sat in saturdays:
        hit = None
        for summary, start, end in review:
            if start is None:
                continue
            sd = start.date()
            if sd != sat:
                continue
            start_ok = (start.hour == 10 and start.minute == 0)
            dur_ok = False
            if end is not None:
                dur = (end - start).total_seconds()
                dur_ok = abs(dur - 3600) < 60
            if start_ok and dur_ok:
                hit = (summary, str(start), str(end))
                break
        if hit:
            matched += 1
        else:
            details.append(str(sat))
    record("4 'CCC Module Review Meeting' events on consecutive Saturdays at 10:00 for 1h",
           matched >= 4,
           f"matched {matched}/4 expected Saturdays {[str(s) for s in saturdays]}; missing: {details}")


def check_email(expected, saturdays):
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    conn.close()

    def recipients_of(to_addr):
        if isinstance(to_addr, list):
            return [str(r).strip().lower() for r in to_addr]
        if isinstance(to_addr, str):
            try:
                parsed = json.loads(to_addr)
                if isinstance(parsed, list):
                    return [str(r).strip().lower() for r in parsed]
            except (json.JSONDecodeError, TypeError):
                pass
            return [to_addr.strip().lower()]
        return []

    target = None
    for subj, from_addr, to_addr, body in emails:
        if "teaching-team@ccc.example.com" in recipients_of(to_addr):
            target = (subj, from_addr, to_addr, body)
            break

    record("Email sent to teaching-team@ccc.example.com", target is not None,
           f"Total emails: {len(emails)}")

    if target is None:
        record("Email to teaching-team@ccc.example.com references tracker, 4 dates and 5 modules",
               False, "No email to teaching-team@ccc.example.com")
        return

    subj, from_addr, to_addr, body = target
    subj_l = (subj or "").lower()
    body = body or ""
    body_l = body.lower()

    record("Email subject contains 'Module Review Setup'",
           "module review setup" in subj_l, f"Subject: {subj}")
    record("Email from ta@university.example.com",
           "ta@university.example.com" in (from_addr or "").lower(),
           f"From: {from_addr}")

    # Tracker reference (name or 'teamly' or RU 'трекер').
    tracker_ref = ("module tracker" in body_l or "ccc fall 2014" in body_l
                   or "teamly" in body_l or "трекер" in body_l)
    # 4 meeting dates: count how many expected Saturdays appear in any common form.
    date_hits = 0
    RU_MONTHS = ["", "января", "февраля", "марта", "апреля", "мая", "июня",
                 "июля", "августа", "сентября", "октября", "ноября", "декабря"]
    for sat in saturdays:
        forms = {
            sat.isoformat(),                                  # 2026-06-06
            f"{sat.day:02d}.{sat.month:02d}.{sat.year}",      # 06.06.2026
            f"{sat.day}.{sat.month}.{sat.year}",              # 6.6.2026
            f"{sat.month:02d}/{sat.day:02d}/{sat.year}",      # 06/06/2026
            sat.strftime("%B %d").lower(),                    # june 06
            f"{sat.strftime('%B').lower()} {sat.day}",        # june 6
            f"{sat.day} {sat.strftime('%B').lower()}",        # 6 june
            f"{sat.day:02d} {sat.strftime('%B').lower()}",    # 06 june
            f"{sat.day} {RU_MONTHS[sat.month]}",              # 6 июня
        }
        if any(f.lower() in body_l for f in forms if f):
            date_hits += 1
    # 5 modules summarised (RU+EN tolerant).
    mods_in_body = sum(1 for al in MODULE_ALIASES.values() if any(a in body_l for a in al))

    record("Email references the tracker", tracker_ref, "No tracker reference found")
    record("Email lists the 4 meeting dates", date_hits >= 4,
           f"{date_hits}/4 dates found; expected {[str(s) for s in saturdays]}")
    record("Email summarises all 5 modules", mods_in_body >= 5,
           f"{mods_in_body}/5 module names found in body")
    record("Email mentions course (Creative Computing / CCC)",
           "creative computing" in body_l or "ccc" in body_l
           or "креативн" in body_l or "вычислен" in body_l, "course name absent")

    # CRITICAL: combined email substance.
    record("Email to teaching-team@ccc.example.com references tracker, 4 dates and 5 modules",
           ("module review setup" in subj_l) and tracker_ref
           and date_hits >= 4 and mods_in_body >= 5,
           f"subject_ok={'module review setup' in subj_l}, tracker={tracker_ref}, "
           f"dates={date_hits}/4, modules={mods_in_body}/5")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("CANVAS MODULE COMPLETION TEAMLY GCAL - EVALUATION")
    print("=" * 70)

    expected = get_expected_modules()
    saturdays = expected_saturdays(args.launch_time)
    print(f"Expected modules (live from Canvas): {expected}")
    print(f"Expected Saturdays (from launch_time={args.launch_time}): {[str(s) for s in saturdays]}")

    check_teamly(expected)
    check_gcal(saturdays)
    check_email(expected, saturdays)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    print(f"\n=== SUMMARY: {PASS_COUNT} passed, {FAIL_COUNT} failed ({accuracy:.1f}%) ===")

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    if critical_failed:
        print(f"  CRITICAL FAILURES ({len(critical_failed)}):")
        for n in critical_failed:
            print(f"    - {n}")

    success = (not critical_failed) and (accuracy >= 70)
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT,
                       "accuracy": accuracy,
                       "critical_failed": critical_failed,
                       "success": success}, f)

    if critical_failed:
        print("  Overall: FAIL (critical check failed)")
        sys.exit(1)
    if accuracy >= 70:
        print("  Overall: PASS")
        sys.exit(0)
    print("  Overall: FAIL")
    sys.exit(1)


if __name__ == "__main__":
    main()
