"""
Evaluation for the Novgorod trip (rzd + kulinar) task.

Re-themed Russian version of the original Beijing->Qufu task:
  - Route: Москва -> Великий Новгород (rzd "Ласточка" seed).
  - Outbound 2026-03-12: train 820А, departs 16:30 (latest evening), arrives 23:20.
  - Return   2026-03-15: train 821А, departs 15:00 (afternoon),     arrives 22:10.
  - Recipes from kulinar: one каша (категория "гарнир") + one овощное блюдо.

Critical checks (CRITICAL_CHECKS): any failure => overall FAIL regardless of
accuracy (sys.exit(1)). Otherwise pass threshold: accuracy >= 70%.

Deliverable: Novgorod_Trip_Itinerary.docx + 3 gcal events.
"""
import json
import os
import sys
import unicodedata
from argparse import ArgumentParser
from zoneinfo import ZoneInfo

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

# Real kulinar recipe-name fragments (lowercased). These are LIVE names present
# in the kulinar recipe DB. A valid Meal Plan must reference one каша-style dish
# (категория "гарнир") AND one овощное/салат блюдо.
KULINAR_PORRIDGE = ["гречневая каша", "перловая каша"]  # категория "гарнир"
KULINAR_VEGGIE = ["икра кабачковая", "винегрет", "греческий салат",
                  "картофель отварной с укропом", "картофельное пюре",
                  "рис отварной"]

# Critical (semantic) checks: any failure here => overall FAIL.
CRITICAL_CHECKS = {
    "Novgorod_Trip_Itinerary.docx exists",
    "Doc contains outbound train code 820А",
    "Doc contains return train code 821А",
    "Outbound train event exists on 2026-03-12 (820А)",
    "Outbound event starts at 16:30",
    "Return train event exists on 2026-03-15 (821А)",
    "Return event starts at 15:00",
    "Cultural visit event on 2026-03-13 (Новгород/кремль/Софийский)",
    "Meal Plan references two live kulinar dishes",
}

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []


def normalize(s: str) -> str:
    """Lowercase + collapse common cyrillic/latin lookalikes (А/A, etc.)
    so '820А' and '820A' compare equal."""
    s = (s or "").lower()
    s = unicodedata.normalize("NFKD", s)
    table = str.maketrans({"а": "a", "о": "o", "е": "e", "р": "p",
                           "с": "c", "у": "y", "к": "k", "х": "x"})
    return s.translate(table)


def local_hour(dt, tzname):
    """gcal stores start_datetime as UTC; a naive dateTime + timeZone is
    converted to UTC and the IANA zone kept in start_timezone (Google Calendar
    semantics). Convert back to the event's local zone before reading the hour
    so both '16:30 + Europe/Moscow' and a verbatim '16:30 UTC' read as 16."""
    if dt is None:
        return -1
    if tzname and dt.tzinfo is not None:
        return dt.astimezone(ZoneInfo(tzname)).hour
    return dt.hour


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_word(agent_workspace):
    print("\n=== Check 1: Novgorod_Trip_Itinerary.docx ===")
    docx_path = os.path.join(agent_workspace, "Novgorod_Trip_Itinerary.docx")
    if not os.path.exists(docx_path):
        record("Novgorod_Trip_Itinerary.docx exists", False, f"Not found at {docx_path}")
        return
    record("Novgorod_Trip_Itinerary.docx exists", True)

    try:
        import docx
        doc = docx.Document(docx_path)
        raw_text = " ".join(p.text for p in doc.paragraphs)
    except Exception as e:
        record("Word document readable", False, str(e))
        return
    record("Word document readable", True)

    low = raw_text.lower()          # keep originals for RU keyword checks
    norm = normalize(raw_text)      # for cyr/lat train-code matching

    # Section headings: accept EN literal OR RU equivalent.
    record("Contains Trip Overview / Обзор поездки section",
           "trip overview" in low or "обзор поездки" in low or "обзор" in low,
           "No Trip Overview / Обзор поездки found")
    record("Contains Outbound Journey / Маршрут туда section",
           "outbound" in low or "маршрут туда" in low or ("маршрут" in low and "туда" in low),
           "No Outbound Journey / Маршрут туда found")
    record("Contains Return Journey / Обратный путь section",
           ("return" in low and "journey" in low) or "обратный путь" in low or "обратн" in low,
           "No Return Journey / Обратный путь found")
    record("Contains Meal Plan / План питания section",
           "meal plan" in low or "meal" in low or "план питания" in low or "питани" in low,
           "No Meal Plan / План питания found")

    # Train codes (CRITICAL) — matched via normalize for А/A.
    record("Doc contains outbound train code 820А", "820a" in norm,
           "No 820А train code found")
    record("Doc contains return train code 821А", "821a" in norm,
           "No 821А train code found")

    # Destination / landmark keyword.
    record("Mentions Новгород / кремль / Софийский",
           any(kw in low for kw in ["новгород", "кремль", "софийск"]),
           "No Новгород/кремль/Софийский mention found")

    # Meal Plan must reference two distinct LIVE kulinar dishes (CRITICAL).
    has_porridge = any(kw in low for kw in KULINAR_PORRIDGE)
    has_veggie = any(kw in low for kw in KULINAR_VEGGIE)
    record("Meal Plan references two live kulinar dishes",
           has_porridge and has_veggie,
           f"porridge={has_porridge}, veggie={has_veggie}")


def check_gcal():
    print("\n=== Check 2: Google Calendar Events ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Outbound train event on 2026-03-12 (820А Москва -> Великий Новгород).
    cur.execute("""
        SELECT summary, start_datetime, end_datetime, start_timezone
        FROM gcal.events
        WHERE start_datetime >= '2026-03-12' AND start_datetime < '2026-03-13'
        ORDER BY start_datetime
    """)
    events_mar12 = cur.fetchall()
    train_events_12 = [e for e in events_mar12
                       if "820a" in normalize(e[0])
                       or "новгород" in (e[0] or "").lower()
                       or "москва" in (e[0] or "").lower()]
    record("Outbound train event exists on 2026-03-12 (820А)", len(train_events_12) >= 1,
           f"Events on Mar 12: {[e[0] for e in events_mar12]}")

    if train_events_12:
        e = train_events_12[0]
        start_hour = local_hour(e[1], e[3])
        record("Outbound event starts at 16:30", start_hour == 16,
               f"Start time: {e[1]} (tz={e[3]})")

    # Cultural visit on 2026-03-13.
    cur.execute("""
        SELECT summary, start_datetime, end_datetime
        FROM gcal.events
        WHERE start_datetime >= '2026-03-13' AND start_datetime < '2026-03-14'
        ORDER BY start_datetime
    """)
    events_mar13 = cur.fetchall()
    cultural_events = [e for e in events_mar13 if
                       any(kw in (e[0] or "").lower() for kw in
                           ["новгород", "кремль", "софийск", "культурн", "визит"])]
    record("Cultural visit event on 2026-03-13 (Новгород/кремль/Софийский)",
           len(cultural_events) >= 1,
           f"Events on Mar 13: {[e[0] for e in events_mar13]}")

    # Return train event on 2026-03-15 (821А Великий Новгород -> Москва).
    cur.execute("""
        SELECT summary, start_datetime, end_datetime, start_timezone
        FROM gcal.events
        WHERE start_datetime >= '2026-03-15' AND start_datetime < '2026-03-16'
        ORDER BY start_datetime
    """)
    events_mar15 = cur.fetchall()
    return_events = [e for e in events_mar15
                     if "821a" in normalize(e[0])
                     or "новгород" in (e[0] or "").lower()
                     or "москва" in (e[0] or "").lower()]
    record("Return train event exists on 2026-03-15 (821А)", len(return_events) >= 1,
           f"Events on Mar 15: {[e[0] for e in events_mar15]}")

    if return_events:
        e = return_events[0]
        start_hour = local_hour(e[1], e[3])
        record("Return event starts at 15:00", start_hour == 15,
               f"Start time: {e[1]} (tz={e[3]})")

    # Total events created (at least 3).
    cur.execute("SELECT COUNT(*) FROM gcal.events WHERE start_datetime >= '2026-03-12'")
    total = cur.fetchone()[0]
    record("At least 3 calendar events created", total >= 3, f"Found {total} events")

    cur.close()
    conn.close()


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_word(args.agent_workspace)
    check_gcal()

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
    }
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    # Critical gate: any failed critical check => FAIL regardless of accuracy.
    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    if critical_failed:
        print("FAIL: critical check(s) failed: " + "; ".join(critical_failed))
        sys.exit(1)

    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
