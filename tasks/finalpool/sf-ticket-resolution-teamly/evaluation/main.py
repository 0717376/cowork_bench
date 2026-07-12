"""
Evaluation for sf-ticket-resolution (ClickHouse + Teamly).

Checks:
1. Teamly "Slow Response Tickets Tracker" space/pages: one page per priority
   (High/Medium/Low) with correct live-computed count and avg SLA utilization.
2. Escalation email to support.manager@company.example.com with subject
   "Slow Response Tickets Report", total, overall percentage and
   per-priority breakdown.

Critical checks (CRITICAL_CHECKS): any failure => overall FAIL regardless of
accuracy. Otherwise PASS requires accuracy >= 70%.

NOTE: sf_data column/identifier names and PRIORITY values (High/Medium/Low)
stay English on purpose — the ClickHouse seed russifies data realia centrally
but leaves PRIORITY/columns English, so these substring checks survive.
"""
import argparse
import json
import os
import re
import sys

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Critical: any failure => overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "Teamly Slow Response Tickets Tracker exists with 3 priority pages",
    "Page High has correct count",
    "Page Medium has correct count",
    "Page Low has correct count",
    "Page High has correct avg utilization",
    "Page Medium has correct avg utilization",
    "Page Low has correct avg utilization",
    "Email addressed to support.manager@company.example.com",
    "Email subject is Slow Response Tickets Report",
    "Email body has correct total slow count",
    "Email body has correct overall slow percentage",
}


def normalize_ru_numbers(text):
    """RU number normalization for substring/regex checks: collapse digit-group
    separators (space/NBSP/NNBSP/dot/comma before a 3-digit group) and turn
    decimal commas into dots ("31 588" -> "31588", "4 586,91" -> "4586.91")."""
    t = str(text or "")
    t = re.sub(r"(?<=\d)[ \xa0\u202f\u2009.,](?=\d{3}\b)", "", t)
    return re.sub(r"(?<=\d),(?=\d)", ".", t)


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_present(value, text, tol=0.5):
    """True if a number close to `value` appears in `text` (RU comma or dot)."""
    if text is None:
        return False
    t = normalize_ru_numbers(text)
    for m in re.findall(r"\d+(?:[.,]\d+)?", t):
        try:
            if abs(float(m.replace(",", ".")) - float(value)) <= tol:
                return True
        except ValueError:
            continue
    return False


def get_expected_data():
    """Query ClickHouse mirror (sf_data) for expected slow response ticket stats."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT "PRIORITY",
               COUNT(*) as count,
               ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 2) as avg_resp,
               ROUND(AVG("SLA_HOURS")::numeric, 2) as avg_sla,
               ROUND(AVG("RESPONSE_TIME_HOURS" / "SLA_HOURS" * 100)::numeric, 2) as avg_util,
               ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2) as avg_csat
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        WHERE "RESPONSE_TIME_HOURS" / "SLA_HOURS" > 0.5
        GROUP BY "PRIORITY"
        ORDER BY "PRIORITY"
    """)
    priority_stats = cur.fetchall()

    cur.execute("""
        SELECT COUNT(*) FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        WHERE "RESPONSE_TIME_HOURS" / "SLA_HOURS" > 0.5
    """)
    total_slow = cur.fetchone()[0]

    cur.execute('SELECT COUNT(*) FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"')
    total_all = cur.fetchone()[0]

    cur.close()
    conn.close()

    return {
        "total_slow": int(total_slow),
        "total_all": int(total_all),
        "slow_pct": round(total_slow / total_all * 100, 2),
        "by_priority": {r[0]: {"count": int(r[1]), "avg_resp": float(r[2]),
                                "avg_sla": float(r[3]), "avg_util": float(r[4]),
                                "avg_csat": float(r[5])} for r in priority_stats},
    }


def check_teamly(expected):
    """Check the Teamly tracker space + per-priority pages (count + avg util)."""
    print("\n=== Checking Teamly ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Agent-created pages only (seed pages have id <= 3).
    cur.execute("SELECT id, title, COALESCE(body, ''), space_id FROM teamly.pages WHERE id > 3")
    pages = cur.fetchall()

    # Tracker container: a space titled with the deliverable marker (RU/EN) OR a
    # parent page carrying it. The English deliverable title is preserved.
    cur.execute("SELECT id, name, COALESCE(description, '') FROM teamly.spaces "
                "WHERE key NOT IN ('TEAM', 'TRIPS')")
    spaces = cur.fetchall()

    def is_tracker_marker(text):
        tl = (text or "").lower()
        return ("slow response tickets tracker" in tl
                or ("slow" in tl and "ticket" in tl)
                or ("медленн" in tl and ("обращен" in tl or "тикет" in tl or "отклик" in tl)))

    tracker_space_ids = {s[0] for s in spaces if is_tracker_marker(s[1]) or is_tracker_marker(s[2])}
    tracker_parent_ids = {p[0] for p in pages if is_tracker_marker(p[1])}

    # Identify the 3 priority pages. Match by exact priority word in the title
    # (PRIORITY stays English). Prefer pages inside the tracker space / under the
    # tracker parent, but accept any agent-created page if container detection is
    # imperfect (broaden, don't over-constrain).
    def find_priority_page(prio):
        prio_l = prio.lower()
        candidates = []
        for pid, title, body, space_id in pages:
            tl = (title or "").lower()
            if re.search(r"\b" + re.escape(prio_l) + r"\b", tl):
                in_tracker = (space_id in tracker_space_ids)
                candidates.append((in_tracker, pid, title, body))
        if not candidates:
            return None
        candidates.sort(key=lambda c: not c[0])  # tracker-scoped first
        return candidates[0]

    prio_pages = {p: find_priority_page(p) for p in ["High", "Medium", "Low"]}
    n_found = sum(1 for v in prio_pages.values() if v is not None)
    has_container = bool(tracker_space_ids or tracker_parent_ids)

    record("Teamly Slow Response Tickets Tracker exists with 3 priority pages",
           has_container and n_found == 3,
           f"container={has_container}, priority pages found={n_found}, "
           f"new pages={[(p[0], p[1]) for p in pages][:10]}")

    # Per-priority count + avg utilization (CRITICAL).
    for prio in ["High", "Medium", "Low"]:
        stats = expected["by_priority"].get(prio)
        page = prio_pages.get(prio)
        text = ((page[2] or "") + " " + (page[3] or "")) if page else ""
        if stats is None:
            record(f"Page {prio} has correct count", False, "no expected stats")
            record(f"Page {prio} has correct avg utilization", False, "no expected stats")
            continue
        record(f"Page {prio} has correct count",
               page is not None and num_present(stats["count"], text, tol=0.4),
               f"expected count {stats['count']}")
        record(f"Page {prio} has correct avg utilization",
               page is not None and num_present(stats["avg_util"], text, tol=0.5),
               f"expected util {stats['avg_util']}")
        # Non-critical extra statistics that were computed-but-unused before.
        record(f"Page {prio} has correct avg response hours",
               page is not None and num_present(stats["avg_resp"], text, tol=0.5),
               f"expected avg_resp {stats['avg_resp']}")
        record(f"Page {prio} has correct avg SLA hours",
               page is not None and num_present(stats["avg_sla"], text, tol=0.5),
               f"expected avg_sla {stats['avg_sla']}")
        record(f"Page {prio} has correct avg CSAT",
               page is not None and num_present(stats["avg_csat"], text, tol=0.2),
               f"expected avg_csat {stats['avg_csat']}")

    cur.close()
    conn.close()


def check_email(expected):
    """Check escalation email (exact recipient + sender, total, percentage, breakdown)."""
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    record("At least 1 email sent", len(emails) >= 1, f"Found {len(emails)}")

    # Pick the report email by subject marker, else fall back to recipient.
    report = None
    for subject, from_addr, to_addr, body_text in emails:
        subj_l = (subject or "").lower()
        to_l = str(to_addr or "").lower()
        if "slow response tickets report" in subj_l or (
            ("ticket" in subj_l or "response" in subj_l or "slow" in subj_l)
            and "support.manager@company.example.com" in to_l
        ):
            report = (subject, from_addr, to_addr, body_text)
            break
    if report is None and emails:
        for e in emails:
            if "support.manager@company.example.com" in str(e[2] or "").lower():
                report = e
                break

    if report is None:
        for n in ["Email addressed to support.manager@company.example.com",
                  "Email body has correct total slow count",
                  "Email body has correct overall slow percentage",
                  "Email subject is Slow Response Tickets Report",
                  "Email body has priority breakdown"]:
            record(n, False, "no report email found")
        cur.close()
        conn.close()
        return

    subject, from_addr, to_addr, body_text = report
    to_l = str(to_addr or "").lower()
    subj_l = (subject or "").lower()
    body = body_text or ""
    body_l = body.lower()

    record("Email addressed to support.manager@company.example.com",
           "support.manager@company.example.com" in to_l, f"To: {to_addr}")
    record("Email subject is Slow Response Tickets Report",
           "slow response tickets report" in subj_l, f"Subject: {subject}")

    record("Email body has correct total slow count",
           num_present(expected["total_slow"], body, tol=0.4),
           f"Expected {expected['total_slow']} in body")
    record("Email body has correct overall slow percentage",
           num_present(expected["slow_pct"], body, tol=0.1),
           f"Expected {expected['slow_pct']}% in body")

    # Priority breakdown: each priority word present (RU prose may surround the
    # English priority codes, so check the English words which stay verbatim).
    breakdown_ok = all(p in body_l for p in ["high", "medium", "low"]) or \
        all(num_present(expected["by_priority"][p]["count"], body, tol=0.4)
            for p in ["High", "Medium", "Low"] if p in expected["by_priority"])
    record("Email body has priority breakdown", breakdown_ok,
           "Missing per-priority breakdown")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected = get_expected_data()
    print(f"[eval] Total slow response: {expected['total_slow']}/{expected['total_all']} "
          f"({expected['slow_pct']}%)")
    print(f"[eval] by_priority: {expected['by_priority']}")

    check_teamly(expected)
    check_email(expected)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    print("\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}%")
    if critical_failed:
        print(f"  CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"    - {n}")

    if args.res_log_file:
        try:
            with open(args.res_log_file, "w") as f:
                json.dump({
                    "total_passed": PASS_COUNT, "total_checks": total,
                    "accuracy": accuracy, "critical_failed": critical_failed,
                }, f, indent=2)
        except Exception:
            pass

    overall = (not critical_failed) and accuracy >= 70
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
