"""Evaluation for canvas-enrollment-teamly.

The agent must:
  1. Build a "Course Enrollment Tracker" page in the Teamly knowledge base whose
     body lists every course with Student_Count / Teacher_Count / TA_Count /
     Total_Enrollment.
  2. Email admin@university.example.com (subject "Course Enrollment Summary
     Report") with the course count, overall total enrollment and the top-5
     courses by Total_Enrollment.

All expected values are recomputed LIVE from canvas.enrollments — nothing is
hardcoded — so the eval stays honest if the seed changes.

CRITICAL_CHECKS (semantic): any failure => overall FAIL regardless of accuracy.
Otherwise pass threshold: accuracy >= 70%.
"""
import argparse
import json
import os
import re
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Critical semantic checks — reflect the task's substance (correct totals,
# correct course count, correct top-5), not mere structure.
CRITICAL_CHECKS = {
    "Tracker page exists in Teamly",
    "Per-course Total_Enrollment correct for >=18 courses",
    "Email total course count correct",
    "Email overall total enrollment correct",
    "Email lists the top-5 courses with their totals",
    "Email sent to admin@university.example.com with correct subject",
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


def num_variants(n):
    """RU + EN renderings of an integer (thousands separators tolerated)."""
    n = int(n)
    s = str(n)
    grouped_space = f"{n:,}".replace(",", " ")   # 32 663
    grouped_nbsp = f"{n:,}".replace(",", " ")
    grouped_comma = f"{n:,}"                       # 32,663
    grouped_dot = f"{n:,}".replace(",", ".")      # 32.663
    return {s, grouped_space, grouped_nbsp, grouped_comma, grouped_dot}


def body_has_number(text, n):
    """True if any RU/EN rendering of n appears as a standalone token in text."""
    for v in num_variants(n):
        # word-ish boundary so e.g. 369 doesn't match inside 2369
        pat = r"(?<!\d)" + re.escape(v) + r"(?!\d)"
        if re.search(pat, text):
            return True
    return False


def short_name(course_name):
    """Course name without the '(Term Year)' suffix, lowercased."""
    return course_name.split("(")[0].strip().lower()


def get_expected():
    """Recompute per-course student/teacher/TA/total counts live from canvas."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""SELECT e.course_id, c.name, e.type, COUNT(*)
        FROM canvas.enrollments e JOIN canvas.courses c ON c.id=e.course_id
        GROUP BY e.course_id, c.name, e.type ORDER BY e.course_id""")
    courses = {}
    for cid, name, etype, cnt in cur.fetchall():
        if cid not in courses:
            courses[cid] = {"name": name, "students": 0, "teachers": 0, "tas": 0}
        if "Student" in etype:
            courses[cid]["students"] = cnt
        elif "Teacher" in etype:
            courses[cid]["teachers"] = cnt
        elif "Ta" in etype:
            courses[cid]["tas"] = cnt
    for c in courses.values():
        c["total"] = c["students"] + c["teachers"] + c["tas"]
    conn.close()
    return courses


def check_teamly(expected):
    print("\n=== Checking Teamly Tracker Page ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Find the tracker page. English title 'Course Enrollment Tracker' is
    # preserved per task.md; accept a few lenient fallbacks.
    cur.execute("""
        SELECT id, title, COALESCE(body, '')
        FROM teamly.pages
        WHERE title ILIKE '%course enrollment tracker%'
           OR title ILIKE '%enrollment tracker%'
           OR (title ILIKE '%enrollment%' AND title ILIKE '%course%')
    """)
    pages = cur.fetchall()

    if not pages:
        cur.execute("SELECT COUNT(*) FROM teamly.pages")
        total = cur.fetchone()[0]
        record("Tracker page exists in Teamly", False,
               f"Found {total} pages, none matching 'Course Enrollment Tracker'")
        conn.close()
        return
    record("Tracker page exists in Teamly", True)

    # Combine the body text of matching pages.
    body = "\n".join(str(b) for _, _, b in pages)
    body_lower = body.lower()

    record("Tracker page has non-trivial body", len(body) >= 200,
           f"Longest/combined body is {len(body)} chars")

    # Column header presence (preserved English property names).
    for col in ["Student_Count", "Teacher_Count", "TA_Count", "Total_Enrollment"]:
        record(f"Tracker mentions column {col}", col.lower() in body_lower,
               f"'{col}' not found in page body")

    # Course-name coverage (structural).
    found_names = sum(1 for c in expected.values() if short_name(c["name"]) in body_lower)
    record("Tracker lists most course names", found_names >= 18,
           f"{found_names}/{len(expected)} course names present")

    # CRITICAL: per-course Total_Enrollment correctness. For each course whose
    # name is present, require its correct total to appear in the body.
    correct_totals = 0
    for c in expected.values():
        if short_name(c["name"]) in body_lower and body_has_number(body, c["total"]):
            correct_totals += 1
    record("Per-course Total_Enrollment correct for >=18 courses",
           correct_totals >= 18,
           f"{correct_totals}/{len(expected)} courses have a correct Total_Enrollment in the body")

    conn.close()


def check_email(expected):
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    record("At least 1 email sent", len(emails) >= 1, f"Found {len(emails)}")

    # Locate the summary email by subject; fall back to the only email.
    summary = None
    for subj, to, body in emails:
        if subj and "enrollment" in subj.lower() and "summary" in subj.lower():
            summary = (subj, to, body)
            break
    if summary is None and emails:
        summary = emails[0]

    if summary is None:
        conn.close()
        return

    subj, to, body = summary
    to_str = json.dumps(to).lower() if isinstance(to, list) else str(to).lower()
    subj_ok = bool(subj) and "course enrollment summary report" in subj.lower()
    to_ok = "admin@university.example.com" in to_str
    record("Email sent to admin@university.example.com with correct subject",
           subj_ok and to_ok, f"subject={subj!r}, to={to}")

    body = body or ""

    # CRITICAL: total course count appears in the body.
    total_courses = len(expected)
    record("Email total course count correct",
           body_has_number(body, total_courses),
           f"Expected course count {total_courses} in body")

    # CRITICAL: overall total enrollment appears (RU/EN number forms).
    overall = sum(c["total"] for c in expected.values())
    record("Email overall total enrollment correct",
           body_has_number(body, overall),
           f"Expected overall total {overall} (RU form '{overall:,}'.replace(',', ' ')) in body")

    # CRITICAL: top-5 courses by total — both name and total must appear.
    top5 = sorted(expected.values(), key=lambda c: c["total"], reverse=True)[:5]
    body_lower = body.lower()
    top5_ok = 0
    details = []
    for c in top5:
        name_ok = short_name(c["name"]) in body_lower
        total_ok = body_has_number(body, c["total"])
        if name_ok and total_ok:
            top5_ok += 1
        else:
            details.append(f"{c['name']}={c['total']} (name={name_ok}, total={total_ok})")
    record("Email lists the top-5 courses with their totals",
           top5_ok >= 5, f"matched {top5_ok}/5; missing: {details}")

    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", default=".")
    parser.add_argument("--groundtruth_workspace", default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected = get_expected()
    check_teamly(expected)
    check_email(expected)

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
