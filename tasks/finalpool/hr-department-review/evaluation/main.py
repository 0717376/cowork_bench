"""
Evaluation script for hr-department-review task.

Checks:
1. Word document Department_Review_2025.docx exists and contains all 7 departments
   with correct headcount, avg salary, avg performance rating (each value bound to
   its specific department section).
2. 7 emails sent to correct department managers with correct content.

ClickHouse HR_ANALYTICS (sf_data schema) data values are russified centrally;
this script queries English column/identifier names and computes expected values
at evaluation time, so it stays in sync with seed/groundtruth.

Scoring: CRITICAL semantic checks must all pass (any critical fail => sys.exit(1)
before the accuracy gate). Otherwise PASS requires accuracy >= 70.
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
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []


def load_expected_depts():
    """Query PostgreSQL to compute expected department data from HR_ANALYTICS tables."""
    conn = psycopg2.connect(**DB_CONFIG)
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT
                e."DEPARTMENT",
                COUNT(e."EMPLOYEE_ID") AS headcount,
                ROUND(AVG(e."SALARY")::numeric, 2) AS avg_salary,
                ROUND(AVG(e."PERFORMANCE_RATING")::numeric, 2) AS avg_perf,
                LOWER(REPLACE(mgr."EMPLOYEE_NAME", ' ', '')) || '@company.com' AS manager_email
            FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES" e
            JOIN sf_data."HR_ANALYTICS__PUBLIC__DEPARTMENTS" d
                ON e."DEPARTMENT" = d."DEPARTMENT_NAME"
            JOIN sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES" mgr
                ON d."MANAGER_ID" = mgr."EMPLOYEE_ID"
            GROUP BY e."DEPARTMENT", mgr."EMPLOYEE_NAME"
            ORDER BY e."DEPARTMENT"
        """)
        rows = cur.fetchall()
        cur.close()

        depts = {}
        for dept, headcount, avg_salary, avg_perf, manager_email in rows:
            depts[dept] = {
                "headcount": int(headcount),
                "avg_salary": float(avg_salary),
                "avg_perf": float(avg_perf),
                "manager_email": manager_email,
            }
        return depts
    finally:
        conn.close()


# Expected department data computed from HR_ANALYTICS at evaluation time
EXPECTED_DEPTS = load_expected_depts()


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    tag = "CRITICAL " if critical else ""
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {tag}{name}")
    else:
        FAIL_COUNT += 1
        if critical:
            CRITICAL_FAILED.append(name)
        detail_str = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL] {tag}{name}{detail_str}")


def normalize_text(text):
    """Normalize text for comparison: lowercase, collapse whitespace."""
    return re.sub(r"\s+", " ", text.lower().strip())


def number_appears(value, text):
    """Check if a number appears in text, with or without comma formatting."""
    val_str = str(value)
    # Try exact match
    if val_str in text:
        return True
    # Try with commas (e.g., 7,096)
    try:
        int_val = int(value)
        formatted = f"{int_val:,}"
        if formatted in text:
            return True
    except (ValueError, TypeError):
        pass
    # Try float with 2 decimals
    try:
        float_val = float(value)
        formatted_2d = f"{float_val:.2f}"
        if formatted_2d in text:
            return True
        # Also try with comma thousands
        int_part = int(float_val)
        dec_part = formatted_2d.split(".")[1]
        formatted_comma = f"{int_part:,}.{dec_part}"
        if formatted_comma in text:
            return True
    except (ValueError, TypeError):
        pass
    return False


def dept_variants(dept):
    """Lowercase variants accepted for a department name."""
    dl = dept.lower()
    if dept == "R&D":
        return ["r&d", "r & d", "r and d", "r&amp;d"]
    return [dl]


def dept_in(dept, text_lower):
    return any(v in text_lower for v in dept_variants(dept))


def find_dept_pos(dept, text_lower):
    """Return earliest index where any variant of dept appears, or -1."""
    positions = [text_lower.find(v) for v in dept_variants(dept)]
    positions = [p for p in positions if p >= 0]
    return min(positions) if positions else -1


def check_word_doc(agent_workspace):
    """Check Department_Review_2025.docx."""
    print("\n=== Checking Word Document ===")
    from docx import Document

    doc_path = os.path.join(agent_workspace, "Department_Review_2025.docx")
    # Structural (non-critical): file existence / readability.
    check("Word file exists", os.path.isfile(doc_path), f"Expected {doc_path}")
    if not os.path.isfile(doc_path):
        # Cannot evaluate the critical content checks without the file.
        for dept in EXPECTED_DEPTS:
            check(f"Department '{dept}' section has correct headcount/salary/perf",
                  False, "Word file missing", critical=True)
        return False

    try:
        doc = Document(doc_path)
    except Exception as e:
        check("Word file readable", False, str(e))
        for dept in EXPECTED_DEPTS:
            check(f"Department '{dept}' section has correct headcount/salary/perf",
                  False, "Word file unreadable", critical=True)
        return False

    check("Word file readable", True)

    # Collect all paragraphs (preserve order for per-section binding).
    paras = [p.text.strip() for p in doc.paragraphs]
    all_text = " ".join(paras)
    all_text_lower = all_text.lower()

    # Structural (non-critical): title heading present. Accept EN or RU wording.
    title_variants = [
        "2025 department performance review",
        "department performance review 2025",
    ]
    has_title = any(v in all_text_lower for v in title_variants)
    # RU acceptance: a heading mentioning 2025 + review/performance of departments.
    if not has_title:
        has_title = ("2025" in all_text_lower and
                     ("отдел" in all_text_lower) and
                     ("эффективност" in all_text_lower or "обзор" in all_text_lower
                      or "аттестац" in all_text_lower or "performance" in all_text_lower))
    check("Title heading present (EN or RU)", has_title)

    # Structural (non-critical): all 7 department names appear somewhere.
    for dept in EXPECTED_DEPTS:
        check(f"Department '{dept}' mentioned in document",
              dept_in(dept, all_text_lower))

    # CRITICAL: bind each dept's headcount AND avg_salary AND avg_perf to that
    # department's section. The section spans from this dept's heading position to
    # the next dept heading (in document order).
    # Build ordered list of (pos, dept) for depts present in the text.
    present = []
    for dept in EXPECTED_DEPTS:
        pos = find_dept_pos(dept, all_text_lower)
        if pos >= 0:
            present.append((pos, dept))
    present.sort()
    ordered_depts = [d for _, d in present]

    def section_text_for(dept):
        """Text belonging to a department's section (heading -> next dept heading)."""
        pos = find_dept_pos(dept, all_text_lower)
        if pos < 0:
            return ""
        # find next department's position after this one
        next_positions = [
            find_dept_pos(o, all_text_lower)
            for o in EXPECTED_DEPTS
            if o != dept and find_dept_pos(o, all_text_lower) > pos
        ]
        end = min(next_positions) if next_positions else len(all_text)
        return all_text[pos:end]

    for dept, data in EXPECTED_DEPTS.items():
        sect = section_text_for(dept)
        hc = data["headcount"]
        avg_sal = data["avg_salary"]
        avg_perf = data["avg_perf"]
        perf_str = f"{avg_perf:.2f}"

        has_hc = number_appears(hc, sect)
        has_sal = number_appears(avg_sal, sect)
        has_perf = (perf_str in sect) or number_appears(avg_perf, sect)

        check(
            f"Department '{dept}' section has correct headcount/salary/perf",
            has_hc and has_sal and has_perf,
            f"hc={hc}:{has_hc} sal={avg_sal}:{has_sal} perf={perf_str}:{has_perf}",
            critical=True,
        )

    return True


def check_emails():
    """Check that 7 emails were sent to the correct managers."""
    print("\n=== Checking Emails ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
        WHERE folder_id = 2 OR folder_id IS NULL
    """)
    all_emails = cur.fetchall()

    # If no emails in folder 2, try all emails
    if len(all_emails) == 0:
        cur.execute("""
            SELECT subject, from_addr, to_addr, body_text
            FROM email.messages
        """)
        all_emails = cur.fetchall()

    def parse_recipients(to_addr):
        """Parse to_addr field into a list of lowercase email strings."""
        if to_addr is None:
            return []
        if isinstance(to_addr, list):
            return [str(r).strip().lower() for r in to_addr]
        to_str = str(to_addr).strip()
        try:
            parsed = json.loads(to_str)
            if isinstance(parsed, list):
                return [str(r).strip().lower() for r in parsed]
            return [to_str.lower()]
        except (json.JSONDecodeError, TypeError):
            return [to_str.lower()]

    def find_email_for(target_email):
        for subj, from_addr, to_addr, body in all_emails:
            recipients = parse_recipients(to_addr)
            if target_email.lower() in recipients:
                return subj, from_addr, to_addr, body
        return None

    # CRITICAL: exactly one email to each of the 7 correct managers (none missing).
    all_recipient_emails = set()
    for subj, from_addr, to_addr, body in all_emails:
        for r in parse_recipients(to_addr):
            all_recipient_emails.add(r)
    expected_emails = {d["manager_email"].lower() for d in EXPECTED_DEPTS.values()}
    missing = expected_emails - all_recipient_emails
    check("All 7 manager emails sent (none missing)", len(missing) == 0,
          f"Missing: {missing}", critical=True)

    for dept, data in EXPECTED_DEPTS.items():
        manager_email = data["manager_email"]
        result = find_email_for(manager_email)
        # presence of the email is part of the missing-managers critical check above;
        # here we evaluate per-email content critically when the email exists.
        if not result:
            check(f"Email to {dept}: subject prefix + dept name", False,
                  f"No email to {manager_email}", critical=True)
            check(f"Email to {dept}: body has correct headcount AND avg salary",
                  False, f"No email to {manager_email}", critical=True)
            continue

        subj, from_addr, to_addr, body = result
        subj_lower = (subj or "").lower()
        body_lower = (body or "").lower()
        body_text = body or ""

        # CRITICAL: subject starts with the required prefix AND contains dept name.
        subj_raw = subj or ""
        prefix_ok = subj_raw.startswith("Department Performance Review - ")
        dept_in_subject = dept_in(dept, subj_lower)
        check(f"Email to {dept}: subject prefix + dept name",
              prefix_ok and dept_in_subject,
              f"Subject: {subj_raw[:100]} prefix={prefix_ok} dept={dept_in_subject}",
              critical=True)

        # Non-critical: body mentions the department name.
        check(f"Email to {dept}: body mentions department name",
              dept_in(dept, body_lower),
              f"Body start: {body_text[:100]}")

        # CRITICAL: body has correct headcount AND avg salary (strengthened from OR).
        hc = data["headcount"]
        avg_sal = data["avg_salary"]
        has_hc = number_appears(hc, body_text)
        has_sal = number_appears(avg_sal, body_text)
        check(f"Email to {dept}: body has correct headcount AND avg salary",
              has_hc and has_sal,
              f"hc={hc}:{has_hc} sal={avg_sal}:{has_sal}",
              critical=True)

    # Non-critical: no large number of unexpected recipients.
    unexpected = all_recipient_emails - expected_emails
    check("No more than 2 unexpected email recipients",
          len(unexpected) <= 2,
          f"Unexpected recipients: {unexpected}")

    cur.close()
    conn.close()
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_word_doc(args.agent_workspace)
    check_emails()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}%")

    if CRITICAL_FAILED:
        print(f"  CRITICAL checks failed: {CRITICAL_FAILED}")
        print(f"  Overall: FAIL (critical)")
        sys.exit(1)

    all_passed = accuracy >= 70
    print(f"  Overall: {'PASS' if all_passed else 'FAIL'}")
    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
