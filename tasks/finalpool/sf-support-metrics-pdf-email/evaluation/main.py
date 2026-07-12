"""Evaluation for sf-support-metrics-pdf-email (ClickHouse / RU).

A subset of checks are CRITICAL (core semantic deliverables). Any critical
failure forces FAIL (sys.exit(1)) regardless of overall accuracy. Non-critical
checks contribute to an accuracy rate; PASS requires accuracy >= 70% AND no
critical failure.

ISSUE_TYPE data values in sf_data are russified centrally by
db/zzz_clickhouse_after_init.sql, so they are read LIVE from the DB rather than
hardcoded. PRIORITY values (High/Medium/Low) are NOT russified and stay English.
"""
import argparse
import os
import re
import sys

import psycopg2


def normalize_ru_numbers(text):
    """RU number normalization for substring/regex checks: collapse digit-group
    separators (space/NBSP/NNBSP/dot/comma before a 3-digit group) and turn
    decimal commas into dots ("31 588" -> "31588", "4 586,91" -> "4586.91")."""
    t = str(text or "")
    t = re.sub(r"(?<=\d)[ \xa0\u202f\u2009.,](?=\d{3}\b)", "", t)
    return re.sub(r"(?<=\d),(?=\d)", ".", t)

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}

CRITICAL_FAILURES = []
PASSED = 0
TOTAL = 0


def check(name, condition, detail="", critical=False):
    global PASSED, TOTAL
    TOTAL += 1
    tag = "CRITICAL " if critical else ""
    if condition:
        PASSED += 1
        print(f"  [PASS] {tag}{name}")
    else:
        print(f"  [FAIL] {tag}{name}" + (f" -- {detail}" if detail else ""))
        if critical:
            CRITICAL_FAILURES.append(name)


def extract_pdf_text(path):
    """Extract text from PDF using available libraries."""
    try:
        from pypdf import PdfReader
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text
    except ImportError:
        pass
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    except ImportError:
        pass
    with open(path, "rb") as f:
        return f.read().decode("latin-1", errors="ignore")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    # --- Check PDF exists ---
    agent_pdf = os.path.join(args.agent_workspace, "Support_Metrics.pdf")
    if not os.path.exists(agent_pdf):
        print(f"FAIL: Agent output not found: {agent_pdf}")
        sys.exit(1)

    print("  Checking Support_Metrics.pdf...")
    text = normalize_ru_numbers(extract_pdf_text(agent_pdf).lower())

    # --- Title (accept EN literal or RU equivalent) ---
    title_ok = ("support metrics report" in text
                or "отчёт по метрикам поддержки" in text
                or "отчет по метрикам поддержки" in text
                or "метрик" in text and "поддержк" in text)
    check("PDF has title 'Support Metrics Report' (EN or RU)", title_ok)

    # --- Sections (accept EN headers or RU translations) ---
    check("PDF has 'Overview' section (EN/RU)",
          "overview" in text or "обзор" in text)
    check("PDF has 'Issue Type' section (EN/RU)",
          "issue type" in text or "тип обращени" in text or "типу обращени" in text)
    check("PDF has 'Priority' section (EN/RU)",
          "priority" in text or "приоритет" in text)

    # --- Query DB live ---
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT "ISSUE_TYPE", COUNT(*) as cnt,
          ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 2) as avg_resp,
          ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2) as avg_csat
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        GROUP BY "ISSUE_TYPE"
        ORDER BY cnt DESC
    """)
    issue_types = cur.fetchall()

    cur.execute("""
        SELECT "PRIORITY", COUNT(*) as cnt,
          ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 2) as avg_resp,
          ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2) as avg_csat
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        GROUP BY "PRIORITY"
        ORDER BY "PRIORITY"
    """)
    priorities = cur.fetchall()

    cur.execute("""
        SELECT ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2)
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
    """)
    overall_csat = cur.fetchone()[0]
    conn.close()

    total_tickets = sum(r[1] for r in issue_types)

    # --- CRITICAL: total ticket count present (derived from DB) ---
    check(f"PDF contains total ticket count ({total_tickets})",
          str(total_tickets) in text,
          f"expected {total_tickets}", critical=True)

    # --- CRITICAL: every LIVE ISSUE_TYPE value (russified) appears in PDF ---
    missing_issue = [r[0] for r in issue_types if str(r[0]).lower() not in text]
    check("PDF contains all DB ISSUE_TYPE values (live, russified)",
          not missing_issue,
          f"missing: {missing_issue}", critical=True)

    # --- CRITICAL: per-issue-type ticket counts each appear ---
    missing_issue_cnt = [(r[0], r[1]) for r in issue_types if str(r[1]) not in text]
    check("PDF contains per-issue-type ticket counts",
          not missing_issue_cnt,
          f"missing counts: {missing_issue_cnt}", critical=True)

    # --- Priorities present (High/Medium/Low stay English in DB) ---
    for p in priorities:
        check(f"PDF contains priority '{p[0]}'",
              str(p[0]).lower() in text, f"priority {p[0]} not found")

    # --- Priority ticket counts ---
    for r in priorities:
        check(f"PDF contains priority '{r[0]}' ticket count ({r[1]})",
              str(r[1]) in text, f"count {r[1]} not found")

    # --- CRITICAL: overall avg customer satisfaction appears ---
    csat_str = f"{overall_csat:.2f}"
    check(f"PDF contains overall avg satisfaction ({csat_str})",
          csat_str in text or str(overall_csat) in text,
          f"expected {csat_str}", critical=True)

    # --- CRITICAL: email deliverable (recipient + subject + body content) ---
    print("  Checking email deliverable...")
    email_ok = False
    email_detail = "no email found"
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("""
            SELECT subject, to_addr, body_text
            FROM email.messages
            WHERE LOWER(to_addr::text) LIKE '%manager@supportcenter.com%'
        """)
        rows = cur.fetchall()
        conn.close()

        cand = []
        for subj, to_addr, body in rows:
            subj = subj or ""
            body = body or ""
            subj_ok = "monthly support metrics report" in subj.lower()
            body_low = normalize_ru_numbers(body.lower())
            body_ok = (str(total_tickets) in body_low
                       and (f"{overall_csat:.2f}" in body_low or str(overall_csat) in body_low))
            cand.append((subj_ok, body_ok, subj))
            if subj_ok and body_ok:
                email_ok = True
        if not email_ok:
            email_detail = f"recipient/subject/body candidates={[(s, b, subj[:40]) for s, b, subj in cand][:5]}"
    except Exception as e:
        email_detail = f"email DB error: {e}"

    check("Email to manager@supportcenter.com, subject 'Monthly Support Metrics Report', "
          f"body contains total count ({total_tickets}) + avg satisfaction ({csat_str})",
          email_ok, email_detail, critical=True)

    # --- Result gate ---
    pass_rate = PASSED / TOTAL if TOTAL else 0.0
    critical_ok = len(CRITICAL_FAILURES) == 0
    success = critical_ok and pass_rate >= 0.70

    print(f"\n  Passed {PASSED}/{TOTAL} ({pass_rate*100:.1f}%)")
    if not critical_ok:
        print(f"  CRITICAL FAILURES: {CRITICAL_FAILURES}")

    result = {
        "passed": PASSED,
        "total": TOTAL,
        "pass_rate": pass_rate,
        "critical_failures": CRITICAL_FAILURES,
        "success": success,
    }
    if args.res_log_file:
        try:
            import json
            with open(args.res_log_file, "w") as f:
                json.dump(result, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    if not critical_ok:
        print("\n=== RESULT: FAIL (critical) ===")
        sys.exit(1)
    if success:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)
    print("\n=== RESULT: FAIL ===")
    sys.exit(1)


if __name__ == "__main__":
    main()
