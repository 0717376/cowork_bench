"""
Evaluation for insales-inventory-report-pdf-gcal task.

Checks:
1. PDF file Inventory_Audit.pdf exists and contains expected data
2. Google Calendar event created (non-blocking)
"""

import argparse
import json
import os
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

# Semantic critical failures => immediate FAIL regardless of accuracy
CRITICAL_FAILURES = []

# Known out-of-stock product IDs
OOS_IDS = [20, 39, 45, 71, 79]
# Known low-stock product IDs (qty 1-5)
LOW_IDS = [5, 7, 15, 19, 21, 22, 27, 28, 30, 31, 32, 50, 52, 53, 54, 61, 63, 70, 77, 82]

# RU+EN substring alternatives (the agent writes the PDF in Russian after the InSales swap)
OOS_TERMS = ["out of stock", "нет в наличии", "отсутству", "не в наличии"]
LOW_TERMS = ["low stock", "низкий остаток", "низкий запас", "малый остаток"]
TITLE_TERMS_EN = ["inventory", "audit"]
TITLE_TERMS_RU = ["аудит", "запас"]
TOTAL_TERMS = ["total products", "всего товаров", "всего товара", "общее количество товаров"]


def record(name, passed, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        tag = " [CRITICAL]" if critical else ""
        print(f"  [FAIL]{tag} {name}{msg}")
        if critical:
            CRITICAL_FAILURES.append(name)


def any_in(terms, text_lower):
    return any(t in text_lower for t in terms)


def check_pdf(agent_workspace):
    """Check the PDF file exists and has key content."""
    print("\n=== Checking Inventory_Audit.pdf ===")

    pdf_path = os.path.join(agent_workspace, "Inventory_Audit.pdf")
    if not os.path.isfile(pdf_path):
        record("PDF file exists", False, f"Not found: {pdf_path}", critical=True)
        return False
    record("PDF file exists", True)

    # Check file size is reasonable (at least 1KB)
    size = os.path.getsize(pdf_path)
    record("PDF file size reasonable (>1KB)", size > 1024, f"Size: {size} bytes", critical=True)

    # Try to extract text
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            text = ""
            for page in reader.pages:
                text += page.extract_text() or ""
    except ImportError:
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as p:
                text = ""
                for page in p.pages:
                    text += page.extract_text() or ""
        except ImportError:
            print("  [WARN] No PDF reader available (PyPDF2 or pdfplumber). Checking file existence only.")
            return True

    text_lower = text.lower()

    # Check title (RU "Аудит запасов" OR EN "Inventory Audit")
    title_ok = (all(t in text_lower for t in TITLE_TERMS_EN)
                or all(t in text_lower for t in TITLE_TERMS_RU))
    record("PDF contains title (RU/EN)", title_ok,
           "Expected 'Inventory Audit' or 'Аудит запасов' in PDF")

    # Check summary section: total products count (language-neutral digits)
    record("PDF contains total products count (82)",
           "82" in text or any_in(TOTAL_TERMS, text_lower),
           "Expected total products count 82 / 'всего товаров'", critical=True)

    record("PDF mentions out of stock (RU/EN)",
           any_in(OOS_TERMS, text_lower),
           "Expected 'out of stock' / 'нет в наличии' text")

    record("PDF mentions low stock (RU/EN)",
           any_in(LOW_TERMS, text_lower),
           "Expected 'low stock' / 'низкий остаток' text")

    # Check all 5 OOS product IDs appear (language-neutral) - CRITICAL
    found_oos = 0
    for pid in OOS_IDS:
        if str(pid) in text:
            found_oos += 1
    record("PDF lists all out-of-stock product IDs",
           found_oos >= 5, f"Found {found_oos}/5 OOS product IDs", critical=True)

    # Check low-stock product IDs (full list of 20) - CRITICAL >= 15
    found_low = 0
    for pid in LOW_IDS:
        if str(pid) in text:
            found_low += 1
    record("PDF lists low-stock product IDs",
           found_low >= 15, f"Found {found_low}/20 low-stock product IDs", critical=True)

    return True


def check_gcal():
    """Check Google Calendar restock event on 2026-03-12 - RECORDED & CRITICAL."""
    print("\n=== Checking Google Calendar restock event ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("""
            SELECT summary, description, start_datetime
              FROM gcal.events
             WHERE start_datetime >= '2026-03-12'
               AND start_datetime < '2026-03-13'
             ORDER BY start_datetime
        """)
        events = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Calendar restock event exists on 2026-03-12", False,
               f"DB error: {e}", critical=True)
        return

    # Restock event by RU or EN title keyword
    restock_events = []
    for ev in events:
        s = (ev[0] or "").lower()
        if ("restock" in s or "пополнен" in s or
                ("планиров" in s and "запас" in s)):
            restock_events.append(ev)

    record("Calendar restock event exists on 2026-03-12",
           len(restock_events) > 0,
           f"{len(events)} events on 2026-03-12, none restock-related"
           if not restock_events else "", critical=True)

    if not restock_events:
        return

    ev = restock_events[0]
    desc = (ev[1] or "")
    print(f"    Title: {ev[0]}, Start: {ev[2]}")

    # Description must mention the OOS count (5) and the low-stock count (20)
    desc_ok = ("5" in desc and "20" in desc)
    record("Restock event description mentions stock counts (5 OOS / 20 low)",
           desc_ok, f"Desc: {desc[:150]}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_pdf(args.agent_workspace)
    check_gcal()  # now recorded & critical

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (100.0 * PASS_COUNT / total) if total else 0.0

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}%")

    # CRITICAL gate: any semantic critical failure => immediate FAIL
    if CRITICAL_FAILURES:
        print(f"  CRITICAL FAILURES: {CRITICAL_FAILURES}")
        print(f"  Overall: FAIL (critical check failed)")
        if args.res_log_file:
            with open(args.res_log_file, "w") as f:
                json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT,
                           "accuracy": accuracy, "success": False,
                           "critical_failures": CRITICAL_FAILURES}, f, indent=2)
        sys.exit(1)

    overall = accuracy >= 70.0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    if args.res_log_file:
        result = {"passed": PASS_COUNT, "failed": FAIL_COUNT,
                  "accuracy": accuracy, "success": overall}
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
