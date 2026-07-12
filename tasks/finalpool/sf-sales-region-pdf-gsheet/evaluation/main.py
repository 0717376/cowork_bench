"""Evaluation for sf-sales-region-pdf-gsheet."""
import argparse
import os
import sys

import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}

# Region values are russified CENTRALLY by the seed (zzz_clickhouse_after_init.sql),
# so the live DB returns Cyrillic REGION labels while a legitimate report may use the
# canonical English label (or vice versa). The region-name critical checks must be
# language-agnostic: a region passes if ANY of its aliases (Cyrillic OR English)
# appears in the extracted PDF text.
REGION_ALIASES = {
    "europe": ["europe", "европа"],
    "asia pacific": ["asia pacific", "азиатско-тихоокеанский регион", "азиатско тихоокеанский регион"],
    "north america": ["north america", "северная америка"],
    "middle east": ["middle east", "ближний восток"],
    "latin america": ["latin america", "латинская америка"],
}
# variant (lowercase) -> canonical english key
_VARIANT_TO_CANON = {v.strip().lower(): canon for canon, variants in REGION_ALIASES.items() for v in variants}


def region_aliases(name):
    """Return the lowercase alias set for a region (English or Russian); falls back to the name itself."""
    canon = _VARIANT_TO_CANON.get(str(name).strip().lower())
    if canon is None:
        return {str(name).strip().lower()}
    return {v.strip().lower() for v in REGION_ALIASES[canon]}


def region_in_text(name, text):
    """True if ANY alias (Cyrillic or English) of the region appears in text."""
    return any(alias in text for alias in region_aliases(name))


def num_close(a, b, tol=1.0):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


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

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    all_errors = []

    # --- Check PDF ---
    agent_pdf = os.path.join(args.agent_workspace, "Regional_Sales_Report.pdf")
    if not os.path.exists(agent_pdf):
        print(f"FAIL: Agent output not found: {agent_pdf}")
        sys.exit(1)

    print("  Checking Regional_Sales_Report.pdf...")
    text = extract_pdf_text(agent_pdf).lower()

    critical_failures = []

    # --- Read ground-truth aggregates from DB (russified REGION values are read live) ---
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT c."REGION",
          COUNT(DISTINCT o."ORDER_ID") as order_count,
          ROUND(SUM(o."TOTAL_AMOUNT")::numeric, 2) as total_revenue,
          COUNT(DISTINCT o."CUSTOMER_ID") as customer_count
        FROM sf_data."SALES_DW__PUBLIC__ORDERS" o
        JOIN sf_data."SALES_DW__PUBLIC__CUSTOMERS" c ON o."CUSTOMER_ID" = c."CUSTOMER_ID"
        WHERE o."STATUS" = 'Доставлен'
        GROUP BY c."REGION"
        ORDER BY c."REGION"
    """)
    db_regions = cur.fetchall()
    conn.close()

    if not db_regions:
        print("FAIL: no delivered orders found in DB (sf_data SALES_DW)")
        sys.exit(1)

    grand_total = sum(float(r[2]) for r in db_regions)
    top_region = max(db_regions, key=lambda x: float(x[2]))

    def revenue_variants(val):
        """Plain and comma-grouped string forms of a 2-decimal revenue value."""
        f = round(float(val), 2)
        plain = f"{f:.2f}"
        grouped = f"{f:,.2f}"
        return {plain.lower(), grouped.lower(), plain.replace(".", ",").lower(), grouped.replace(",", " ").lower()}

    # Check title
    if "regional sales report" not in text:
        all_errors.append("PDF missing title 'Regional Sales Report'")

    # Check summary section
    if "summary" not in text:
        all_errors.append("PDF missing 'Summary' section")

    # --- CRITICAL: regions present (read live from DB; russified Cyrillic values).
    #     Language-agnostic: pass if the Cyrillic OR canonical English alias appears. ---
    for r in db_regions:
        region_name = str(r[0])
        if not region_in_text(region_name, text):
            msg = f"PDF missing region: {region_name}"
            all_errors.append(msg)
            critical_failures.append(msg)

    # --- CRITICAL: per-region delivered-order COUNT appears in PDF ---
    for r in db_regions:
        region_name, order_count = str(r[0]), str(r[1])
        if order_count not in text:
            msg = f"PDF missing order count {order_count} for {region_name}"
            all_errors.append(msg)
            critical_failures.append(msg)

    # --- CRITICAL: per-region total revenue appears in PDF (plain or comma-formatted) ---
    for r in db_regions:
        region_name = str(r[0])
        if not any(v in text for v in revenue_variants(r[2])):
            msg = f"PDF missing revenue {r[2]} for {region_name}"
            all_errors.append(msg)
            critical_failures.append(msg)

    # --- CRITICAL: total delivered orders ---
    total_orders = str(sum(r[1] for r in db_regions))
    if total_orders not in text:
        msg = f"PDF missing total orders: {total_orders}"
        all_errors.append(msg)
        critical_failures.append(msg)

    # --- CRITICAL: top region by revenue mentioned (Cyrillic OR English alias) ---
    if not region_in_text(str(top_region[0]), text):
        msg = f"PDF missing top region: {top_region[0]}"
        all_errors.append(msg)
        critical_failures.append(msg)

    # --- Non-critical: total revenue appears in PDF ---
    if not any(v in text for v in revenue_variants(grand_total)):
        all_errors.append(f"PDF missing total revenue: {grand_total}")

    # --- Google Sheet deliverable (now BLOCKING + CRITICAL) ---
    print("  Checking Google Sheet...")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute(
            "SELECT id, title FROM gsheet.spreadsheets "
            "WHERE LOWER(title) LIKE '%regional%sales%' OR LOWER(title) LIKE '%регионал%'"
        )
        sheets = cur.fetchall()
        if sheets:
            print(f"    Found spreadsheet: {sheets[0][1]}")
            cur.execute(
                "SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s AND LOWER(title) = 'overview'",
                (sheets[0][0],),
            )
            overview = cur.fetchall()
            if overview:
                print("    Found 'Overview' sheet")
            else:
                msg = "Google Sheet 'Overview' sheet not found"
                all_errors.append(msg)
                critical_failures.append(msg)
        else:
            msg = "Google Sheet 'Regional Sales Dashboard' spreadsheet not found"
            all_errors.append(msg)
            critical_failures.append(msg)
        conn.close()
    except Exception as e:
        msg = f"GSheet DB check error: {e}"
        all_errors.append(msg)
        critical_failures.append(msg)

    # --- CRITICAL gate: any critical failure => immediate FAIL ---
    if critical_failures:
        print(f"\n=== RESULT: FAIL (critical) — {len(critical_failures)} critical issue(s) ===")
        for e in critical_failures[:15]:
            print(f"  CRITICAL: {e}")
        sys.exit(1)

    # --- Accuracy gate (>=70) over all checks ---
    total_checks = 4 + 4 * len(db_regions) + 3  # title, summary, total_orders, total_revenue + per-region(4 groups) + top + gsheet
    failed = len(all_errors)
    accuracy = max(0.0, 100.0 * (total_checks - failed) / total_checks) if total_checks else 0.0
    print(f"\nAccuracy: {accuracy:.1f}% ({total_checks - failed}/{total_checks} checks passed)")

    if accuracy >= 70.0:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)
    else:
        print(f"\n=== RESULT: FAIL ({failed} errors) ===")
        for e in all_errors[:15]:
            print(f"  {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
