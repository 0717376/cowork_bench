"""
Evaluation script for fetch-arxiv-conference-schedule-gcal-teamly task.

Checks:
1. Teamly knowledge-base page "Conference Reading List" with >= 6 entries,
   real source diversity (ArXiv + Scholar), entry titles matching injected
   source papers, and Relevance_Score values.
2. Google Calendar: exactly the 3 reading-group sessions on the specified
   dates/times, descriptions referencing reading-list paper titles.
3. Conference_Reading_Summary.xlsx in the agent workspace with header + >= 6 rows.

CRITICAL_CHECKS (semantic): any failure => overall FAIL regardless of accuracy.
Pass threshold otherwise: accuracy >= 70%.
"""

import argparse
import json
import os
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

# Critical checks: any failure => overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "Conference Reading List page has >= 6 entries with real source diversity (ArXiv + Scholar)",
    "Entry titles match >= 4 distinct injected source papers",
    "Exactly 3 reading-group events on 2026-03-18/19/20 at 14:00-15:30 with correct topics",
    "Each reading-group description references a reading-list paper title",
    "Relevance scores 1..5 present and Conference_Reading_Summary.xlsx has header + >= 6 rows",
}

# Titles of the injected source papers (arxiv + scholarly). English preserved.
SOURCE_PAPER_TITLES = [
    "Sparse Attention Transformers",
    "Multi-Scale Hierarchical Transformers",
    "Adaptive Learning Rate Methods",
    "Memory-Enhanced Attention Networks",
    "Gradient-Free Optimization Techniques",
    "Self-Attention Distillation",
    "A Survey on Attention Mechanisms",
    "Mixed Precision Training",
    "Vision-Language Pretraining",
]


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


def get_conn():
    return psycopg2.connect(**DB_CONFIG)


def check_teamly():
    """Check the Teamly 'Conference Reading List' page."""
    print("\n=== Checking Teamly Reading List ===")
    try:
        conn = get_conn()
        cur = conn.cursor()
        # User pages only (seed pages id<=3). Title carries English marker.
        cur.execute("SELECT id, title, COALESCE(body, '') FROM teamly.pages WHERE id > 3")
        pages = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Teamly database accessible", False, str(e))
        record("Conference Reading List page has >= 6 entries with real source diversity (ArXiv + Scholar)", False, str(e))
        record("Entry titles match >= 4 distinct injected source papers", False, str(e))
        record("Relevance scores 1..5 present and Conference_Reading_Summary.xlsx has header + >= 6 rows", False, str(e))
        return None

    # Find the reading-list page; skip the noise page ('архив протоколов').
    page = None
    for pid, title, body in pages:
        tl = (title or "").lower()
        if "архив протокол" in tl:
            continue
        if "conference reading list" in tl or ("reading" in tl and "list" in tl) \
                or ("conference" in tl and "reading" in tl) \
                or ("список" in tl and "чтени" in tl):
            page = (pid, title, body)
            break

    if page is None:
        record("Conference Reading List page exists", False,
               f"pages: {[(p[0], p[1]) for p in pages]}")
        record("Conference Reading List page has >= 6 entries with real source diversity (ArXiv + Scholar)", False, "no page")
        record("Entry titles match >= 4 distinct injected source papers", False, "no page")
        return None

    record("Conference Reading List page exists", True)
    body = page[2] or ""
    text = body.lower()

    # Source diversity (CRITICAL): both ArXiv and Scholar appear, and there are
    # >= 6 entries. Count entries by number of distinct Read_By_Date / Title
    # markers — use the number of "relevance" occurrences as entry proxy is
    # fragile, so count source markers and matched titles separately.
    has_arxiv = "arxiv" in text
    has_scholar = "scholar" in text
    # Entry count proxy: occurrences of the 'source' field marker, or title hits.
    title_hits = [t for t in SOURCE_PAPER_TITLES if t.lower() in text]
    n_arxiv = text.count("arxiv")
    n_scholar = text.count("scholar")
    enough_entries = len(title_hits) >= 6 or (n_arxiv + n_scholar) >= 6
    record(
        "Conference Reading List page has >= 6 entries with real source diversity (ArXiv + Scholar)",
        has_arxiv and has_scholar and enough_entries,
        f"arxiv={has_arxiv}({n_arxiv}), scholar={has_scholar}({n_scholar}), title_hits={len(title_hits)}",
    )

    # CRITICAL: >= 4 distinct injected source paper titles present.
    record(
        "Entry titles match >= 4 distinct injected source papers",
        len(title_hits) >= 4,
        f"matched titles ({len(title_hits)}): {title_hits}",
    )

    # Required field identifiers present (non-critical structural).
    field_markers = ["title", "authors", "source", "conference_session",
                     "relevance_score", "read_by_date"]
    present_fields = [f for f in field_markers if f in text]
    record(
        "Page lists required English field identifiers",
        len(present_fields) >= 5,
        f"present: {present_fields}",
    )

    return page


def check_relevance_and_xlsx(page, workspace):
    """CRITICAL: relevance scores 1..5 on the page AND xlsx header + >= 6 rows."""
    print("\n=== Checking Relevance Scores + XLSX ===")

    # --- Relevance scores on the teamly page ---
    import re
    rel_ok = False
    if page is not None:
        body = page[2] or ""
        text = body.lower()
        # Look for digits 1..5 appearing near a 'relevance' marker, or simply
        # presence of the Relevance_Score field plus single-digit 1..5 values.
        has_rel_field = "relevance" in text
        digits = re.findall(r"\brelevance[_ ]?score\b[^\n|]*?([1-5])", text)
        if len(digits) < 6:
            # fallback: any standalone 1..5 tokens in the body (markdown table)
            digits = max(digits, re.findall(r"\b([1-5])\b", text), key=len)
        rel_ok = has_rel_field and len(digits) >= 6
    record("Relevance_Score field with 1..5 values present on page", rel_ok)

    # --- XLSX ---
    xlsx_ok = False
    detail = ""
    try:
        import openpyxl
    except ImportError:
        record("openpyxl available", False, "Cannot import openpyxl")
        openpyxl = None

    rows_ok = False
    header_ok = False
    if openpyxl is not None:
        xlsx_path = os.path.join(workspace, "Conference_Reading_Summary.xlsx")
        if not os.path.isfile(xlsx_path):
            record("Conference_Reading_Summary.xlsx exists", False, f"Not found: {xlsx_path}")
        else:
            record("Conference_Reading_Summary.xlsx exists", True)
            try:
                wb = openpyxl.load_workbook(xlsx_path, data_only=True)
                best_rows = 0
                hdr_found = False
                for ws in wb.worksheets:
                    rows = [r for r in ws.iter_rows(values_only=True)
                            if any(c is not None for c in r)]
                    if not rows:
                        continue
                    header = " ".join(str(c).lower() for c in rows[0] if c is not None)
                    if "title" in header and "source" in header:
                        hdr_found = True
                    data_rows = len(rows) - 1
                    best_rows = max(best_rows, data_rows)
                wb.close()
                header_ok = hdr_found
                rows_ok = best_rows >= 6
                detail = f"header_ok={hdr_found}, data_rows={best_rows}"
            except Exception as e:
                detail = str(e)
                record("XLSX readable", False, str(e))

    record("Conference_Reading_Summary.xlsx has expected header", header_ok, detail)
    record("Conference_Reading_Summary.xlsx has >= 6 data rows", rows_ok, detail)

    xlsx_ok = header_ok and rows_ok
    # Aggregate CRITICAL check.
    record(
        "Relevance scores 1..5 present and Conference_Reading_Summary.xlsx has header + >= 6 rows",
        rel_ok and xlsx_ok,
        f"rel_ok={rel_ok}, xlsx_ok={xlsx_ok}",
    )


def check_calendar(page):
    """Check Google Calendar reading-group sessions."""
    print("\n=== Checking Google Calendar ===")
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "SELECT summary, description, start_datetime, end_datetime FROM gcal.events"
        )
        events = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Calendar DB accessible", False, str(e))
        record("Exactly 3 reading-group events on 2026-03-18/19/20 at 14:00-15:30 with correct topics", False, str(e))
        record("Each reading-group description references a reading-list paper title", False, str(e))
        return

    reading_events = [
        e for e in events if "reading group" in (e[0] or "").lower()
    ]
    record("3 reading group events exist", len(reading_events) >= 3,
           f"Found {len(reading_events)} reading group events")

    # Build expected (date -> topic keyword) mapping.
    expected = {
        "2026-03-18": "transformer",
        "2026-03-19": "attention",
        "2026-03-20": "optim",
    }

    matched = {}
    for summary, description, start_dt, end_dt in reading_events:
        s_lower = (summary or "").lower()
        sd = str(start_dt) if start_dt else ""
        ed = str(end_dt) if end_dt else ""
        for date_key, topic in expected.items():
            if date_key in sd and topic in s_lower:
                # Verify 14:00 start and 15:30 end.
                start_ok = "14:00" in sd
                end_ok = "15:30" in ed
                if start_ok and end_ok:
                    matched[date_key] = (summary, description)

    record(
        "Exactly 3 reading-group events on 2026-03-18/19/20 at 14:00-15:30 with correct topics",
        len(matched) == 3,
        f"matched dates: {sorted(matched.keys())}",
    )

    # Non-critical per-topic info.
    for date_key, topic in expected.items():
        record(f"Reading group '{topic}' on {date_key} (14:00-15:30) present",
               date_key in matched)

    # CRITICAL: each matched description references a reading-list paper title.
    page_text = (page[2].lower() if page else "")
    desc_ref_ok = 0
    for date_key, (summary, description) in matched.items():
        desc = (description or "").lower()
        # A paper title from the source list that appears both in the page and
        # in this event description.
        for t in SOURCE_PAPER_TITLES:
            tl = t.lower()
            if tl in desc and (not page_text or tl in page_text):
                desc_ref_ok += 1
                break
    record(
        "Each reading-group description references a reading-list paper title",
        desc_ref_ok >= 3 and len(matched) == 3,
        f"descriptions referencing a reading-list title: {desc_ref_ok}/3",
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    page = check_teamly()
    check_relevance_and_xlsx(page, args.agent_workspace)
    check_calendar(page)

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT} ({accuracy:.1f}%)")
    if critical_failed:
        print(f"  CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"    - {n}")

    if args.res_log_file:
        try:
            with open(args.res_log_file, "w") as f:
                json.dump({
                    "total_passed": PASS_COUNT,
                    "total_checks": total,
                    "accuracy": accuracy,
                    "critical_failed": critical_failed,
                }, f, indent=2)
        except Exception:
            pass

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
