"""
Evaluation script for insales-inventory-alert-gcal task.

Conventions:
- 2-5 SEMANTIC CRITICAL checks run with sys.exit(1) BEFORE the accuracy>=70 gate.
- A non-doing agent that emits only the bare IDs (20,39,45,71,79) and the
  literal '5' MUST fail a CRITICAL (it has no Names, no SKUs, no russified
  Categories, no correctly-timed/titled events, no formatted email lines).

Checks:
1. out_of_stock_report.txt: each out-of-stock product's Name + SKU + russified
   Category present (CRITICAL).
2. Google Calendar: per product a 'Restock: [Name]' event on 2026-03-10,
   09:00-10:00 America/New_York (= 13:00-14:00 UTC, EDT) (CRITICAL).
3. Email: sorted-ascending 'ID [id]: [name] (SKU: [sku])' body lines +
   'Total out-of-stock products: 5' (CRITICAL).
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone

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
CRITICAL_FAILS = []

OOS_PRODUCT_IDS = [20, 39, 45, 71, 79]

# Ground-truth out-of-stock products. Category VALUES are russified centrally
# (the whole point of the swap). Names use a distinctive substring (full names
# are long); SKUs are exact. Identifiers (SKU) stay English by convention.
OOS_PRODUCTS = {
    20: {
        "name_substr": "Infinity (JBL Glide 500",
        "sku": "INFINITY-JBL-GLIDE-1020",
        "category": "Электроника",
    },
    39: {
        "name_substr": "BOXTUDIO",
        "sku": "BOXTUDIO-LIGHTBOX-TA-1039",
        "category": "Камеры",
    },
    45: {
        "name_substr": "TYPE-C Earphone for OnePlus 8 Pro",
        "sku": "TYPEC-EARPHONE-FOR-1045",
        "category": "Аудио",
    },
    71: {
        "name_substr": "30kg digital scale for shop and kitchen",
        "sku": "30KG-DIGITAL-SCALE-1071",
        "category": "Бытовая техника",
    },
    79: {
        "name_substr": "Silencer Panels",
        "sku": "SILENCER-PANELS-FLAM-1079",
        "category": "ТВ и домашний кинотеатр",
    },
}

# 2026-03-10 09:00 America/New_York. DST began 2026-03-08, so the date is in
# EDT (UTC-4): 09:00 EDT == 13:00 UTC, 10:00 EDT == 14:00 UTC.
TARGET_START_UTC = datetime(2026, 3, 10, 13, 0, tzinfo=timezone.utc)
TARGET_END_UTC = datetime(2026, 3, 10, 14, 0, tzinfo=timezone.utc)


def record(name, passed, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    tag = "CRIT-" if critical else ""
    if passed:
        PASS_COUNT += 1
        print(f"  [{tag}PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [{tag}FAIL] {name}{msg}")
        if critical:
            CRITICAL_FAILS.append(name)


def connect():
    return psycopg2.connect(**DB_CONFIG)


def to_utc(dt):
    """Normalize a (possibly naive) datetime to aware UTC for comparison."""
    if dt is None:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


# ============================================================================
# Check 1: out_of_stock_report.txt  (Name + SKU + russified Category)
# ============================================================================

def check_text_file(agent_workspace, groundtruth_workspace):
    print("\n=== Checking out_of_stock_report.txt ===")

    agent_file = os.path.join(agent_workspace, "out_of_stock_report.txt")

    if not os.path.isfile(agent_file):
        record("CRITICAL Report file exists", False,
               f"Not found: {agent_file}", critical=True)
        return

    with open(agent_file, encoding="utf-8") as f:
        content = f.read()
    lines = [l.strip() for l in content.splitlines() if l.strip()]
    record("Report file exists", True)
    record("Report line count == 5", len(lines) == 5,
           f"Got {len(lines)} lines")

    # CRITICAL: every product must show ID + Name + SKU + russified Category.
    # A non-doing agent emitting only IDs has none of Name/SKU/Category.
    for pid in OOS_PRODUCT_IDS:
        info = OOS_PRODUCTS[pid]
        # Locate the line for this product by SKU (unique, exact).
        line = next((l for l in lines if info["sku"] in l), None)
        if line is None:
            record(f"CRITICAL Report product {pid}: line w/ SKU {info['sku']}",
                   False, "no line containing this SKU", critical=True)
            continue
        has_id = bool(re.search(rf"(?<!\d){pid}(?!\d)", line))
        has_name = info["name_substr"].lower() in line.lower()
        has_cat = info["category"] in line
        record(
            f"CRITICAL Report product {pid}: ID+Name+SKU+RU-Category",
            has_id and has_name and has_cat,
            f"id={has_id} name={has_name} category({info['category']})={has_cat} "
            f"| line={line[:160]}",
            critical=True,
        )

    # Sorted ascending by leading product ID.
    leading_ids = []
    for l in lines:
        m = re.match(r"\s*(\d+)", l)
        if m:
            leading_ids.append(int(m.group(1)))
    record("Report sorted ascending by ID",
           leading_ids == sorted(leading_ids) and leading_ids == OOS_PRODUCT_IDS,
           f"leading ids = {leading_ids}")


# ============================================================================
# Check 2: Google Calendar  ('Restock: [Name]' @ correct time)
# ============================================================================

def check_gcal():
    print("\n=== Checking Google Calendar ===")

    try:
        conn = connect()
        cur = conn.cursor()
        cur.execute("""
            SELECT summary, description, start_datetime, end_datetime
            FROM gcal.events
            ORDER BY summary
        """)
        events = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("CRITICAL Calendar reachable", False, str(e), critical=True)
        return

    print(f"[check_gcal] Found {len(events)} calendar events.")
    record("At least 5 calendar events created", len(events) >= 5,
           f"Found {len(events)}")

    restock_events = [e for e in events if "restock" in (e[0] or "").lower()]
    record("At least 5 'Restock' events", len(restock_events) >= 5,
           f"Found {len(restock_events)}")

    for pid in OOS_PRODUCT_IDS:
        info = OOS_PRODUCTS[pid]
        name_lc = info["name_substr"].lower()
        sku = info["sku"]

        # Match this product's event by SKU in the description (unique).
        ev = next((e for e in restock_events if sku in (e[1] or "")), None)
        # Fallback: match by distinctive name substring in the summary.
        if ev is None:
            ev = next(
                (e for e in restock_events if name_lc in (e[0] or "").lower()),
                None,
            )

        if ev is None:
            record(f"CRITICAL gcal product {pid}: Restock event present",
                   False, f"no Restock event for SKU {sku}", critical=True)
            continue

        summary, description, sdt, edt = ev
        summary = summary or ""
        description = description or ""

        # Title: 'Restock: [Name]' (Name truncated to <=50 chars). Verify the
        # 'Restock:' prefix AND that the (truncated) product name follows.
        title_prefix_ok = summary.lower().startswith("restock:")
        # First 50 chars of the name, lower-cased, as the truncation target.
        name_head = name_lc[:50].strip()
        # Allow either the head (when name <=50 stays whole) or a strong prefix.
        title_name_ok = name_head[:30] in summary.lower()
        record(
            f"CRITICAL gcal product {pid}: title 'Restock: [Name]'",
            title_prefix_ok and title_name_ok,
            f"summary={summary[:80]!r}",
            critical=True,
        )

        # Time: 09:00-10:00 America/New_York on 2026-03-10 (== 13:00-14:00 UTC).
        s = to_utc(sdt)
        e = to_utc(edt)
        # Accept either UTC-stored (13:00) or naive-local-stored (09:00).
        time_ok = False
        if s is not None:
            for target_s, target_e in (
                (TARGET_START_UTC, TARGET_END_UTC),
                (datetime(2026, 3, 10, 9, 0, tzinfo=timezone.utc),
                 datetime(2026, 3, 10, 10, 0, tzinfo=timezone.utc)),
            ):
                start_ok = abs((s - target_s).total_seconds()) <= 60
                end_ok = e is not None and abs((e - target_e).total_seconds()) <= 60
                if start_ok and end_ok:
                    time_ok = True
                    break
        record(
            f"CRITICAL gcal product {pid}: 09:00-10:00 America/New_York 2026-03-10",
            time_ok,
            f"start={s} end={e}",
            critical=True,
        )

        # Russified Category appears in the description (semantic core of swap).
        record(f"gcal product {pid}: RU Category in description",
               info["category"] in description,
               f"category {info['category']!r} missing")


# ============================================================================
# Check 3: Email  (sorted lines + total count)
# ============================================================================

def check_emails():
    print("\n=== Checking Emails ===")

    try:
        conn = connect()
        cur = conn.cursor()
        cur.execute("""
            SELECT subject, from_addr, to_addr, body_text
            FROM email.messages
        """)
        all_emails = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("CRITICAL Email reachable", False, str(e), critical=True)
        return

    print(f"[check_emails] Found {len(all_emails)} total emails.")
    record("At least 1 email sent", len(all_emails) >= 1, f"Found {len(all_emails)}")

    target = None
    for subject, from_addr, to_addr, body_text in all_emails:
        subj = (subject or "").lower()
        if "out of stock" in subj or "stock alert" in subj:
            target = (subject, from_addr, to_addr, body_text)
            break

    if target is None:
        record("CRITICAL Out-of-stock alert email found", False,
               "no email with 'out of stock'/'stock alert' subject",
               critical=True)
        return
    record("Out-of-stock alert email found", True)

    subject, from_addr, to_addr, body_text = target
    body = body_text or ""

    record("email from purchasing@ourstore.com",
           "purchasing@ourstore.com" in str(from_addr or "").lower(),
           f"From: {from_addr}")
    record("email to warehouse@ourstore.com",
           "warehouse@ourstore.com" in str(to_addr or "").lower(),
           f"To: {to_addr}")

    # CRITICAL: each product line 'ID [id]: [name] (SKU: [sku])'. A non-doing
    # agent that wrote only bare IDs has neither names nor 'SKU:' tokens.
    line_order = []
    for pid in OOS_PRODUCT_IDS:
        info = OOS_PRODUCTS[pid]
        # Expected line shape, matched case-insensitively & whitespace-tolerant.
        pat = re.compile(
            rf"id\s*{pid}\s*:.*sku\s*:\s*{re.escape(info['sku'].lower())}",
            re.IGNORECASE | re.DOTALL,
        )
        # Constrain to a single logical line so other products don't bleed in.
        matched_line = None
        for raw in body.splitlines():
            l = raw.strip()
            if re.match(rf"id\s*{pid}\s*:", l, re.IGNORECASE) and \
               info["sku"].lower() in l.lower() and "sku" in l.lower():
                matched_line = l
                break
        name_ok = matched_line is not None and \
            info["name_substr"].lower() in matched_line.lower()
        record(
            f"CRITICAL email line product {pid}: 'ID {pid}: [name] (SKU: {info['sku']})'",
            matched_line is not None and name_ok,
            f"line={matched_line!r}" if matched_line else "no matching line",
            critical=True,
        )
        if matched_line is not None:
            line_order.append(pid)

    # Lines sorted ascending by ID.
    record("email lines sorted ascending by ID",
           line_order == sorted(line_order) and line_order == OOS_PRODUCT_IDS,
           f"order = {line_order}")

    # CRITICAL: exact total line. Bare '5' anywhere must NOT satisfy this.
    total_ok = bool(re.search(
        r"total\s+out-of-stock\s+products\s*:\s*5\b",
        body, re.IGNORECASE))
    record("CRITICAL email 'Total out-of-stock products: 5'",
           total_ok, "exact total line missing", critical=True)


# ============================================================================
# Main
# ============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_text_file(args.agent_workspace, args.groundtruth_workspace)
    check_gcal()
    check_emails()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0

    print("\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}%")

    # CRITICAL semantic gate runs BEFORE the accuracy gate.
    if CRITICAL_FAILS:
        print(f"  CRITICAL FAILURES: {CRITICAL_FAILS}")
        success = False
    else:
        success = accuracy >= 70.0

    print(f"  Overall: {'PASS' if success else 'FAIL'}")

    if args.res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "accuracy": accuracy,
            "critical_failures": CRITICAL_FAILS,
            "success": success,
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
