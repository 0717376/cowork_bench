"""Evaluation for insales-vip-customer-gsheet-gcal-email.

Checks:
1. Google Sheet "VIP Customer Tracker" with >=10 rows including Gold/Silver/Bronze tiers
2. GCal event "VIP Appreciation Day" ~30 days from launch_time
3. 3 emails sent to Gold tier customers (Scarlett Wright, Ethan Martinez, Olivia Wilson)
   from vip@store.example.com with subject containing VIP and discount/20%
"""
import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []

# Names of checks that must pass; any failure forces overall FAIL regardless of accuracy.
CRITICAL_CHECKS = {
    "GSheet has >=10 rows across all three tiers",
    "GSheet contains top Gold customer (email or RU name)",
    "All 3 Gold tier customers emailed",
    "VIP Appreciation Day is ~30 days from launch",
}

# Top 3 gold customers (by actual WC data: completed orders)
GOLD_CUSTOMERS = [
    {"name": "Скарлетт Рыжов", "email": "scarlett.wright@x.dummyjson.com"},
    {"name": "Артём Мартынов", "email": "ethan.martinez@x.dummyjson.com"},
    {"name": "Оливия Виноградов", "email": "olivia.wilson@x.dummyjson.com"},
]


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    crit = " [CRITICAL]" if name in CRITICAL_CHECKS else ""
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS]{crit} {name}")
    else:
        FAIL_COUNT += 1
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILED.append(name)
        detail_str = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL]{crit} {name}{detail_str}")


def parse_recipients(to_addr):
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


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE LOWER(title) LIKE '%vip%' AND LOWER(title) LIKE '%customer%'
    """)
    sheets = cur.fetchall()
    check("VIP Customer Tracker spreadsheet exists", len(sheets) >= 1,
          f"Found {len(sheets)} matching spreadsheets")

    if sheets:
        ss_id = sheets[0][0]
        cur.execute("""
            SELECT c.value FROM gsheet.cells c
            JOIN gsheet.sheets s ON c.spreadsheet_id = s.spreadsheet_id AND c.sheet_id = s.id
            WHERE c.spreadsheet_id = %s
        """, (ss_id,))
        cells = cur.fetchall()
        all_values = " ".join(str(c[0]) for c in cells if c[0])
        all_lower = all_values.lower()

        g, s, b = all_lower.count("gold"), all_lower.count("silver"), all_lower.count("bronze")
        check("GSheet has >=10 rows across all three tiers",
              (g + s + b) >= 10 and g >= 1 and s >= 1 and b >= 1,
              f"Tiers found: gold={g}, silver={s}, bronze={b}")
        check("GSheet contains 'Gold' tier", "gold" in all_lower, "Gold not found")
        check("GSheet contains 'Silver' tier", "silver" in all_lower, "Silver not found")
        check("GSheet contains 'Bronze' tier", "bronze" in all_lower, "Bronze not found")
        # Top Gold customer: matched by stable English email OR russified name
        # (seed russifies Wright -> Рыжов / first name Скарлетт).
        top_present = (
            "scarlett.wright@x.dummyjson.com" in all_lower
            or "scarlett" in all_lower
            or "скарлетт" in all_lower
            or "рыжов" in all_lower
        )
        check("GSheet contains top Gold customer (email or RU name)",
              top_present, "Top Gold customer not found by email or RU name")
        check("GSheet contains Total_Spent data",
              "3328" in all_values or "3053" in all_values or "2942" in all_values,
              "Spend values not found")

    cur.close()
    conn.close()


def check_gcal(launch_time_str=None):
    print("\n=== Checking Google Calendar ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT summary, start_datetime, end_datetime, description
        FROM gcal.events
        WHERE LOWER(summary) LIKE '%vip%' AND LOWER(summary) LIKE '%appreciation%'
    """)
    events = cur.fetchall()
    check("VIP Appreciation Day event created", len(events) >= 1,
          f"Found {len(events)} matching events")

    if events and launch_time_str:
        try:
            launch_time = datetime.fromisoformat(launch_time_str)
            if launch_time.tzinfo is None:
                launch_time = launch_time.replace(tzinfo=timezone.utc)
            target_date = launch_time + timedelta(days=30)
            for event in events:
                event_start = event[1]
                if event_start.tzinfo is None:
                    event_start = event_start.replace(tzinfo=timezone.utc)
                diff_days = abs((event_start.date() - target_date.date()).days)
                if diff_days <= 2:
                    check("VIP Appreciation Day is ~30 days from launch", True)
                    break
            else:
                check("VIP Appreciation Day is ~30 days from launch", False,
                      f"Closest event at {events[0][1]}, expected ~{target_date.date()}")
        except Exception as e:
            check("VIP Appreciation Day date check", False, str(e))

    cur.close()
    conn.close()


def check_emails():
    print("\n=== Checking VIP Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_emails = cur.fetchall()
    conn.close()

    gold_emails_found = 0
    for customer in GOLD_CUSTOMERS:
        customer_email = customer["email"].lower()
        found = False
        for subj, from_addr, to_addr, body in all_emails:
            recipients = parse_recipients(to_addr)
            if customer_email in recipients:
                found = True
                subj_lower = (subj or "").lower()
                from_lower = (from_addr or "").lower()
                body_lower = (body or "").lower()

                check(f"Email sent to {customer['name']} ({customer['email']})", True)
                check(f"Email from vip@store.example.com for {customer['name']}",
                      "vip@store.example.com" in from_lower,
                      f"From: {from_addr}")
                check(f"Subject contains 'VIP' for {customer['name']}",
                      "vip" in subj_lower, f"Subject: {subj}")
                check(f"Subject contains 'discount' or '20%' for {customer['name']}",
                      "discount" in subj_lower or "20%" in subj_lower,
                      f"Subject: {subj}")
                gold_emails_found += 1
                break
        if not found:
            check(f"Email sent to {customer['name']} ({customer['email']})", False,
                  "No matching email found")

    check("All 3 Gold tier customers emailed", gold_emails_found == 3,
          f"Found {gold_emails_found}/3 emails")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("WC VIP CUSTOMER GSHEET GCAL EMAIL - EVALUATION")
    print("=" * 70)

    check_gsheet()
    check_gcal(args.launch_time)
    check_emails()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}%")

    if CRITICAL_FAILED:
        print(f"  CRITICAL checks failed: {CRITICAL_FAILED}")
        print(f"  Overall: FAIL (critical check failed)")
        sys.exit(1)

    overall = accuracy >= 70.0
    print(f"  Overall: {'PASS' if overall else 'FAIL'} (threshold 70%)")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
