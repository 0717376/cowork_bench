"""Evaluation for sf-hr-turnover-notion (ClickHouse + Teamly fork).

Deliverable:
  1. A Teamly knowledge-base page titled (English literal) "HR Department
     Workforce Analysis" whose body holds a "Department Metrics" table with one
     row per department: Department, Employee_Count, Avg_Salary, Min_Salary,
     Max_Salary, Salary_Range. Below it, a paragraph naming the highest-avg and
     lowest-avg departments with their values.
  2. An email to hr-director@company.com (subject English literal
     "HR Department Workforce Analysis Ready") whose body states the total
     employee count and the highest-avg department.

sf_data (ClickHouse HR_ANALYTICS) data VALUES are russified centrally by
db/zzz_clickhouse_after_init.sql, so DEPARTMENT_NAME comes back as Russian
(Инженерия/Финансы/Кадры/Операции/НИОКР/Продажи/Поддержка). This script computes
all expected metrics from sf_data at evaluation time -> always in sync with
seed/groundtruth. Numeric column/identifier names stay English.

Scoring: CRITICAL semantic checks must all pass (any critical fail => sys.exit(1)
before the accuracy gate). Otherwise PASS requires accuracy >= 70.
"""
import argparse
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
CRITICAL_FAILED = []


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
        d = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL] {tag}{name}{d}")


def load_expected_depts(cur):
    """Compute per-department metrics from sf_data. Keys = russified DEPARTMENT_NAME."""
    # EMPLOYEES carries its own russified DEPARTMENT text column (== DEPARTMENTS.
    # DEPARTMENT_NAME), so group on it directly (matches sibling HR tasks).
    cur.execute(
        """
        SELECT e."DEPARTMENT",
               COUNT(e."EMPLOYEE_ID")              AS cnt,
               ROUND(AVG(e."SALARY")::numeric, 2)  AS avg_salary,
               MIN(e."SALARY")                     AS min_salary,
               MAX(e."SALARY")                     AS max_salary,
               MAX(e."SALARY") - MIN(e."SALARY")   AS salary_range
        FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES" e
        GROUP BY e."DEPARTMENT"
        ORDER BY avg_salary DESC
        """
    )
    rows = cur.fetchall()
    depts = {}
    for name, cnt, avg, mn, mx, rng in rows:
        depts[name] = {
            "count": int(cnt),
            "avg": float(avg),
            "min": float(mn),
            "max": float(mx),
            "range": float(rng),
        }
    return depts


def num_variants(value, decimals=None):
    """String forms a number may legitimately take in free text / table cells."""
    out = set()
    try:
        f = float(value)
    except (TypeError, ValueError):
        return out
    is_int = abs(f - round(f)) < 1e-9
    if is_int:
        iv = int(round(f))
        out.add(str(iv))
        out.add(f"{iv:,}")          # 7,096
        out.add(f"{iv:.2f}")        # 7096.00
        out.add(f"{iv:,.2f}")       # 7,096.00
    else:
        out.add(f"{f:.2f}")         # 58991.61
        out.add(f"{f:,.2f}")        # 58,991.61
        out.add(str(f))
    return {v for v in out if v}


def num_in(value, text):
    return any(v in text for v in num_variants(value))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    expected = load_expected_depts(cur)
    print("Expected departments (avg desc):")
    for name, m in expected.items():
        print(f"  {name}: {m}")

    total_employees = sum(m["count"] for m in expected.values())
    ranked = sorted(expected.items(), key=lambda kv: kv[1]["avg"], reverse=True)
    high_name, high_m = ranked[0]
    low_name, low_m = ranked[-1]

    # ---- Locate the Teamly page (English title literal; broadened grep) ----
    print("\n=== Checking Teamly page ===")
    cur.execute(
        """
        SELECT title, COALESCE(body, '')
        FROM teamly.pages
        WHERE title ILIKE '%%hr department workforce analysis%%'
           OR (title ILIKE '%%workforce%%' AND title ILIKE '%%analysis%%')
           OR (title ILIKE '%%department%%' AND title ILIKE '%%metric%%')
        """
    )
    pages = cur.fetchall()
    if not pages:
        cur.execute("SELECT COUNT(*) FROM teamly.pages")
        total = cur.fetchone()[0]
        check("Teamly page 'HR Department Workforce Analysis' exists", False,
              f"{total} pages exist but none match the expected title", critical=True)
        body = ""
    else:
        check("Teamly page 'HR Department Workforce Analysis' exists", True)
        # Pick the longest matching body (the real deliverable).
        title, body = max(pages, key=lambda tb: len(tb[1]))
    body_l = body.lower()

    check("Teamly page body is non-trivial", len(body) >= 150,
          f"longest matching body is {len(body)} chars")

    # 'Department Metrics' table label present (English literal kept per task).
    check("Body references 'Department Metrics' table",
          "department metrics" in body_l)

    # ---- Per-department metric values present in the body ----
    print("\n=== Checking department metrics in page body ===")
    for name, m in expected.items():
        nl = name.lower()
        present = nl in body_l
        check(f"Department '{name}' present in body", present)

        # CRITICAL: every metric value for this dept must appear in the body.
        # Values are unique enough (6-7k counts, 5-6 digit salaries) that this
        # verifies the agent computed real aggregates, not garbage.
        check(f"{name}: Employee_Count {m['count']}",
              num_in(m["count"], body), critical=True)
        check(f"{name}: Avg_Salary {m['avg']:.2f}",
              num_in(m["avg"], body), critical=True)
        check(f"{name}: Min_Salary {int(m['min'])}",
              num_in(m["min"], body), critical=True)
        check(f"{name}: Max_Salary {int(m['max'])}",
              num_in(m["max"], body), critical=True)
        # Salary_Range == Max - Min (the core derived rule).
        check(f"{name}: Salary_Range {int(m['range'])} (== Max-Min)",
              num_in(m["range"], body), critical=True)

    # ---- Highest / lowest avg paragraph ----
    print("\n=== Checking highest/lowest paragraph ===")
    check(f"Body names highest-avg dept '{high_name}' with value {high_m['avg']:.2f}",
          high_name.lower() in body_l and num_in(high_m["avg"], body), critical=True)
    check(f"Body names lowest-avg dept '{low_name}' with value {low_m['avg']:.2f}",
          low_name.lower() in body_l and num_in(low_m["avg"], body), critical=True)

    # ---- Email ----
    print("\n=== Checking email ===")
    cur.execute("SELECT to_addr, subject, body_text FROM email.messages")
    messages = cur.fetchall()
    found_email = False
    for to_addr, subject, btext in messages:
        to = str(to_addr).lower() if to_addr else ""
        subj = str(subject).lower() if subject else ""
        ebody = str(btext) if btext else ""
        ebody_l = ebody.lower()
        if "hr-director@company.com" in to and "workforce" in subj and "analysis" in subj:
            found_email = True
            check("Email subject is 'HR Department Workforce Analysis Ready'",
                  "hr department workforce analysis ready" in subj,
                  f"got '{subject}'")
            # CRITICAL: total employee count (~50000) stated.
            check("Email body states total employee count",
                  num_in(total_employees, ebody)
                  or "50000" in ebody_l or "50,000" in ebody_l,
                  f"expected total ~{total_employees}", critical=True)
            # CRITICAL: highest-avg department named (RU value from ClickHouse,
            # but accept the English original too in case the agent un-localizes).
            en_high = {
                "инженерия": "engineering", "финансы": "finance", "кадры": "hr",
                "операции": "operations", "ниокр": "r&d", "продажи": "sales",
                "поддержка": "support",
            }.get(high_name.lower(), "")
            named = (high_name.lower() in ebody_l) or (en_high and en_high in ebody_l)
            check("Email body names the highest-avg department (RU or EN)",
                  named, f"highest='{high_name}'", critical=True)
            # The highest-avg value should appear too.
            check("Email body includes the highest-avg salary value",
                  num_in(high_m["avg"], ebody),
                  f"expected ~{high_m['avg']:.2f}")
            break
    if not found_email:
        check("Email to hr-director@company.com with workforce-analysis subject sent",
              False, "no matching email found", critical=True)

    cur.close()
    conn.close()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    print("\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}, Accuracy: {accuracy:.1f}%")

    if CRITICAL_FAILED:
        print(f"  CRITICAL FAILURES ({len(CRITICAL_FAILED)}):")
        for n in CRITICAL_FAILED:
            print(f"    - {n}")
        print("  Overall: FAIL (critical)")
        sys.exit(1)

    if accuracy >= 70.0:
        print("  Overall: PASS")
        sys.exit(0)
    else:
        print("  Overall: FAIL (accuracy < 70%)")
        sys.exit(1)


if __name__ == "__main__":
    main()
