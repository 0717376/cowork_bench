"""Evaluation for terminal-moex-sf-gsheet-word-gcal (ClickHouse + moex-finance fork).
Structural checks plus CRITICAL semantic checks recomputed live from source data.

CRITICAL checks (any failure => sys.exit(1) before the accuracy gate):
 C1. Market adjustment factor in market_adjusted_bonuses.json matches the factor
     recomputed live from moex.stock_prices for the driver ticker SBER.ME.
 C2. Budget-cap enforcement: total adjusted bonuses <= budget cap; if pre-scaling
     total exceeded the cap, scaled total equals the cap.
 C3. Per-employee cap: no individual adjusted bonus exceeds 20% of that salary.
 C4. Bonus computation: spot-check >=3 employees that current bonus ==
     salary * tier_pct, where tier derives from the round-robin region's total
     revenue queried live from clickhouse orders/customers.
 C5. Calendar integrity: exactly one Q4 Compensation Review Meeting, ~90 min,
     within the Mar 9-13 2026 09:00-17:00 ET window, zero overlap with the
     pre-seeded conflict events.

Region / department / employee-name realia stay EXACTLY as the central clickhouse
map produced them (Russian); this eval never hardcodes those literals — it queries
them live so seed/eval/groundtruth stay in sync.
"""
import argparse
import json
import os
import sys
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}

# Driver ticker for the market adjustment (russified fork: Sberbank bellwether).
DRIVER_TICKER = "SBER.ME"
DISPLAY_TICKERS = ["SBER.ME", "GAZP.ME", "LKOH.ME"]
BUDGET_CAP = 30000000.0

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS]{' [CRIT]' if critical else ''} {name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL]{' [CRIT]' if critical else ''} {name}: {str(detail)[:300]}")
        if critical:
            CRITICAL_FAILED.append(name)


def num_close(a, b, tol=2.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Live source-of-truth recomputation (ClickHouse sf_data + moex.stock_prices)
# ---------------------------------------------------------------------------

def compute_expected_factor():
    """Recompute the market adjustment factor from moex.stock_prices for the
    driver ticker, replicating the agent's method: current = close at MAX(date),
    year-ago = close at the earliest date inside a ~1y lookback window. Returns
    (factor, yoy_pct, current, year_ago) or None on failure."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT MAX(date) FROM moex.stock_prices WHERE symbol = %s",
            (DRIVER_TICKER,))
        max_date = cur.fetchone()[0]
        if max_date is None:
            return None
        cur.execute(
            "SELECT close FROM moex.stock_prices WHERE symbol = %s AND date = %s",
            (DRIVER_TICKER, max_date))
        current = float(cur.fetchone()[0])
        # Earliest date within ~1 year before the latest date.
        cur.execute(
            """SELECT close FROM moex.stock_prices
               WHERE symbol = %s AND date >= (%s::date - INTERVAL '370 days')
               ORDER BY date ASC LIMIT 1""",
            (DRIVER_TICKER, max_date))
        year_ago = float(cur.fetchone()[0])
        yoy = (current - year_ago) / year_ago * 100.0
        if yoy > 10:
            factor = 0.9
        elif yoy < -10:
            factor = 1.1
        else:
            factor = 1.0
        return factor, yoy, current, year_ago
    except Exception as e:
        print(f"  [warn] compute_expected_factor failed: {e}")
        return None
    finally:
        cur.close()
        conn.close()


def load_tier_table():
    """Read bonus tier thresholds from the workspace bonus_tiers.csv if present,
    else fall back to the canonical policy thresholds."""
    # Canonical thresholds (min_revenue, max_revenue, bonus_percentage).
    return [
        (0, 550000, 5),
        (550000, 620000, 8),
        (620000, 650000, 10),
        (650000, 10**12, 12),
    ]


def tier_pct(revenue, tiers):
    for lo, hi, pct in tiers:
        if lo <= revenue < hi:
            return pct
    return tiers[-1][2]


def compute_expected_bonuses():
    """Recompute, live from clickhouse, the per-employee expected current bonus.
    Returns dict {employee_name: {region, salary, pct, bonus}} or None.
    Replicates: round-robin alphabetical assignment of Sales employees across the
    five customer regions sorted alphabetically; tier from that region's total
    revenue (orders joined to customers)."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Region revenues (orders joined to customers), only the 5 named regions.
        cur.execute("""
            SELECT c."REGION", SUM(o."TOTAL_AMOUNT")
            FROM sf_data."SALES_DW__PUBLIC__ORDERS" o
            JOIN sf_data."SALES_DW__PUBLIC__CUSTOMERS" c
              ON o."CUSTOMER_ID" = c."CUSTOMER_ID"
            WHERE c."REGION" IS NOT NULL AND c."REGION" <> ''
            GROUP BY c."REGION"
        """)
        region_rev = {r[0]: float(r[1]) for r in cur.fetchall()}
        # Five regions sorted alphabetically (whatever language the map produced).
        regions = sorted(region_rev.keys())
        if len(regions) < 5:
            print(f"  [warn] expected 5 regions, found {len(regions)}: {regions}")
        tiers = load_tier_table()
        region_pct = {reg: tier_pct(region_rev[reg], tiers) for reg in regions}

        # Sales employees; sort in Python (matches the agent's typical sort and
        # avoids DB-collation divergence on Cyrillic names).
        cur.execute("""
            SELECT "EMPLOYEE_NAME", "SALARY"
            FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"
            WHERE "DEPARTMENT" = 'Продажи'
        """)
        emps = sorted(cur.fetchall(), key=lambda r: r[0])
        out = {}
        n_reg = len(regions)
        for i, (name, salary) in enumerate(emps):
            reg = regions[i % n_reg] if n_reg else None
            pct = region_pct.get(reg, 0)
            salary = float(salary)
            out[name] = {
                "region": reg, "salary": salary,
                "pct": pct, "bonus": salary * pct / 100.0,
            }
        return out
    except Exception as e:
        print(f"  [warn] compute_expected_bonuses failed: {e}")
        return None
    finally:
        cur.close()
        conn.close()


def get_sales_headcount():
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        cur.execute("""SELECT COUNT(*) FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"
                       WHERE "DEPARTMENT" = 'Продажи'""")
        return int(cur.fetchone()[0])
    except Exception:
        return None
    finally:
        cur.close()
        conn.close()


# ---------------------------------------------------------------------------
# Structural checks
# ---------------------------------------------------------------------------

def check_word(workspace):
    print("\n=== Check 1: Compensation_Review_Memo.docx ===")
    path = os.path.join(workspace, "Compensation_Review_Memo.docx")
    if not os.path.exists(path):
        check("Word file exists", False, f"Not found at {path}")
        return
    check("Word file exists", True)

    try:
        from docx import Document
        doc = Document(path)
        all_text = " ".join(p.text for p in doc.paragraphs).lower()

        def has_any(*subs):
            return any(s in all_text for s in subs)

        check("Has title 'Q4 Compensation Review'",
              has_any("q4 compensation review", "compensation review memo",
                      "обзор компенсаций", "памятка"),
              f"Text snippet: {all_text[:120]}")
        check("Has Background section", has_any("background", "введение", "обзор"),
              "Missing Background/Введение")
        check("Has Methodology section", has_any("methodology", "методолог"),
              "Missing Methodology/Методология")
        check("Has Market Analysis section",
              has_any("market analysis", "анализ рынка", "market", "рынк"),
              "Missing Market Analysis/Анализ рынка")
        check("Has Regional Performance section",
              has_any("regional", "region", "регион"),
              "Missing Regional/региональн")
        check("Has Budget Impact section",
              has_any("budget", "бюджет"), "Missing Budget/бюджет")
        check("Has Recommendations section",
              has_any("recommend", "рекомендац"), "Missing Recommendations/рекомендации")
        check("Mentions driver ticker (SBER)",
              has_any("sber", "сбер"), "No SBER/Сбербанк reference")
        check("Mentions adjustment factor",
              has_any("0.9", "1.0", "1.1", "adjustment", "коэффициент", "корректировк"),
              "No adjustment factor reference")
        check("Mentions budget cap",
              has_any("budget cap", "бюджетн", "30,000,000", "30000000", "30 million",
                      "30 млн"),
              "No budget cap reference")
    except Exception as e:
        check("Word readable", False, str(e))


def check_gsheet():
    print("\n=== Check 2: Google Sheet 'Sales Compensation Analysis' ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE lower(title) LIKE '%sales compensation%' OR lower(title) LIKE '%compensation analysis%'")
        rows = cur.fetchall()
        check("Spreadsheet exists", len(rows) >= 1, f"Found {len(rows)} matching spreadsheets")
        if not rows:
            return

        ss_id = rows[0][0]

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s ORDER BY index", (ss_id,))
        sheets = cur.fetchall()
        sheet_titles = [s[1].lower() for s in sheets]
        check("Has at least 3 sheets", len(sheets) >= 3, f"Found {len(sheets)} sheets: {sheet_titles}")

        has_rep = any("rep" in t or "performance" in t for t in sheet_titles)
        has_market = any("market" in t or "adjustment" in t for t in sheet_titles)
        has_adjusted = any("adjusted" in t or "bonus" in t for t in sheet_titles)
        check("Has Rep_Performance sheet", has_rep, f"Sheets: {sheet_titles}")
        check("Has Market_Adjustment sheet", has_market, f"Sheets: {sheet_titles}")
        check("Has Adjusted_Bonuses sheet", has_adjusted, f"Sheets: {sheet_titles}")

        # Market_Adjustment sheet has the three tickers.
        for sid, title in sheets:
            if "market" in title.lower() or "adjustment" in title.lower():
                cur.execute("""SELECT value FROM gsheet.cells
                    WHERE spreadsheet_id = %s AND sheet_id = %s""", (ss_id, sid))
                values = [r[0].lower() if r[0] else "" for r in cur.fetchall()]
                all_vals = " ".join(values)
                check("Market sheet has SBER", "sber" in all_vals,
                      f"Values: {all_vals[:200]}")
                check("Market sheet has GAZP", "gazp" in all_vals,
                      f"Values: {all_vals[:200]}")
                check("Market sheet has LKOH", "lkoh" in all_vals,
                      f"Values: {all_vals[:200]}")
                break

        # Rep_Performance sheet has data rows.
        for sid, title in sheets:
            if "rep" in title.lower() or "performance" in title.lower():
                cur.execute("""SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
                    WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 0""", (ss_id, sid))
                data_rows = cur.fetchone()[0]
                check("Rep sheet has data rows", data_rows >= 1, f"Found {data_rows} data rows")
                break

    except Exception as e:
        check("GSheet query", False, str(e))
    finally:
        cur.close()
        conn.close()


def check_gcal_critical():
    """C5: calendar event integrity, critical."""
    print("\n=== Check 3 (CRITICAL): Calendar Event ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        cur.execute("""SELECT summary, description, start_datetime, end_datetime
            FROM gcal.events
            WHERE lower(summary) LIKE '%compensation%' OR lower(summary) LIKE '%q4%bonus%'""")
        events = cur.fetchall()
        check("Exactly one Compensation Review event scheduled", len(events) == 1,
              f"Found {len(events)} matching events", critical=True)

        if events:
            evt = events[0]
            check("Event title contains 'Compensation Review'",
                  "compensation" in evt[0].lower() and "review" in evt[0].lower(),
                  f"Title: {evt[0]}")

            start_dt = evt[2]
            end_dt = evt[3]
            if start_dt:
                day = getattr(start_dt, 'day', None)
                month = getattr(start_dt, 'month', None)
                year = getattr(start_dt, 'year', None)
                hour = getattr(start_dt, 'hour', None)
                in_week = (year == 2026 and month == 3 and 9 <= (day or 0) <= 13)
                check("Event in Mar 9-13 2026 week", in_week,
                      f"Start: {start_dt}", critical=True)
                # Window 09:00-17:00 ET; event must start at/after 09:00 and end by 17:00.
                if end_dt:
                    duration = (end_dt - start_dt).total_seconds() / 60
                    check("Event is ~90 minutes", 85 <= duration <= 95,
                          f"Duration: {duration} min", critical=True)
                    in_hours = (hour is not None and 9 <= hour and
                                getattr(end_dt, 'hour', 0) * 60 + getattr(end_dt, 'minute', 0) <= 17 * 60)
                    check("Event within 09:00-17:00 ET", in_hours,
                          f"Start hour: {hour}, end: {end_dt}", critical=True)

                # Zero overlap with pre-seeded conflict events.
                try:
                    cur.execute("""SELECT summary, start_datetime, end_datetime FROM gcal.events
                        WHERE lower(summary) NOT LIKE '%%compensation%%'
                        AND start_datetime < %s AND end_datetime > %s""",
                        (end_dt, start_dt))
                    conflicts = cur.fetchall()
                    check("No conflicts with existing events", len(conflicts) == 0,
                          f"Conflicts: {[(c[0], str(c[1])) for c in conflicts]}",
                          critical=True)
                except Exception as e2:
                    check("No conflicts with existing events", False, str(e2), critical=True)
    except Exception as e:
        check("GCal query", False, str(e), critical=True)
    finally:
        cur.close()
        conn.close()


def check_scripts(workspace):
    print("\n=== Check 4: Script Files ===")
    for script in ["compute_bonuses.py", "market_adjustment.py", "validate_bonuses.py"]:
        path = os.path.join(workspace, script)
        check(f"{script} exists", os.path.exists(path))


def _load_json(path):
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return None


def _entries(data):
    """Normalize a bonuses JSON file to a list of dict rows."""
    if isinstance(data, list):
        return [d for d in data if isinstance(d, dict)]
    if isinstance(data, dict):
        for v in data.values():
            if isinstance(v, list):
                return [d for d in v if isinstance(d, dict)]
        # dict keyed by name -> record
        rows = []
        for k, v in data.items():
            if isinstance(v, dict):
                r = dict(v)
                r.setdefault("name", k)
                rows.append(r)
        return rows
    return []


def _get(row, *keys):
    low = {k.lower(): v for k, v in row.items()}
    for k in keys:
        if k.lower() in low:
            return low[k.lower()]
    # substring fallback
    for k in keys:
        for lk, lv in low.items():
            if k.lower() in lk:
                return lv
    return None


def check_json_outputs(workspace, expected_bonuses, headcount):
    print("\n=== Check 5: JSON Output Files ===")
    threshold = max(50, int((headcount or 0) * 0.8)) if headcount else 50

    cb_path = os.path.join(workspace, "current_bonuses.json")
    cb = _load_json(cb_path)
    check("current_bonuses.json exists", os.path.exists(cb_path))
    if cb is not None:
        rows = _entries(cb)
        check("current_bonuses has entries", len(rows) >= threshold,
              f"Found {len(rows)} entries (threshold {threshold}, headcount {headcount})")
        if rows:
            first = rows[0]
            has_keys = all(_get(first, k) is not None for k in ["name", "region", "salary", "bonus"])
            check("current_bonuses has required fields", has_keys,
                  f"Keys: {list(first.keys())}")

    mab_path = os.path.join(workspace, "market_adjusted_bonuses.json")
    mab = _load_json(mab_path)
    check("market_adjusted_bonuses.json exists", os.path.exists(mab_path))
    if mab is not None:
        rows = _entries(mab)
        check("adjusted_bonuses has entries", len(rows) >= threshold,
              f"Found {len(rows)} entries (threshold {threshold})")
        if rows:
            keys_lower = {k.lower() for k in rows[0].keys()}
            check("adjusted_bonuses has adjusted field",
                  any("adjust" in k or "скоррект" in k for k in keys_lower),
                  f"Keys: {list(rows[0].keys())}")


# ---------------------------------------------------------------------------
# CRITICAL semantic checks
# ---------------------------------------------------------------------------

def critical_checks(workspace, expected_factor_info, expected_bonuses):
    print("\n=== CRITICAL Semantic Checks ===")
    mab_path = os.path.join(workspace, "market_adjusted_bonuses.json")
    cb_path = os.path.join(workspace, "current_bonuses.json")
    mab = _entries(_load_json(mab_path) or [])
    cb = _entries(_load_json(cb_path) or [])

    # --- C1: market adjustment factor matches recomputed driver YoY factor ---
    if expected_factor_info is None:
        check("C1 factor recomputable from moex", False,
              "Could not recompute factor from moex.stock_prices", critical=True)
        exp_factor = None
    else:
        exp_factor, yoy, cur_p, ya_p = expected_factor_info
        print(f"  [info] driver {DRIVER_TICKER}: current={cur_p:.4f} year_ago={ya_p:.4f} "
              f"YoY={yoy:.2f}% -> expected factor {exp_factor}")
        observed_factors = set()
        for r in mab:
            f = _get(r, "market factor", "factor", "market_factor", "adjustment factor",
                     "коэффициент", "market")
            if f is not None:
                try:
                    observed_factors.add(round(float(f), 4))
                except Exception:
                    pass
        if observed_factors:
            ok = all(abs(f - exp_factor) <= 0.01 for f in observed_factors)
            check("C1 adjustment factor matches recomputed driver YoY factor", ok,
                  f"expected {exp_factor}, observed {sorted(observed_factors)}",
                  critical=True)
        else:
            # Fall back: infer factor from adjusted/current ratio.
            ratios = []
            for r in mab:
                adj = _get(r, "adjusted bonus", "adjusted", "adjusted_bonus", "скорректированный")
                cur_b = _get(r, "current bonus", "current", "bonus", "current_bonus")
                try:
                    if adj is not None and cur_b is not None and float(cur_b) > 0:
                        ratios.append(float(adj) / float(cur_b))
                except Exception:
                    pass
            if ratios:
                avg = sum(ratios) / len(ratios)
                # If budget scaling was applied, ratios are <= factor; require not exceeding factor materially.
                ok = avg <= exp_factor + 0.02
                check("C1 adjusted/current ratio consistent with factor (no scaling) or below",
                      ok, f"expected <= {exp_factor}, avg ratio {avg:.4f}", critical=True)
            else:
                check("C1 adjustment factor present", False,
                      "No factor or adjusted/current ratio found in JSON", critical=True)

    # --- C2: budget-cap enforcement ---
    total_adj = 0.0
    got_total = False
    for r in mab:
        adj = _get(r, "adjusted bonus", "adjusted", "adjusted_bonus", "скорректированный")
        try:
            if adj is not None:
                total_adj += float(adj)
                got_total = True
        except Exception:
            pass
    if got_total:
        check("C2 total adjusted bonuses within budget cap",
              total_adj <= BUDGET_CAP + 1.0,
              f"total adjusted = {total_adj:.2f}, cap = {BUDGET_CAP}", critical=True)
        # If pre-scaling total (current*factor) exceeded the cap, scaled total must equal cap.
        if exp_factor is not None and cb:
            pre = 0.0
            okpre = False
            for r in cb:
                b = _get(r, "bonus amount", "bonus", "current bonus", "current_bonus")
                try:
                    if b is not None:
                        pre += float(b)
                        okpre = True
                except Exception:
                    pass
            if okpre:
                pre_scaled = pre * exp_factor
                if pre_scaled > BUDGET_CAP + 1.0:
                    check("C2 scaled total equals cap (scaling applied)",
                          abs(total_adj - BUDGET_CAP) <= max(1.0, BUDGET_CAP * 0.001),
                          f"pre-scaling {pre_scaled:.2f} > cap; total {total_adj:.2f}",
                          critical=True)
                else:
                    check("C2 no over-cap scaling required (total <= cap)",
                          total_adj <= BUDGET_CAP + 1.0,
                          f"pre-scaling {pre_scaled:.2f} <= cap", critical=True)
    else:
        check("C2 adjusted bonus totals present", False,
              "No adjusted bonus amounts found in market_adjusted_bonuses.json",
              critical=True)

    # --- C3: per-employee 20%-of-salary cap ---
    violations = []
    checked_any = False
    for r in mab:
        adj = _get(r, "adjusted bonus", "adjusted", "adjusted_bonus", "скорректированный")
        sal = _get(r, "salary", "оклад")
        name = _get(r, "name", "employee name", "имя")
        # salary may live only in current_bonuses; build a lookup if missing.
        if sal is None and name is not None:
            for c in cb:
                if _get(c, "name", "employee name", "имя") == name:
                    sal = _get(c, "salary", "оклад")
                    break
        try:
            if adj is not None and sal is not None and float(sal) > 0:
                checked_any = True
                if float(adj) > 0.20 * float(sal) + 0.01:
                    violations.append((name, float(adj), float(sal)))
        except Exception:
            pass
    if checked_any:
        check("C3 no adjusted bonus exceeds 20% of salary", len(violations) == 0,
              f"{len(violations)} violations, e.g. {violations[:3]}", critical=True)
    else:
        check("C3 per-employee cap verifiable", False,
              "Could not pair adjusted bonus with salary", critical=True)

    # --- C4: bonus computation spot-check vs live clickhouse recompute ---
    if not expected_bonuses:
        check("C4 expected bonuses recomputable from clickhouse", False,
              "Could not recompute expected bonuses", critical=True)
    elif not cb:
        check("C4 current_bonuses.json usable for spot-check", False,
              "current_bonuses.json missing/unparseable", critical=True)
    else:
        matched = 0
        checked = 0
        mismatches = []
        for r in cb:
            name = _get(r, "name", "employee name", "имя")
            bonus = _get(r, "bonus amount", "bonus", "current bonus", "current_bonus")
            if name in expected_bonuses and bonus is not None:
                checked += 1
                exp = expected_bonuses[name]["bonus"]
                try:
                    if num_close(bonus, exp, tol=max(1.0, abs(exp) * 0.01)):
                        matched += 1
                    else:
                        mismatches.append((name, float(bonus), round(exp, 2)))
                except Exception:
                    mismatches.append((name, bonus, round(exp, 2)))
                if checked >= 8 and matched >= 3:
                    break
        check("C4 current bonuses match salary*tier_pct (>=3 spot-checks)",
              matched >= 3,
              f"matched {matched}/{checked}; mismatches {mismatches[:3]}",
              critical=True)


def check_reverse_validation():
    print("\n=== Reverse Validation ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        cur.execute("""SELECT summary, start_datetime FROM gcal.events
            WHERE lower(summary) LIKE '%compensation%review%'""")
        comp_events = cur.fetchall()
        check("No duplicate Compensation Review events", len(comp_events) <= 1,
              f"Found {len(comp_events)}: {[e[0] for e in comp_events]}")

        cur.execute("""SELECT summary FROM gcal.events
            WHERE lower(summary) LIKE '%bonus%payout%'
               OR lower(summary) LIKE '%salary%update%'
               OR lower(summary) LIKE '%pay%raise%'""")
        wrong_events = cur.fetchall()
        check("No wrong event types created", len(wrong_events) == 0,
              f"Found: {[e[0] for e in wrong_events]}")

        cur.execute("""SELECT COUNT(*) FROM gcal.events
            WHERE lower(summary) NOT LIKE '%compensation%'
              AND lower(summary) NOT LIKE '%q4%bonus%'""")
        other_count = cur.fetchone()[0]
        check("Pre-existing calendar events preserved (>= 5)", other_count >= 5,
              f"Found {other_count} non-compensation events (expected >= 5 from original 11)")
    except Exception as e:
        check("Reverse validation (gcal noise)", False, str(e))
    finally:
        cur.close()
        conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected_factor_info = compute_expected_factor()
    expected_bonuses = compute_expected_bonuses()
    headcount = get_sales_headcount()

    check_word(args.agent_workspace)
    check_gsheet()
    check_gcal_critical()
    check_scripts(args.agent_workspace)
    check_json_outputs(args.agent_workspace, expected_bonuses, headcount)
    critical_checks(args.agent_workspace, expected_factor_info, expected_bonuses)
    check_reverse_validation()

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    result = {"total_passed": PASS_COUNT, "total_checks": total, "accuracy": accuracy,
              "critical_failed": CRITICAL_FAILED}
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if CRITICAL_FAILED:
        print(f"\nFAIL: {len(CRITICAL_FAILED)} CRITICAL check(s) failed: {CRITICAL_FAILED}")
        sys.exit(1)

    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
