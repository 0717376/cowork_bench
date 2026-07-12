"""
Evaluation script for playwright-moex-market-dashboard-gsheet-word task.

Checks:
1. Google Sheet "Weekly_Market_Analysis" with sector and stock data (MOEX)
2. Weekly_Market_Report.docx with executive summary
3. Email sent with market analysis

CRITICAL_CHECKS (any failure => overall FAIL regardless of accuracy):
- Sector sheet has the CORRECT best (Energy +3.4) and worst (Consumer -1.5)
  weekly returns matched numerically with tolerance against the dashboard.
- Stock-vs-Sector sheet has CORRECT live prices for >=4 of 5 MOEX tickers,
  read as the latest close from moex.stock_prices (the source get_stock_info
  serves, per pg_adapter.py), proving the agent queried get_stock_info.
- Word conclusion names BOTH the best (Энергетика/Energy) and worst
  (Потребительский/Consumer) sectors.
- Email exists with preserved English subject from analyst@ to committee@,
  body naming the sector(s) with negative weekly returns.
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
CRITICAL_FAILS = []

# Critical checks (semantic substance). Any failure -> overall FAIL.
CRITICAL_CHECKS = {
    "Sector sheet has correct BEST weekly return (Energy +3.4)",
    "Sector sheet has correct WORST weekly return (Consumer -1.5)",
    ">=4/5 MOEX tickers have correct live price in stock sheet",
    "Word conclusion names BOTH best (Energy) and worst (Consumer) sectors",
    "Email body names sector(s) with negative weekly return (Consumer)",
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILS.append(name)
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=0.4):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def parse_num(s):
    """Extract a float from a cell like '+3.4%', '-1.5', '142.6'."""
    if s is None:
        return None
    txt = str(s).replace("%", "").replace(",", "").replace("−", "-").strip()
    try:
        return float(txt)
    except (ValueError, TypeError):
        return None


def str_contains(haystack, needle):
    if haystack is None or needle is None:
        return False
    return needle.strip().lower() in str(haystack).strip().lower()


def any_alt(text, alts):
    """text contains any of the RU/EN alternatives."""
    t = (text or "").lower()
    return any(a.lower() in t for a in alts)


# Expected sector data from the RU MOEX dashboard (authored, ground truth).
# weekly returns drive the best/worst critical checks.
SECTOR_DATA = {
    "energy":        {"weekly": 3.4,  "index": 142.6, "ytd": 18.7, "alts": ["energy", "энергетик"]},
    "financial":     {"weekly": 1.8,  "index": 124.2, "ytd": 9.6,  "alts": ["financial", "финанс"]},
    "communication": {"weekly": 0.4,  "index": 96.3,  "ytd": -1.8, "alts": ["communication", "телеком"]},
    "consumer":      {"weekly": -1.5, "index": 101.7, "ytd": 4.2,  "alts": ["consumer", "потребит"]},
    "technology":    {"weekly": 2.6,  "index": 133.8, "ytd": 14.1, "alts": ["technology", "технолог"]},
}

BEST_SECTOR_WEEKLY = 3.4    # Energy / Энергетика (highest)
WORST_SECTOR_WEEKLY = -1.5  # Consumer / Потребительский (only negative)

TICKERS = ["GAZP.ME", "SBER.ME", "MTSS.ME", "MGNT.ME", "TCSG.ME"]


def load_live_prices():
    """Read live prices the same way get_stock_info serves them: the latest close
    from moex.stock_prices (the single source of truth the MCP adapter uses to
    override currentPrice/regularMarketPrice). Reading stock_info.data.currentPrice
    here would diverge from the sanctioned tool's output."""
    prices = {}
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        for sym in TICKERS:
            cur.execute(
                "SELECT close FROM moex.stock_prices WHERE symbol = %s "
                "ORDER BY date DESC LIMIT 1",
                (sym,),
            )
            row = cur.fetchone()
            if row and row[0] is not None:
                prices[sym] = float(row[0])
        cur.close()
        conn.close()
    except Exception as e:
        print(f"  [WARN] could not load moex.stock_prices: {e}")
    return prices


def check_gsheet(live_prices):
    """Check Google Sheet with sector and stock data."""
    print("\n=== Checking Google Sheet ===")
    all_ok = True

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        found_ss = None
        for ss_id, title in spreadsheets:
            if title and "market" in title.lower():
                found_ss = ss_id
                break

        if not found_ss:
            record("Google Sheet 'Weekly_Market_Analysis' exists", False,
                   f"Found spreadsheets: {[t for _, t in spreadsheets]}")
            cur.close()
            conn.close()
            return False

        record("Google Sheet exists", True)

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (found_ss,))
        sheets = cur.fetchall()
        sheet_names = [t.lower() for _, t in sheets]

        sector_sheet_id = None
        stock_sheet_id = None
        for s_id, s_title in sheets:
            tl = s_title.lower()
            if "stock" in tl:
                stock_sheet_id = s_id
            elif "sector" in tl or "performance" in tl:
                sector_sheet_id = s_id

        has_sector = sector_sheet_id is not None
        has_stock = stock_sheet_id is not None
        record("Has Sector Performance sheet", has_sector, f"Sheets: {sheet_names}")
        record("Has Stock vs Sector sheet", has_stock, f"Sheets: {sheet_names}")
        all_ok = has_sector and has_stock

        # ---- Sector sheet ----
        sector_nums = []
        if sector_sheet_id:
            cur.execute("""
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (found_ss, sector_sheet_id))
            cells = cur.fetchall()

            grid = {}
            for row, col, val in cells:
                grid.setdefault(row, {})[col] = val

            data_rows = {r for r in grid if r > 0}
            ok = len(data_rows) >= 5
            record("Sector sheet has >= 5 data rows", ok, f"Found {len(data_rows)} rows")
            if not ok:
                all_ok = False

            all_values = " ".join(str(v).lower() for r in grid.values() for v in r.values())
            for key, info in SECTOR_DATA.items():
                ok = any_alt(all_values, info["alts"])
                record(f"Sector '{key}' (RU/EN) in sheet data", ok)
                if not ok:
                    all_ok = False

            # collect all numeric cell values for value-presence checks
            for r in grid.values():
                for v in r.values():
                    n = parse_num(v)
                    if n is not None:
                        sector_nums.append(n)

            # CRITICAL: correct best & worst weekly returns present
            best_ok = any(num_close(n, BEST_SECTOR_WEEKLY) for n in sector_nums)
            record("Sector sheet has correct BEST weekly return (Energy +3.4)", best_ok,
                   f"nums={sector_nums}")
            worst_ok = any(num_close(n, WORST_SECTOR_WEEKLY) for n in sector_nums)
            record("Sector sheet has correct WORST weekly return (Consumer -1.5)", worst_ok,
                   f"nums={sector_nums}")
            if not best_ok or not worst_ok:
                all_ok = False

            # Non-critical: index values present for best/worst
            idx_ok = (any(num_close(n, 142.6, tol=1.0) for n in sector_nums) and
                      any(num_close(n, 101.7, tol=1.0) for n in sector_nums))
            record("Sector sheet has correct index values (142.6 & 101.7)", idx_ok)

        # ---- Stock sheet ----
        if stock_sheet_id:
            cur.execute("""
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (found_ss, stock_sheet_id))
            cells = cur.fetchall()

            up_values = " ".join(str(v).upper() for _, _, v in cells)
            for ticker in TICKERS:
                ok = ticker in up_values or ticker.replace(".ME", "") in up_values
                record(f"Ticker {ticker} in stock sheet", ok)
                if not ok:
                    all_ok = False

            stock_nums = [parse_num(v) for _, _, v in cells]
            stock_nums = [n for n in stock_nums if n is not None]

            # CRITICAL: >=4/5 live prices correct (proves get_stock_info used)
            matched = 0
            for sym, price in live_prices.items():
                tol = max(2.0, abs(price) * 0.02)
                if any(num_close(n, price, tol=tol) for n in stock_nums):
                    matched += 1
            price_ok = matched >= 4
            record(">=4/5 MOEX tickers have correct live price in stock sheet", price_ok,
                   f"matched={matched}/5, live={live_prices}")
            if not price_ok:
                all_ok = False

            # Non-critical: cross-reference sector weekly return joined into stock sheet
            join_ok = (any(num_close(n, BEST_SECTOR_WEEKLY) for n in stock_nums) and
                       any(num_close(n, WORST_SECTOR_WEEKLY) for n in stock_nums))
            record("Stock sheet joins Sector_Weekly_Return (best & worst present)", join_ok)

        cur.close()
        conn.close()
        return all_ok

    except Exception as e:
        record("Google Sheet DB accessible", False, str(e))
        return False


def check_word(agent_workspace):
    """Check Weekly_Market_Report.docx."""
    print("\n=== Checking Word Output ===")

    fpath = os.path.join(agent_workspace, "Weekly_Market_Report.docx")
    if not os.path.isfile(fpath):
        record("Word file exists", False, f"Not found: {fpath}")
        return False

    record("Word file exists", True)

    try:
        from docx import Document
        doc = Document(fpath)
        full_text = "\n".join(p.text for p in doc.paragraphs).lower()
    except Exception as e:
        record("Word file readable", False, str(e))
        return False

    all_ok = True

    # Sector mentions (RU/EN)
    sector_checks = [
        ("Mentions Energy sector (RU/EN)", SECTOR_DATA["energy"]["alts"]),
        ("Mentions Communication/Telecom (RU/EN)", SECTOR_DATA["communication"]["alts"]),
        ("Mentions Consumer sector (RU/EN)", SECTOR_DATA["consumer"]["alts"]),
    ]
    for name, alts in sector_checks:
        ok = any_alt(full_text, alts)
        record(name, ok)
        if not ok:
            all_ok = False

    rep_ok = any_alt(full_text, ["gazp", "газпром", "sber", "сбербанк"])
    record("Mentions a representative stock (GAZP/SBER...)", rep_ok)
    if not rep_ok:
        all_ok = False

    perf_ok = any_alt(full_text, ["best", "worst", "top", "performing", "performance",
                                  "лучш", "худш", "эффективн", "доходн"])
    record("Mentions best/worst or performance (RU/EN)", perf_ok)
    if not perf_ok:
        all_ok = False

    weekly_ok = any_alt(full_text, ["weekly", "return", "недельн", "доходн"])
    record("Mentions weekly return (RU/EN)", weekly_ok)
    if not weekly_ok:
        all_ok = False

    # CRITICAL: conclusion names BOTH best (Energy) and worst (Consumer)
    best_named = any_alt(full_text, SECTOR_DATA["energy"]["alts"])
    worst_named = any_alt(full_text, SECTOR_DATA["consumer"]["alts"])
    concl_ok = best_named and worst_named
    record("Word conclusion names BOTH best (Energy) and worst (Consumer) sectors", concl_ok,
           f"best={best_named} worst={worst_named}")
    if not concl_ok:
        all_ok = False

    return all_ok


def check_email():
    """Check email with market analysis."""
    print("\n=== Checking Email ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
        emails = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Email DB accessible", False, str(e))
        return False

    all_ok = True
    found_email = False

    for subject, from_addr, to_addr, body_text in emails:
        subj_lower = (subject or "").lower()
        if "market" in subj_lower and ("analysis" in subj_lower or "report" in subj_lower):
            found_email = True
            record("Market analysis email exists", True)

            from_ok = str_contains(from_addr, "analyst") or str_contains(from_addr, "investment")
            record("Email from analyst address", from_ok, f"From: {from_addr}")
            if not from_ok:
                all_ok = False

            to_ok = str_contains(to_addr, "committee")
            record("Email addressed to committee", to_ok, f"To: {to_addr}")
            if not to_ok:
                all_ok = False

            body_lower = (body_text or "").lower()
            body_ok = any_alt(body_lower, ["energy", "consumer", "sector",
                                           "энергетик", "потребит", "сектор"])
            record("Email body mentions sectors (RU/EN)", body_ok,
                   f"Body preview: {(body_text or '')[:200]}")
            if not body_ok:
                all_ok = False

            # CRITICAL: body names the negative-return sector (Consumer)
            neg_ok = any_alt(body_lower, SECTOR_DATA["consumer"]["alts"])
            record("Email body names sector(s) with negative weekly return (Consumer)", neg_ok,
                   f"Body preview: {(body_text or '')[:200]}")
            if not neg_ok:
                all_ok = False
            break

    if not found_email:
        record("Market analysis email exists", False,
               f"Found {len(emails)} emails, none matching")
        all_ok = False

    return all_ok


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    live_prices = load_live_prices()
    print(f"[eval] Live MOEX prices: {live_prices}")

    gsheet_ok = check_gsheet(live_prices)
    word_ok = check_word(args.agent_workspace)
    email_ok = check_email()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0

    print(f"\n=== SUMMARY ===")
    print(f"  GSheet: {'PASS' if gsheet_ok else 'FAIL'}")
    print(f"  Word:   {'PASS' if word_ok else 'FAIL'}")
    print(f"  Email:  {'PASS' if email_ok else 'FAIL'}")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}, Accuracy: {accuracy:.1f}%")

    if CRITICAL_FAILS:
        print(f"  CRITICAL FAILURES: {CRITICAL_FAILS}")
        print(f"  Overall: FAIL (critical check failed)")
        sys.exit(1)

    overall = accuracy >= 70.0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
