"""Evaluation for sf-support-csat-gform-gsheet-email (ClickHouse + forms RU swap).

Checks:
1. Google Sheet "Support Center Performance Dashboard" with SLA_Compliance and Summary sheets
   (compliance rates & Avg_CSAT verified numerically via num_close; Summary Worst_SLA_Priority verified)
2. Google Forms "Customer Support Satisfaction Survey" with 4 questions and correct option content
3. Email analytics@ -> support-management@ with preserved English subject and concrete figures

CRITICAL_CHECKS reflect the task's substance: a single critical failure => overall
FAIL (sys.exit(1)) regardless of accuracy. The accuracy>=70 gate applies afterward.

NOTE: sf_data realia (priority values High/Medium/Low etc.) are russified CENTRALLY by
db/zzz_clickhouse_after_init.sql. Priority VALUES are kept EXACTLY English here so that
seed<->eval<->groundtruth stay in sync. The agent writes English priority labels.
"""
import argparse
import json
import os
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILS = []

# Actual SLA data from sf_data (SUPPORT_CENTER). Row counts unchanged by central
# russification (string map only substitutes realia, not row counts / numeric metrics).
SLA_DATA = {
    "high":   {"total": 6466,  "compliant": 778,  "rate": 12.03, "csat": 3.26},
    "medium": {"total": 15774, "compliant": 1645, "rate": 10.43, "csat": 3.26},
    "low":    {"total": 9348,  "compliant": 4204, "rate": 44.97, "csat": 3.25},
}
# Worst SLA compliance = Medium (10.43). Best CSAT is a tie (High 3.26 == Medium 3.26),
# so Best_Priority is NOT a strict critical check; only Worst_SLA_Priority=Medium is unambiguous.
WORST_SLA_PRIORITY = "medium"


def record(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    tag = "CRITICAL " if critical else ""
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {tag}{name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {tag}{name}: {str(detail)[:200]}")
        if critical:
            CRITICAL_FAILS.append(name)


def check(name, condition, detail=""):
    record(name, condition, detail, critical=False)


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _row_contains_number(values, target, tol):
    """True if any cell in the row is numerically within tol of target."""
    for v in values:
        if v is None:
            continue
        s = str(v).replace("%", "").replace(",", ".").strip()
        if num_close(s, target, tol):
            return True
    return False


def check_gsheet():
    print("\n=== Проверка 1: Google-таблица 'Support Center Performance Dashboard' ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE (title ILIKE '%support%' AND title ILIKE '%performance%')
           OR (title ILIKE '%support%' AND title ILIKE '%dashboard%')
           OR title ILIKE '%performance dashboard%'
    """)
    sheets = cur.fetchall()
    check("Таблица Support Center Performance Dashboard существует", len(sheets) >= 1,
          f"Found: {[s[1] for s in sheets]}")

    crit_rates = "SLA_Compliance: проценты соблюдения SLA верны (High~12.03, Medium~10.43, Low~44.97)"
    crit_csat = "SLA_Compliance: средний CSAT верен (High~3.26, Medium~3.26, Low~3.25)"
    crit_worst = "Summary: Worst_SLA_Priority = Medium (наименьший процент соблюдения SLA)"

    if not sheets:
        record(crit_rates, False, "no spreadsheet", critical=True)
        record(crit_csat, False, "no spreadsheet", critical=True)
        record(crit_worst, False, "no spreadsheet", critical=True)
        cur.close()
        conn.close()
        return False

    ss_id = sheets[0][0]

    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
    tab_rows = cur.fetchall()
    sheet_tabs = [r[1] for r in tab_rows]
    tab_id_by_kind = {}
    for tid, ttitle in tab_rows:
        tl = (ttitle or "").lower()
        if "sla" in tl and "sla" not in tab_id_by_kind:
            tab_id_by_kind["sla"] = tid
        if "summary" in tl and "summary" not in tab_id_by_kind:
            tab_id_by_kind["summary"] = tid

    has_sla = "sla" in tab_id_by_kind
    has_summary = "summary" in tab_id_by_kind
    check("Есть лист SLA_Compliance", has_sla, f"Tabs: {sheet_tabs}")
    check("Есть лист Summary", has_summary, f"Tabs: {sheet_tabs}")

    # ---- SLA_Compliance: load cells scoped to the SLA tab, grouped by row ----
    sla_rows = {}
    if has_sla:
        cur.execute("""
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
            ORDER BY row_index, col_index
        """, (ss_id, tab_id_by_kind["sla"]))
        for r_idx, c_idx, val in cur.fetchall():
            sla_rows.setdefault(r_idx, []).append(val)

    sla_all = " ".join(str(v) for vals in sla_rows.values() for v in vals if v is not None).lower()
    check("Лист SLA_Compliance содержит приоритет High", "high" in sla_all, "Not found")
    check("Лист SLA_Compliance содержит приоритет Medium", "medium" in sla_all, "Not found")
    check("Лист SLA_Compliance содержит приоритет Low", "low" in sla_all, "Not found")
    check("Лист SLA_Compliance содержит количество тикетов",
          any(str(v) in sla_all for v in [6466, 15774, 9348]), "Ticket counts not found")

    # Map each priority to its data row, then verify rate & CSAT numerically.
    def _find_row(priority):
        for vals in sla_rows.values():
            joined = " ".join(str(v) for v in vals if v is not None).lower()
            if priority in joined:
                return vals
        return None

    rates_ok = True
    csat_ok = True
    rate_detail, csat_detail = [], []
    for prio, data in SLA_DATA.items():
        row = _find_row(prio)
        if row is None:
            rates_ok = csat_ok = False
            rate_detail.append(f"{prio}: row missing")
            csat_detail.append(f"{prio}: row missing")
            continue
        if not _row_contains_number(row, data["rate"], tol=0.5):
            rates_ok = False
            rate_detail.append(f"{prio}: expected {data['rate']} in {row}")
        if not _row_contains_number(row, data["csat"], tol=0.1):
            csat_ok = False
            csat_detail.append(f"{prio}: expected {data['csat']} in {row}")
    record(crit_rates, has_sla and rates_ok, "; ".join(rate_detail), critical=True)
    record(crit_csat, has_sla and csat_ok, "; ".join(csat_detail), critical=True)

    # ---- Summary: verify Worst_SLA_Priority = Medium (scoped to Summary tab) ----
    summary_all = ""
    worst_row_vals = []
    if has_summary:
        cur.execute("""
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
            ORDER BY row_index, col_index
        """, (ss_id, tab_id_by_kind["summary"]))
        srows = {}
        for r_idx, c_idx, val in cur.fetchall():
            srows.setdefault(r_idx, []).append(val)
        summary_all = " ".join(str(v) for vals in srows.values() for v in vals if v is not None).lower()
        for vals in srows.values():
            joined = " ".join(str(v) for v in vals if v is not None).lower()
            if "worst" in joined or "наимень" in joined or "худш" in joined:
                worst_row_vals = [str(v).lower() for v in vals if v is not None]
                break

    worst_ok = has_summary and any(WORST_SLA_PRIORITY in v for v in worst_row_vals)
    record(crit_worst, worst_ok,
           f"Worst_SLA row: {worst_row_vals}", critical=True)

    cur.close()
    conn.close()
    return has_sla and has_summary


def _option_values(config):
    """Extract option text values from a question config (RU forms-mcp / gform shape)."""
    vals = []
    if not config:
        return vals
    cfg = config if isinstance(config, dict) else None
    if cfg is None:
        try:
            cfg = json.loads(config)
        except Exception:
            return vals
    opts = cfg.get("options") if isinstance(cfg, dict) else None
    if isinstance(opts, list):
        for o in opts:
            if isinstance(o, dict):
                v = o.get("value")
                if v is not None:
                    vals.append(str(v))
            else:
                vals.append(str(o))
    return vals


def check_gform():
    print("\n=== Проверка 2: Google Forms 'Customer Support Satisfaction Survey' ===")
    crit_opts = ("Форма: Q1 содержит варианты Yes/No/Partially И есть Yes/No вопрос "
                 "(контент вариантов, не только число)")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, title FROM gform.forms
        WHERE title ILIKE '%customer%support%'
           OR title ILIKE '%support%satisfaction%'
           OR title ILIKE '%satisfaction survey%'
           OR title ILIKE '%удовлетвор%'
           OR title ILIKE '%опрос%поддержк%'
    """)
    forms = cur.fetchall()
    if not forms:
        # Fallback: any form (RU title could differ); pick the first.
        cur.execute("SELECT id, title FROM gform.forms")
        forms = cur.fetchall()
    check("Опрос Customer Support Satisfaction Survey существует", len(forms) >= 1,
          f"Found: {[f[1] for f in forms]}")

    if not forms:
        record(crit_opts, False, "no form", critical=True)
        cur.close()
        conn.close()
        return

    form_id = forms[0][0]
    cur.execute("""
        SELECT title, question_type, config FROM gform.questions
        WHERE form_id = %s ORDER BY position
    """, (form_id,))
    questions = cur.fetchall()
    cur.close()
    conn.close()

    check("Форма содержит ровно 4 вопроса", len(questions) == 4, f"Got {len(questions)}")

    parsed = []
    for q_title, q_type, q_config in questions:
        opts = [v.lower() for v in _option_values(q_config)]
        parsed.append({"title": (q_title or "").lower(), "type": q_type, "options": opts})

    # Q1: resolved satisfactorily -> Yes / No / Partially
    has_ynp = any(
        any("yes" in o or "да" in o for o in p["options"])
        and any(o == "no" or "no" in o or "нет" in o for o in p["options"])
        and any("partial" in o or "частич" in o for o in p["options"])
        for p in parsed
    )
    # A final Yes/No question (contact again) -> exactly Yes + No present, no Partially
    has_yes_no = any(
        any("yes" in o or "да" in o for o in p["options"])
        and any(o == "no" or "нет" in o for o in p["options"])
        and not any("partial" in o or "частич" in o for o in p["options"])
        for p in parsed
    )
    record(crit_opts, has_ynp and has_yes_no,
           f"has_ynp={has_ynp} has_yes_no={has_yes_no} options={[p['options'] for p in parsed]}",
           critical=True)


def check_email():
    print("\n=== Проверка 3: Письмо ===")
    crit_email = ("Письмо analytics@ -> support-management@ с темой 'Support Center Performance "
                  "Report - SLA Compliance Analysis' и конкретными цифрами SLA в теле")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
        WHERE (subject ILIKE '%support%' AND subject ILIKE '%performance%')
           OR subject ILIKE '%sla compliance%'
           OR to_addr::text ILIKE '%support-management@company.example.com%'
        ORDER BY date DESC
    """)
    emails = cur.fetchall()
    cur.close()
    conn.close()

    check("Письмо о работе центра поддержки найдено", len(emails) >= 1, f"Found {len(emails)}")

    if not emails:
        record(crit_email, False, "no email", critical=True)
        return

    e = emails[0]
    to_str = str(e[2]).lower()
    from_str = (e[1] or "").lower()
    subj = (e[0] or "").lower()
    body = (e[3] or "").lower()

    check("Письмо на support-management@company.example.com",
          "support-management@company.example.com" in to_str, f"to: {to_str}")
    check("Письмо с analytics@company.example.com",
          "analytics@company.example.com" in from_str, f"from: {e[1]}")
    check("Тело письма упоминает SLA/соблюдение/CSAT",
          any(kw in body for kw in ["sla", "compliance", "соблюд", "csat", "satisfaction",
                                    "удовлетвор", "high", "medium", "low"]),
          "Body missing key terms")

    # CRITICAL: correct addresses + preserved English subject + concrete figures in body.
    to_ok = "support-management@company.example.com" in to_str
    from_ok = "analytics@company.example.com" in from_str
    subj_ok = "support" in subj and ("performance" in subj or "sla compliance" in subj)
    # Concrete figures: at least one correct rate number, OR all three priority levels named.
    rate_strs = ["12.03", "10.43", "44.97"]
    has_rate = any(rs in body for rs in rate_strs)
    has_all_prios = all(p in body for p in ["high", "medium", "low"])
    body_ok = has_rate or has_all_prios
    record(crit_email, to_ok and from_ok and subj_ok and body_ok,
           f"to={to_ok} from={from_ok} subj={subj_ok} has_rate={has_rate} "
           f"has_all_prios={has_all_prios}", critical=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=== Evaluation: sf-support-csat-gform-gsheet-email (ClickHouse + forms) ===")

    check_gsheet()
    check_gform()
    check_email()

    total = PASS_COUNT + FAIL_COUNT
    print(f"\n=== SUMMARY: {PASS_COUNT} passed, {FAIL_COUNT} failed ===")

    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"Overall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
        "critical_failures": CRITICAL_FAILS,
    }
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if CRITICAL_FAILS:
        print(f"\nFAIL: critical checks failed: {CRITICAL_FAILS}")
        sys.exit(1)

    if accuracy >= 70:
        print("\nPASS")
        sys.exit(0)
    print("\nFAIL: accuracy below 70%")
    sys.exit(1)


if __name__ == "__main__":
    main()
