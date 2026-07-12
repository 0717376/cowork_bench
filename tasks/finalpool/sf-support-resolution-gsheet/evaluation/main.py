"""Evaluation for sf-support-resolution-gsheet (ClickHouse warehouse, sf_data schema).

Scoring model: collect (name, passed, critical) checks.
- Any critical check failing => FAIL (sys.exit(1)) before the accuracy gate.
- Otherwise PASS iff accuracy >= 70%.

ISSUE_TYPE values are russified centrally by db/zzz_clickhouse_after_init.sql
(Bug->Ошибка, ...). PRIORITY (High/Medium/Low) and STATUS (Closed) are NOT
russified and stay English.
"""
import argparse
import os
import sys
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}

# Russian issue-type literals as produced by the central russification map.
EXPECTED_ISSUE_TYPES = {
    "Ошибка":                       {"total": 6804, "avg_res_hrs": 16.0, "avg_sat": 2.96},
    "Проблема производительности":  {"total": 8128, "avg_res_hrs": 15.2, "avg_sat": 3.38},
    "Техническая проблема":         {"total": 1515, "avg_res_hrs": 14.7, "avg_sat": 3.35},
    "Обслуживание":                 {"total": 1558, "avg_res_hrs": 14.7, "avg_sat": 3.33},
    "Инцидент":                     {"total": 4463, "avg_res_hrs": 14.6, "avg_sat": 3.31},
    "Запрос функции":               {"total": 6118, "avg_res_hrs": 14.6, "avg_sat": 3.31},
    "Запрос обслуживания":          {"total": 3002, "avg_res_hrs": 14.3, "avg_sat": 3.31},
}

# Issue type with the LONGEST average resolution hours (the email headline).
LONGEST_ISSUE_TYPE = "Ошибка"  # Bug, ~16.0h

# PRIORITY values are intentionally NOT russified.
EXPECTED_PRIORITIES = {
    "High":   {"count": 6466,  "avg_response": 6.2,  "sla_pct": 100.0},
    "Low":    {"count": 9348,  "avg_response": 25.8, "sla_pct": 100.0},
    "Medium": {"count": 15774, "avg_response": 12.3, "sla_pct": 100.0},
}


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_contains(haystack, needle):
    if haystack is None or needle is None:
        return False
    return needle.lower() in str(haystack).lower()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    # checks: list of (name, passed, is_critical)
    checks = []

    def add(name, passed, critical=False):
        checks.append((name, bool(passed), critical))

    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except Exception as e:
        print(f"=== RESULT: FAIL (could not connect to PostgreSQL: {e}) ===")
        sys.exit(1)

    # 1. Spreadsheet exists (CRITICAL core deliverable)
    cur.execute("SELECT id FROM gsheet.spreadsheets WHERE title ILIKE '%Support Resolution Analysis%'")
    ss_rows = cur.fetchall()
    if not ss_rows:
        add("Spreadsheet 'Support Resolution Analysis' exists", False, critical=True)
        ss_id = None
    else:
        add("Spreadsheet 'Support Resolution Analysis' exists", True, critical=False)
        ss_id = ss_rows[0][0]

    issue_grid = {}
    if ss_id is not None:
        # 2. "By Issue Type" sheet
        cur.execute("SELECT id FROM gsheet.sheets WHERE spreadsheet_id = %s AND title ILIKE '%%Issue Type%%'", (ss_id,))
        sheet_rows = cur.fetchall()
        add("'By Issue Type' sheet exists", bool(sheet_rows), critical=False)
        if sheet_rows:
            sheet_id = sheet_rows[0][0]
            cur.execute("""
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (ss_id, sheet_id))
            for row_idx, col_idx, value in cur.fetchall():
                issue_grid.setdefault(row_idx, {})[col_idx] = value

            sorted_rows = sorted(issue_grid.keys())
            data_rows = sorted_rows[1:] if len(sorted_rows) > 1 else []

            # map russian issue type -> data row values
            type_rows = {}
            for row_idx in data_rows:
                row = issue_grid[row_idx]
                cols = sorted(row.keys())
                if len(cols) < 5:
                    continue
                issue_type = str(row[cols[0]] or "")
                for exp_type in EXPECTED_ISSUE_TYPES:
                    if str_contains(issue_type, exp_type):
                        type_rows[exp_type] = [row[c] for c in cols]

            # CRITICAL: all 7 russified issue types present with values in tolerance
            for exp_type, exp_vals in EXPECTED_ISSUE_TYPES.items():
                vals = type_rows.get(exp_type)
                if vals is None:
                    add(f"Issue '{exp_type}' present", False, critical=True)
                    continue
                add(f"Issue '{exp_type}' Total_Tickets",
                    num_close(vals[1], exp_vals["total"], 50), critical=True)
                add(f"Issue '{exp_type}' Avg_Resolution_Hours",
                    num_close(vals[3], exp_vals["avg_res_hrs"], 1.0), critical=False)
                add(f"Issue '{exp_type}' Avg_Satisfaction",
                    num_close(vals[4], exp_vals["avg_sat"], 0.1), critical=False)

            # CRITICAL: the LONGEST-resolution issue type (sorted desc -> first data row)
            #           must be 'Ошибка' (~16.0h)
            top_ok = False
            if data_rows:
                first = issue_grid[sorted_rows[1]]
                fcols = sorted(first.keys())
                if len(fcols) >= 5:
                    top_type = str(first[fcols[0]] or "")
                    top_ok = (str_contains(top_type, LONGEST_ISSUE_TYPE)
                              and num_close(first[fcols[3]], 16.0, 1.0))
            add("By Issue Type top row is longest-resolution 'Ошибка' (~16.0h)",
                top_ok, critical=True)

    # 3. "By Priority" sheet
    if ss_id is not None:
        cur.execute("SELECT id FROM gsheet.sheets WHERE spreadsheet_id = %s AND title ILIKE '%%Priority%%'", (ss_id,))
        sheet_rows = cur.fetchall()
        add("'By Priority' sheet exists", bool(sheet_rows), critical=False)
        if sheet_rows:
            sheet_id = sheet_rows[0][0]
            cur.execute("""
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (ss_id, sheet_id))
            grid = {}
            for row_idx, col_idx, value in cur.fetchall():
                grid.setdefault(row_idx, {})[col_idx] = value

            sorted_rows = sorted(grid.keys())
            data_rows = sorted_rows[1:] if len(sorted_rows) > 1 else []

            pri_rows = {}
            for row_idx in data_rows:
                row = grid[row_idx]
                cols = sorted(row.keys())
                if len(cols) < 4:
                    continue
                priority = str(row[cols[0]] or "")
                for exp_pri in EXPECTED_PRIORITIES:
                    if str_contains(priority, exp_pri):
                        pri_rows[exp_pri] = [row[c] for c in cols]

            # CRITICAL: all three priorities present with counts + response in tolerance
            for exp_pri, exp_vals in EXPECTED_PRIORITIES.items():
                vals = pri_rows.get(exp_pri)
                if vals is None:
                    add(f"Priority '{exp_pri}' present", False, critical=True)
                    continue
                add(f"Priority '{exp_pri}' Ticket_Count",
                    num_close(vals[1], exp_vals["count"], 50), critical=True)
                add(f"Priority '{exp_pri}' Avg_Response_Hours",
                    num_close(vals[2], exp_vals["avg_response"], 1.0), critical=True)
                add(f"Priority '{exp_pri}' SLA_Compliance_Pct",
                    num_close(vals[3], exp_vals["sla_pct"], 2.0), critical=False)

    # 4. Email (CRITICAL): to support-lead@company.com, subject, body mentions 'Ошибка'
    cur.execute("""
        SELECT subject, to_addr, body_text FROM email.messages
    """)
    email_rows = cur.fetchall()
    email_ok = False
    for subject, to_addr, body in email_rows:
        subj = str(subject or "").lower()
        to = str(to_addr or "").lower()
        bod = str(body or "").lower()
        if ("support-lead@company.com" in to
                and "support resolution performance report" in subj
                and LONGEST_ISSUE_TYPE.lower() in bod):
            email_ok = True
            break
    add("Email to support-lead w/ subject + body mentions 'Ошибка'", email_ok, critical=True)

    cur.close()
    conn.close()

    # 5. XLSX content (structural, non-critical): file exists with >=2 rows per sheet
    agent_ws = args.agent_workspace or os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    xlsx_path = os.path.join(agent_ws, "Support_Resolution_Analysis.xlsx")
    if not os.path.exists(xlsx_path):
        add("Support_Resolution_Analysis.xlsx exists", False, critical=False)
    else:
        add("Support_Resolution_Analysis.xlsx exists", True, critical=False)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(xlsx_path, data_only=True)
            ok = len(wb.worksheets) >= 1
            for ws in wb.worksheets:
                rows = list(ws.iter_rows(values_only=True))
                if len(rows) < 2:
                    ok = False
            wb.close()
            add("XLSX has >=2 rows per sheet", ok, critical=False)
        except Exception as e:
            add(f"XLSX readable ({e})", False, critical=False)

    # --- Scoring ---
    total = len(checks)
    passed = sum(1 for _, p, _ in checks if p)
    accuracy = (passed / total * 100.0) if total else 0.0
    critical_fail = [n for n, p, c in checks if c and not p]

    print("=== CHECKS ===")
    for name, p, c in checks:
        tag = "CRIT" if c else "    "
        print(f"  [{'PASS' if p else 'FAIL'}] {tag} {name}")
    print(f"\nAccuracy: {passed}/{total} = {accuracy:.1f}%")

    if critical_fail:
        print(f"=== RESULT: FAIL (critical checks failed: {critical_fail}) ===")
        sys.exit(1)
    if accuracy >= 70:
        print("=== RESULT: PASS ===")
        sys.exit(0)
    print(f"=== RESULT: FAIL (accuracy {accuracy:.1f}% < 70%) ===")
    sys.exit(1)


if __name__ == "__main__":
    main()
