"""
Evaluation script for canvas-assignment-calendar task.

Проверки:
1. Текстовый файл (assignment_schedule.txt) совпадает с эталоном
2. События Google Calendar созданы для каждого задания
3. Отправлено сводное письмо с расписанием

CRITICAL-проверки (любой провал => немедленный FAIL):
  C1. assignment_schedule.txt воспроизводит все 52 эталонные строки (100%).
  C2. Файл отсортирован: по коду курса по алфавиту, затем по дате по возрастанию.
  C3. В gcal.events ровно 52 события; spot-check AAA-2014J / TMA 1758 на 2014-10-19.
  C4. Письмо: тема содержит англ. подстроку 'Fall 2014 Assignment Schedule',
      from=coordinator@, to=students@, тело содержит все 7 кодов курсов и '52'.

Тема/тело письма и описания событий могут содержать русскоязычный текст вокруг
обязательных английских идентификаторов — проверки ищут подстроки, не точное равенство.
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

EXPECTED_TOTAL = 52
COURSE_CODES = [
    "AAA-2014J", "BBB-2014J", "CCC-2014J", "DDD-2014J",
    "EEE-2014J", "FFF-2014J", "GGG-2014J",
]

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILURES = []


def record(name, passed, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS]{' [CRIT]' if critical else ''} {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL]{' [CRIT]' if critical else ''} {name}{msg}")
        if critical:
            CRITICAL_FAILURES.append(name)


# ============================================================================
# Check 1: assignment_schedule.txt
# ============================================================================

def check_text_file(agent_workspace, groundtruth_workspace):
    print("\n=== Checking assignment_schedule.txt ===")

    agent_file = os.path.join(agent_workspace, "assignment_schedule.txt")
    gt_file = os.path.join(groundtruth_workspace, "assignment_schedule.txt")

    if not os.path.isfile(agent_file):
        record("Text file exists", False, f"Not found: {agent_file}")
        record("C1: all 52 lines reproduced", False, "no file", critical=True)
        record("C2: sorted correctly", False, "no file", critical=True)
        return False
    record("Text file exists", True)

    if not os.path.isfile(gt_file):
        record("Groundtruth text file exists", False, f"Not found: {gt_file}")
        return False

    with open(agent_file) as f:
        agent_lines = [l.strip() for l in f.readlines() if l.strip()]
    with open(gt_file) as f:
        gt_lines = [l.strip() for l in f.readlines() if l.strip()]

    record("Line count matches", len(agent_lines) == len(gt_lines),
           f"Expected {len(gt_lines)}, got {len(agent_lines)}")

    # Точное совпадение каждой эталонной строки (по трём полям).
    def parts(line):
        return [p.strip().lower() for p in line.split(" - ")]

    matched = 0
    for gt_line in gt_lines:
        gp = parts(gt_line)
        for agent_line in agent_lines:
            ap = parts(agent_line)
            if len(ap) >= 3 and len(gp) >= 3:
                if ap[0] == gp[0] and ap[1] == gp[1] and ap[2] == gp[2]:
                    matched += 1
                    break

    full_match = (matched == len(gt_lines))
    record(f"Assignment lines matched ({matched}/{len(gt_lines)})",
           full_match, f"Matched {matched} of {len(gt_lines)}")

    # CRITICAL C1: все 52 строки воспроизведены (100%).
    record(f"C1: all {EXPECTED_TOTAL} GT lines reproduced",
           full_match and len(gt_lines) == EXPECTED_TOTAL,
           f"matched={matched}, gt={len(gt_lines)}", critical=True)

    # CRITICAL C2: сортировка по коду курса (алфавит), затем по дате (возр.).
    sort_ok = True
    sort_detail = ""
    parsed = []
    for line in agent_lines:
        p = [x.strip() for x in line.split(" - ")]
        if len(p) >= 3:
            parsed.append((p[0], p[-1], line))  # (code, date, line)
        else:
            sort_ok = False
            sort_detail = f"bad line: {line}"
    if sort_ok:
        # Группировка по коду; коды должны идти неубывающим алфавитом,
        # внутри кода даты неубывающие.
        prev_code = None
        prev_date_in_code = None
        seen_codes_order = []
        for code, date, line in parsed:
            if code != prev_code:
                seen_codes_order.append(code)
                prev_date_in_code = None
            if prev_code is not None and code < prev_code and code not in seen_codes_order[:-1]:
                pass
            prev_code = code
            if prev_date_in_code is not None and date < prev_date_in_code:
                sort_ok = False
                sort_detail = f"dates out of order in {code}: {prev_date_in_code} > {date}"
                break
            prev_date_in_code = date
        # Проверка алфавитного порядка кодов (без повторных групп).
        if sort_ok:
            dedup = []
            for c in seen_codes_order:
                if c not in dedup:
                    dedup.append(c)
            if dedup != sorted(dedup) or len(dedup) != len(seen_codes_order):
                sort_ok = False
                sort_detail = f"course codes not alphabetical/grouped: {seen_codes_order}"
    record("C2: sorted by course code then due date", sort_ok,
           sort_detail, critical=True)

    return full_match


# ============================================================================
# Check 2: Google Calendar events
# ============================================================================

def check_gcal():
    print("\n=== Checking Google Calendar ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT summary, description, start_datetime, end_datetime
        FROM gcal.events
        ORDER BY summary
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_gcal] Found {len(events)} calendar events.")

    record("At least 50 calendar events created", len(events) >= 50,
           f"Found {len(events)}")

    # CRITICAL C3a: ровно 52 события (по одному на задание со сроком).
    record(f"C3: gcal has exactly {EXPECTED_TOTAL} events",
           len(events) == EXPECTED_TOTAL, f"Found {len(events)}",
           critical=True)

    # Spot check: по событию на каждый код курса.
    all_ok = True
    for code in COURSE_CODES:
        code_events = [e for e in events if code.lower() in (e[0] or "").lower()]
        record(f"gcal: events exist for {code}",
               len(code_events) > 0, f"No events found for {code}")
        if not code_events:
            all_ok = False

    # CRITICAL C3b: точечная проверка AAA-2014J / TMA 1758 -> 2014-10-19.
    spot_ok = False
    spot_detail = "AAA-2014J - TMA 1758 not found"
    for summary, desc, start_dt, end_dt in events:
        s = (summary or "").lower()
        if "aaa-2014j" in s and "tma 1758" in s:
            start_str = str(start_dt or "")
            if "2014-10-19" in start_str:
                spot_ok = True
                spot_detail = ""
            else:
                spot_detail = f"start_datetime={start_str} (expected 2014-10-19)"
            break
    record("C3: spot-check AAA-2014J/TMA 1758 due 2014-10-19",
           spot_ok, spot_detail, critical=True)

    return all_ok


# ============================================================================
# Check 3: Email sent
# ============================================================================

def check_emails():
    print("\n=== Checking Emails ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_emails = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_emails] Found {len(all_emails)} total emails.")
    record("At least 1 email sent", len(all_emails) >= 1, f"Found {len(all_emails)}")

    all_ok = True
    found_schedule = False
    target = None

    # Письмо-расписание: тема содержит англ. подстроку (допускается RU-текст вокруг).
    for subject, from_addr, to_addr, body_text in all_emails:
        subject_lower = (subject or "").lower()
        if ("fall 2014 assignment schedule" in subject_lower
                or "assignment schedule" in subject_lower
                or "fall 2014" in subject_lower):
            found_schedule = True
            target = (subject, from_addr, to_addr, body_text)

            from_str = str(from_addr or "").lower()
            record("email: from coordinator",
                   "coordinator@openuniversity.ac.uk" in from_str,
                   f"From: {from_addr}")

            to_str = str(to_addr or "").lower()
            record("email: to students",
                   "students@openuniversity.ac.uk" in to_str,
                   f"To: {to_addr}")

            body_lower = (body_text or "").lower()
            record("email: body mentions AAA-2014J",
                   "aaa-2014j" in body_lower, "Missing AAA-2014J in body")
            record("email: body mentions GGG-2014J",
                   "ggg-2014j" in body_lower, "Missing GGG-2014J in body")
            record("email: body mentions total count",
                   "52" in (body_text or ""), "Missing total count of 52")
            break

    record("email: schedule email found", found_schedule,
           "No email with assignment schedule subject")
    if not found_schedule:
        all_ok = False

    # CRITICAL C4: адресация + все 7 кодов курсов + общий счётчик 52.
    if target is None:
        record("C4: email correctly addressed & consolidated", False,
               "no schedule email", critical=True)
        all_ok = False
    else:
        subject, from_addr, to_addr, body_text = target
        subj_l = (subject or "").lower()
        body_l = (body_text or "").lower()
        crit_ok = (
            "fall 2014 assignment schedule" in subj_l
            and "coordinator@openuniversity.ac.uk" in str(from_addr or "").lower()
            and "students@openuniversity.ac.uk" in str(to_addr or "").lower()
            and all(c.lower() in body_l for c in COURSE_CODES)
            and "52" in (body_text or "")
        )
        missing = [c for c in COURSE_CODES if c.lower() not in body_l]
        record("C4: email correctly addressed & consolidated", crit_ok,
               f"subj_ok={'fall 2014 assignment schedule' in subj_l}, "
               f"missing_codes={missing}, has52={'52' in (body_text or '')}",
               critical=True)

    return all_ok


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

    text_ok = check_text_file(args.agent_workspace, args.groundtruth_workspace)
    gcal_ok = check_gcal()
    email_ok = check_emails()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Accuracy: {accuracy:.1f}%")
    print(f"  Critical failures: {CRITICAL_FAILURES}")

    # CRITICAL gate: любой провал критической проверки => немедленный FAIL.
    if CRITICAL_FAILURES:
        print(f"  Overall: FAIL (critical: {CRITICAL_FAILURES})")
        if args.res_log_file:
            with open(args.res_log_file, "w") as f:
                json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT,
                           "accuracy": accuracy, "success": False,
                           "critical_failures": CRITICAL_FAILURES}, f, indent=2)
        sys.exit(1)

    all_passed = accuracy >= 70 and text_ok and gcal_ok and email_ok

    print(f"  Overall: {'PASS' if all_passed else 'FAIL'}")

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT,
                       "accuracy": accuracy, "success": all_passed,
                       "critical_failures": CRITICAL_FAILURES}, f, indent=2)

    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
