"""Оценка задачи terminal-canvas-pdf-gsheet-email-word.

Архетип: keep_foreign + RU-сценарий. Инструменты (canvas/google_sheet/word/emails/
terminal/pdf-tools/filesystem) остаются «иностранными», источник (canvas.*) читается
вживую. Прозу агент может писать по-русски ИЛИ по-английски, поэтому семантические
проверки опираются на ЦИФРЫ, а не на английские названия предметов.

Эталон пересчитывается из canvas.* вживую (БЕЗ хардкода волатильных данных):
    Enrollment       = canvas.courses.total_students
    Assignment_Count = COUNT(canvas.assignments по course_id)
    Quiz_Count       = COUNT(canvas.quizzes по course_id)
    Subject          = название курса с удалённой группой ' (...)' в конце
    Avg_Assignments  = среднее число заданий по курсам одного предмета
    Workload_Rating  = Heavy (>=12), Moderate (8..11), Light (<8)

CRITICAL_CHECKS: любой провал критической проверки => sys.exit(1) до порога accuracy.
Порог: accuracy >= 70 И нет критических провалов => PASS.
"""
import argparse
import json
import os
import re
import sys

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAIL = False


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAIL
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS]{' [CRIT]' if critical else ''} {name}")
    else:
        FAIL_COUNT += 1
        if critical:
            CRITICAL_FAIL = True
        detail_str = str(detail)[:300] if detail else ""
        print(f"  [FAIL]{' [CRIT]' if critical else ''} {name}: {detail_str}")


def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)


def subject_of(name):
    """Предмет = название курса с удалённой одиночной группой ' (...)' в конце."""
    return re.sub(r" \(.*\)$", "", str(name)).strip()


def rating_of(avg):
    if avg >= 12:
        return "heavy"
    if avg >= 8:
        return "moderate"
    return "light"


# ---------------------------------------------------------------------------
# Эталон из canvas (НЕ хардкод)
# ---------------------------------------------------------------------------

def compute_expected():
    """Возвращает (per_course, subjects) пересчётом из canvas.* вживую.

    per_course: list of dict(name, subject, assignments, quizzes, enrollment)
    subjects: dict subject_lower -> dict(num_courses, avg_assignments,
              total_enrollment, rating, subject)
    """
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT c.id, c.name, COALESCE(c.total_students, 0),
               (SELECT COUNT(*) FROM canvas.assignments a WHERE a.course_id = c.id),
               (SELECT COUNT(*) FROM canvas.quizzes q WHERE q.course_id = c.id)
        FROM canvas.courses c
        ORDER BY c.name
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    per_course = []
    groups = {}
    for cid, name, enr, asg, qz in rows:
        subj = subject_of(name)
        per_course.append({
            "name": str(name), "subject": subj,
            "assignments": int(asg), "quizzes": int(qz),
            "enrollment": int(enr),
        })
        g = groups.setdefault(subj.lower(), {"subject": subj, "asg": [], "enr": []})
        g["asg"].append(int(asg))
        g["enr"].append(int(enr))

    subjects = {}
    for key, g in groups.items():
        n = len(g["asg"])
        avg = sum(g["asg"]) / n if n else 0.0
        subjects[key] = {
            "subject": g["subject"],
            "num_courses": n,
            "avg_assignments": avg,
            "total_enrollment": sum(g["enr"]),
            "rating": rating_of(avg),
        }
    return per_course, subjects


# ---------------------------------------------------------------------------
# Чтение листов из gsheet
# ---------------------------------------------------------------------------

def load_sheet_grid(cur, sheet_id):
    """Возвращает (headers_lower, list_of_row_dicts) для листа."""
    cur.execute(
        "SELECT row_index, col_index, value FROM gsheet.cells WHERE sheet_id = %s "
        "ORDER BY row_index, col_index", (sheet_id,))
    cells = cur.fetchall()
    if not cells:
        return [], []
    grid = {}
    maxc = 0
    for r, c, v in cells:
        grid[(r, c)] = v
        maxc = max(maxc, c)
    headers = [str(grid.get((0, c), "")).strip() for c in range(maxc + 1)]
    headers_lower = [h.lower() for h in headers]
    max_row = max(r for r, c in grid.keys())
    rows = []
    for r in range(1, max_row + 1):
        row = {}
        has_val = False
        for c in range(maxc + 1):
            v = grid.get((r, c), "")
            if v is not None and str(v).strip():
                has_val = True
            row[headers_lower[c] if c < len(headers_lower) else str(c)] = v
        if has_val:
            rows.append(row)
    return headers_lower, rows


def col_get(row, *needles):
    """Достаёт значение из row по подстроке в имени столбца."""
    for k, v in row.items():
        for n in needles:
            if n in k:
                return v
    return None


def to_num(v):
    try:
        return float(str(v).replace(",", ".").strip())
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Проверки
# ---------------------------------------------------------------------------

def check_gsheet(expected_courses, expected_subjects):
    print("\n=== Проверка Google Sheets ===")
    try:
        conn = get_conn()
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE %s",
                    ("%faculty%workload%tracker%",))
        rows = cur.fetchall()
        if not rows:
            cur.execute("SELECT id, title FROM gsheet.spreadsheets")
            all_ss = cur.fetchall()
            check("Faculty_Workload_Tracker spreadsheet exists", False,
                  f"Found: {[r[1] for r in all_ss]}")
            cur.close()
            conn.close()
            return
        check("Faculty_Workload_Tracker spreadsheet exists", True)
        ss_id = rows[0][0]

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
        sheets = cur.fetchall()
        sheet_titles = [s[1].lower() for s in sheets]

        workload_sheet_id = None
        summary_sheet_id = None
        for sid, title in sheets:
            t = title.lower()
            if workload_sheet_id is None and ("workload" in t or "course" in t):
                workload_sheet_id = sid
            if summary_sheet_id is None and ("summary" in t or "subject" in t):
                summary_sheet_id = sid

        check("Course_Workload sheet exists", workload_sheet_id is not None, f"Sheets: {sheet_titles}")
        check("Subject_Summary sheet exists", summary_sheet_id is not None, f"Sheets: {sheet_titles}")

        # --- Course_Workload ---
        cw_rows = []
        if workload_sheet_id:
            headers, cw_rows = load_sheet_grid(cur, workload_sheet_id)
            has_name = any("name" in h or "course" in h for h in headers)
            has_assign = any("assign" in h for h in headers)
            check("Course_Workload has course name column", has_name, f"Headers: {headers}")
            check("Course_Workload has assignment column", has_assign, f"Headers: {headers}")
            check(f"Course_Workload has >= {len(expected_courses)} rows",
                  len(cw_rows) >= len(expected_courses),
                  f"Found {len(cw_rows)} data rows, expected {len(expected_courses)}")

            # CRITICAL: пер-курсовые числа совпадают с canvas для конкретных курсов.
            # Сопоставляем по названию курса; проверяем Assignment_Count и Enrollment.
            by_name = {}
            for row in cw_rows:
                nm = col_get(row, "name", "course")
                if nm is not None:
                    by_name[str(nm).strip().lower()] = row
            matched = 0
            mismatches = []
            for ec in expected_courses:
                row = by_name.get(ec["name"].strip().lower())
                if not row:
                    continue
                matched += 1
                asg = to_num(col_get(row, "assign"))
                enr = to_num(col_get(row, "enroll"))
                if asg is not None and abs(asg - ec["assignments"]) > 0.5:
                    mismatches.append(f"{ec['name']}: assign {asg}!={ec['assignments']}")
                if enr is not None and abs(enr - ec["enrollment"]) > 0.5:
                    mismatches.append(f"{ec['name']}: enroll {enr}!={ec['enrollment']}")
            # Требуем, чтобы большинство курсов сопоставилось и значения совпали.
            check("Course_Workload per-course Assignment_Count/Enrollment match canvas",
                  matched >= max(1, int(0.7 * len(expected_courses))) and not mismatches,
                  f"matched={matched}/{len(expected_courses)}, mismatches={mismatches[:5]}",
                  critical=True)

        # --- Subject_Summary ---
        if summary_sheet_id:
            headers, ss_rows = load_sheet_grid(cur, summary_sheet_id)
            check(f"Subject_Summary has >= {len(expected_subjects)} rows",
                  len(ss_rows) >= len(expected_subjects),
                  f"Found {len(ss_rows)} data rows, expected {len(expected_subjects)}")

            all_values = " ".join(str(v).lower() for row in ss_rows for v in row.values())
            check("Subject_Summary mentions Heavy rating",
                  "heavy" in all_values,
                  f"Values sample: {all_values[:120]}")

            # CRITICAL: число строк предметов == число уникальных предметов (группировка верна).
            check("Subject_Summary row count equals distinct subject count",
                  len(ss_rows) == len(expected_subjects),
                  f"rows={len(ss_rows)} vs subjects={len(expected_subjects)}",
                  critical=True)

            # CRITICAL: Workload_Rating и Avg_Assignments корректны для каждого предмета.
            by_subj = {}
            for row in ss_rows:
                sv = col_get(row, "subject")
                if sv is not None:
                    by_subj[str(sv).strip().lower()] = row
            rating_ok = True
            avg_ok = True
            problems = []
            for key, exp in expected_subjects.items():
                row = by_subj.get(key)
                if not row:
                    rating_ok = False
                    avg_ok = False
                    problems.append(f"missing subject {exp['subject']}")
                    continue
                got_rating = str(col_get(row, "rating", "workload") or "").strip().lower()
                if got_rating and exp["rating"] not in got_rating:
                    rating_ok = False
                    problems.append(f"{exp['subject']}: rating '{got_rating}'!={exp['rating']}")
                got_avg = to_num(col_get(row, "avg"))
                if got_avg is not None and abs(got_avg - exp["avg_assignments"]) > 0.6:
                    avg_ok = False
                    problems.append(
                        f"{exp['subject']}: avg {got_avg}!={round(exp['avg_assignments'],2)}")
            check("Subject_Summary Workload_Rating correct per 12/8 thresholds",
                  rating_ok, f"problems={problems[:6]}", critical=True)
            check("Subject_Summary Avg_Assignments equals mean of subject's courses",
                  avg_ok, f"problems={problems[:6]}", critical=True)

        cur.close()
        conn.close()
    except Exception as e:
        check("Google Sheets accessible", False, str(e))


def check_word(agent_workspace, expected_courses, expected_subjects):
    print("\n=== Проверка Word ===")
    word_path = os.path.join(agent_workspace, "Faculty_Workload_Report.docx")
    check("Faculty_Workload_Report.docx exists", os.path.exists(word_path))
    if not os.path.exists(word_path):
        return
    from docx import Document
    doc = Document(word_path)
    text = " ".join(p.text for p in doc.paragraphs)
    low = text.lower()
    check("Word has substantial content", len(text) > 300, f"length: {len(text)}")
    # RU/EN-толерантно: ключевые слова на двух языках.
    check("Word mentions workload",
          "workload" in low or "нагруз" in low)
    check("Word mentions recommendations",
          "recommend" in low or "рекоменд" in low)

    digits = set(re.findall(r"\d+", text))

    total_courses = len(expected_courses)
    total_enrollment = sum(c["enrollment"] for c in expected_courses)
    # CRITICAL: документ ссылается на корректные общие числа (цифры, не слова-предметы).
    # Допускаем совпадение либо total_courses, либо total_enrollment (оба вычислены из источника).
    has_courses_num = str(total_courses) in digits
    has_enroll_num = str(total_enrollment) in digits
    check("Word references computed totals (course count or total enrollment)",
          has_courses_num or has_enroll_num,
          f"need courses={total_courses} or enrollment={total_enrollment}; digits={sorted(digits)[:20]}",
          critical=True)

    # CRITICAL: упомянут хотя бы один предмет с рейтингом Heavy (по avg_assignments).
    heavy = [s for s in expected_subjects.values() if s["rating"] == "heavy"]
    if heavy:
        # Толерантно: имя предмета (как в canvas) ИЛИ его avg-число присутствует в тексте.
        found = False
        for s in heavy:
            if s["subject"].lower() in low:
                found = True
                break
            if str(int(round(s["avg_assignments"]))) in digits:
                found = True
                break
        check("Word references a Heavy subject (name or its avg figure)",
              found,
              f"heavy subjects={[s['subject'] for s in heavy]}",
              critical=True)
    else:
        # Нет тяжёлых предметов — критическую проверку считаем пройденной.
        check("Word references a Heavy subject (name or its avg figure)", True,
              "no Heavy subjects in source", critical=True)


def check_emails(expected_subjects):
    print("\n=== Проверка писем ===")
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT subject, to_addr, body_text FROM email.messages WHERE subject ILIKE %s",
                    ('%faculty%workload%analysis%complete%',))
        emails = cur.fetchall()
        check("Faculty Workload Analysis Complete email sent", len(emails) >= 1, f"found {len(emails)}")
        if emails:
            check("Email to department-chairs",
                  "department-chairs" in str(emails[0][1]).lower() or "chair" in str(emails[0][1]).lower(),
                  f"to: {emails[0][1]}")

        cur.execute("SELECT subject, to_addr, body_text FROM email.messages WHERE subject ILIKE %s",
                    ('%workload%standards%compliance%',))
        compliance = cur.fetchall()
        check("Workload Standards Compliance Report email sent", len(compliance) >= 1,
              f"found {len(compliance)}")
        if compliance:
            check("Compliance email to academic-affairs",
                  "academic-affairs" in str(compliance[0][1]).lower() or "affairs" in str(compliance[0][1]).lower(),
                  f"to: {compliance[0][1]}")

        # CRITICAL: оба письма существуют с верными темами и адресатами; письмо о
        # соответствии нормативам называет хотя бы один предмет, превышающий норматив
        # (согласованно с рейтингом Heavy). Толерантно RU/EN: имя предмета ИЛИ его avg-число.
        both_ok = len(emails) >= 1 and len(compliance) >= 1
        dept_ok = bool(emails) and (
            "department-chairs@university.edu" in str(emails[0][1]).lower()
            or "department-chairs" in str(emails[0][1]).lower())
        aa_ok = bool(compliance) and (
            "academic-affairs@university.edu" in str(compliance[0][1]).lower()
            or "academic-affairs" in str(compliance[0][1]).lower())

        heavy = [s for s in expected_subjects.values() if s["rating"] == "heavy"]
        compliance_mentions_heavy = True
        if compliance and heavy:
            body = (str(compliance[0][2]) or "").lower()
            digits = set(re.findall(r"\d+", body))
            compliance_mentions_heavy = any(
                s["subject"].lower() in body
                or str(int(round(s["avg_assignments"]))) in digits
                for s in heavy)

        check("Both emails: correct English subjects, recipients, and compliance names an exceeding subject",
              both_ok and dept_ok and aa_ok and compliance_mentions_heavy,
              f"both={both_ok} dept={dept_ok} aa={aa_ok} heavy_named={compliance_mentions_heavy}",
              critical=True)

        cur.close()
        conn.close()
    except Exception as e:
        check("Email checks", False, str(e))


def check_script(agent_workspace):
    print("\n=== Проверка скрипта ===")
    check("workload_analyzer.py exists",
          os.path.exists(os.path.join(agent_workspace, "workload_analyzer.py")))


def check_reverse_validation():
    print("\n=== Обратная валидация ===")
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("""
            SELECT to_addr FROM email.messages
            WHERE subject ILIKE '%%faculty%%workload%%'
               OR subject ILIKE '%%workload%%standards%%'
        """)
        emails = cur.fetchall()
        noise_recipients = ["all-staff@university.edu", "students@university.edu", "facilities@university.edu"]
        bad = None
        for email_row in emails:
            to_str = str(email_row[0]).lower()
            for noise in noise_recipients:
                if noise in to_str:
                    bad = noise
                    break
            if bad:
                break
        check("No workload emails sent to wrong recipients", bad is None,
              f"Sent to noise recipient: {bad}")
        cur.close()
        conn.close()
    except Exception as e:
        check("Reverse validation", False, str(e))


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAIL
    PASS_COUNT = 0
    FAIL_COUNT = 0
    CRITICAL_FAIL = False

    try:
        expected_courses, expected_subjects = compute_expected()
        print(f"[eval] canvas: {len(expected_courses)} courses, "
              f"{len(expected_subjects)} subjects")
    except Exception as e:
        print(f"[eval] WARNING: could not compute expected from canvas: {e}")
        expected_courses, expected_subjects = [], {}

    check_gsheet(expected_courses, expected_subjects)
    check_word(agent_workspace, expected_courses, expected_subjects)
    check_emails(expected_subjects)
    check_script(agent_workspace)
    check_reverse_validation()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (100.0 * PASS_COUNT / total) if total else 0.0
    if CRITICAL_FAIL:
        msg = (f"CRITICAL FAIL — провалена критическая проверка. "
               f"Passed {PASS_COUNT}/{total} ({accuracy:.0f}%)")
        return False, msg
    success = accuracy >= 70.0
    msg = f"Passed {PASS_COUNT}/{total} checks ({accuracy:.0f}%); threshold 70%"
    return success, msg


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    success, message = run_evaluation(
        args.agent_workspace, args.groundtruth_workspace,
        args.launch_time, args.res_log_file
    )
    print(message)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
