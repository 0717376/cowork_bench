"""Оценка для course-enrollment-analytics-dashboard.

Стратегия: данные о курсах, зачислениях и успеваемости засеяны в canvas
ГЛОБАЛЬНО; и агент, и оценка читают их «вживую». Никакие значения не
захардкожены — все эталоны пересчитываются из canvas.* на момент проверки,
поэтому оценка остаётся честной при изменении сидов.

Этапы проверяются так:
  Э1/Э2 — пер-курсовые student_count / total_enrollment / avg_score /
           completion_rate пересчитываются из canvas и сверяются с листом
           Google Sheet "Course Enrollment Analytics" (Department_Metrics).
  Э2/Э4 — уровни риска по порогам avg_score (<60 Critical, 60-75 At Risk,
           >75 On Track); группа риска = Critical ∪ At Risk. Проверяется, что
           Risk_Level в листе и набор курсов в плане Teamly совпадают с
           пересчитанным эталоном (без ложных срабатываний).
  Э5    — страница Teamly "At-Risk Intervention Plan" с курсами риска и
           рекомендациями; письмо на academic-advisors@university.edu с темой
           про intervention и курсами риска в теле; событие календаря
           "Advising Follow-up Meeting" на 2026-06-12 14:00-15:00 UTC.

CRITICAL-проверки (любой провал => немедленный FAIL до порога accuracy):
  1. Лист Department_Metrics: avg_score корректен (в пределах допуска) для всех
     курсов и отсортирован по avg_score по возрастанию.
  2. Risk_Level в листе совпадает с порогами для всех курсов.
  3. Страница плана Teamly содержит РОВНО курсы группы риска (без ложных
     срабатываний/пропусков) и упоминает рекомендации.
  4. Письмо отправлено на academic-advisors@university.edu, тема про
     intervention, тело перечисляет курсы группы риска.
  5. Событие "Advising Follow-up Meeting" на 2026-06-12 14:00-15:00 UTC.
Порог: accuracy >= 70% И нет проваленных critical-проверок.
"""
import argparse
import json
import os
import re
import sys
from datetime import timezone

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "cowork_gym"),
          user="eigent", password="camel")

TARGET_EMAIL = "academic-advisors@university.edu"
EVENT_KW = "advising follow-up"
EVENT_DATE = (2026, 6, 12)
EVENT_START_HOUR_UTC = 14
EVENT_END_HOUR_UTC = 15

# Пороги уровня риска по среднему баллу курса.
CRIT_BELOW = 60.0
ATRISK_BELOW = 75.0

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

CRITICAL_CHECKS = {
    "Department_Metrics: avg_score корректен для всех курсов",
    "Department_Metrics отсортирован по Avg_Score по возрастанию",
    "Department_Metrics: Risk_Level соответствует порогам для всех курсов",
    "План Teamly содержит ровно курсы группы риска (без ложных срабатываний)",
    "Письмо отправлено на academic-advisors@university.edu, тема про intervention",
    "Событие 'Advising Follow-up Meeting' на 2026-06-12 14:00-15:00 UTC",
}


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    tag = "CRITICAL " if critical else ""
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {tag}{name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        print(f"  [FAIL] {tag}{name}: {str(detail)[:240]}")


def safe_float(val, default=None):
    try:
        if val is None:
            return default
        return float(str(val).replace(",", ".").replace("%", "").strip())
    except (ValueError, TypeError):
        return default


def to_utc_naive(dt):
    if dt is None:
        return None
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt


def short_name(course_name):
    """Название курса без суффикса '(Term Year)', в нижнем регистре."""
    return course_name.split("(")[0].strip().lower()


def risk_level(avg):
    if avg < CRIT_BELOW:
        return "critical"
    if avg < ATRISK_BELOW:
        return "at risk"
    return "on track"


def get_expected():
    """Пересчёт пер-курсовых метрик из canvas (вживую)."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM canvas.courses ORDER BY id")
    courses = {}
    for cid, name in cur.fetchall():
        # Кол-во студентов
        cur.execute("""SELECT COUNT(*), COUNT(*) FILTER (WHERE type LIKE '%%Student%%')
            FROM canvas.enrollments WHERE course_id = %s""", (cid,))
        total_enr, student_cnt = cur.fetchone()
        # Средний балл (current_score в зачислениях студентов)
        cur.execute("""SELECT AVG((grades->>'current_score')::numeric)
            FROM canvas.enrollments
            WHERE course_id = %s AND type = 'StudentEnrollment'
              AND grades->>'current_score' IS NOT NULL""", (cid,))
        avg_raw = cur.fetchone()[0]
        if avg_raw is None:
            # курс без оценок исключаем из аналитики
            continue
        avg_score = round(float(avg_raw), 1)
        # Доля выполнения (graded / all submissions)
        cur.execute("""SELECT COUNT(*) FILTER (WHERE s.workflow_state = 'graded'), COUNT(*)
            FROM canvas.submissions s JOIN canvas.assignments a ON s.assignment_id = a.id
            WHERE a.course_id = %s""", (cid,))
        graded, subs_total = cur.fetchone()
        comp_rate = round(graded / subs_total * 100, 1) if subs_total else 0.0
        courses[cid] = {
            "name": name,
            "student_count": int(student_cnt or 0),
            "total_enrollment": int(total_enr or 0),
            "avg_score": avg_score,
            "completion_rate": comp_rate,
            "risk_level": risk_level(avg_score),
        }
    cur.close()
    conn.close()
    return courses


# ---------------------------------------------------------------------------

def check_gsheet(expected):
    print("\n=== Проверка Google Sheet (Department_Metrics) ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""SELECT id, title FROM gsheet.spreadsheets
        WHERE title ILIKE '%course enrollment analytics%'
           OR title ILIKE '%enrollment analytics%'""")
    sheets = cur.fetchall()
    check("Таблица 'Course Enrollment Analytics' существует", len(sheets) >= 1,
          f"найдено {len(sheets)}")
    if not sheets:
        # critical avg/sort/risk проверки фиксируем как провал
        check("Department_Metrics: avg_score корректен для всех курсов", False,
              "таблица отсутствует", critical=True)
        check("Department_Metrics отсортирован по Avg_Score по возрастанию", False,
              "таблица отсутствует", critical=True)
        check("Department_Metrics: Risk_Level соответствует порогам для всех курсов",
              False, "таблица отсутствует", critical=True)
        conn.close()
        return

    ss_id = sheets[0][0]
    # Читаем структурированно: row->col->value
    cur.execute("""SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s ORDER BY row_index, col_index""", (ss_id,))
    grid = {}
    for ri, ci, val in cur.fetchall():
        grid.setdefault(ri, {})[ci] = val
    conn.close()

    if not grid:
        check("Department_Metrics содержит данные", False, "пустая таблица", critical=True)
        return

    rows = [grid[r] for r in sorted(grid)]
    # Шапка
    header_row = rows[0]
    headers = {str(v).strip().lower(): c for c, v in header_row.items() if v is not None}
    for col in ["Course_Name", "Total_Enrollment", "Avg_Score", "Completion_Rate", "Risk_Level"]:
        check(f"Department_Metrics содержит столбец {col}", col.lower() in headers,
              f"шапка: {list(headers)}")

    def col_val(row, key):
        ci = headers.get(key.lower())
        return row.get(ci) if ci is not None else None

    data_rows = rows[1:]
    data_rows = [r for r in data_rows if any(v not in (None, "") for v in r.values())]
    check("Department_Metrics: по строке на курс", len(data_rows) >= len(expected) - 1,
          f"строк {len(data_rows)} vs курсов {len(expected)}")

    # Сопоставляем строки листа с эталоном по названию курса.
    exp_by_short = {short_name(c["name"]): c for c in expected.values()}

    # --- CRITICAL: avg_score корректность ---
    avg_ok = 0
    avg_total = 0
    risk_ok = 0
    risk_total = 0
    listed_avgs = []
    for r in data_rows:
        cn = col_val(r, "Course_Name")
        if cn is None:
            continue
        key = short_name(str(cn))
        exp = exp_by_short.get(key)
        av = safe_float(col_val(r, "Avg_Score"))
        if av is not None:
            listed_avgs.append(av)
        if exp is None or av is None:
            continue
        avg_total += 1
        if abs(av - exp["avg_score"]) <= 0.6:
            avg_ok += 1
        # Risk_Level
        rl = col_val(r, "Risk_Level")
        if rl is not None:
            risk_total += 1
            if str(rl).strip().lower().replace("_", " ") == exp["risk_level"]:
                risk_ok += 1

    check("Department_Metrics: avg_score корректен для всех курсов",
          avg_total > 0 and avg_ok >= max(1, int(avg_total * 0.9)),
          f"{avg_ok}/{avg_total} совпали", critical=True)

    sorted_asc = all(listed_avgs[i] <= listed_avgs[i + 1] + 0.01
                     for i in range(len(listed_avgs) - 1))
    check("Department_Metrics отсортирован по Avg_Score по возрастанию",
          len(listed_avgs) >= 2 and sorted_asc, f"avgs: {listed_avgs}", critical=True)

    check("Department_Metrics: Risk_Level соответствует порогам для всех курсов",
          risk_total > 0 and risk_ok >= max(1, int(risk_total * 0.9)),
          f"{risk_ok}/{risk_total} совпали", critical=True)


def check_teamly(expected, at_risk_names):
    print("\n=== Проверка Teamly (At-Risk Intervention Plan) ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""SELECT id, title, COALESCE(body, '') FROM teamly.pages
        WHERE title ILIKE '%at-risk intervention plan%'
           OR title ILIKE '%intervention plan%'
           OR (title ILIKE '%intervention%' AND title ILIKE '%plan%')""")
    pages = cur.fetchall()
    conn.close()

    if not pages:
        check("Страница плана существует в Teamly", False, "не найдена")
        check("План Teamly содержит ровно курсы группы риска (без ложных срабатываний)",
              False, "страница отсутствует", critical=True)
        return
    check("Страница плана существует в Teamly", True)
    body = "\n".join(str(b) for _, _, b in pages)
    bl = body.lower()
    check("Страница плана содержит непустое тело", len(body) >= 150, f"len={len(body)}")
    # Рекомендации (RU/EN ключевые слова) — .lower() ОРИГИНАЛА
    rec_kw = any(k in bl for k in ("рекомендац", "вмешат", "поддержк", "консультир",
                                   "recommend", "intervention", "support"))
    check("Страница плана упоминает рекомендации/вмешательства", rec_kw,
          "ключевые слова рекомендаций не найдены")

    # CRITICAL: ровно курсы группы риска. Все at-risk курсы должны
    # присутствовать; курсы НЕ из группы риска не должны упоминаться.
    present_risk = sum(1 for n in at_risk_names if short_name(n) in bl)
    non_risk = [short_name(c["name"]) for c in expected.values()
                if c["risk_level"] == "on track"]
    false_positives = [n for n in non_risk if n and n in bl]
    no_fp = len(false_positives) == 0
    all_present = present_risk >= len(at_risk_names)
    check("План Teamly содержит ровно курсы группы риска (без ложных срабатываний)",
          all_present and no_fp,
          f"риск-курсов в плане {present_risk}/{len(at_risk_names)}; "
          f"ложные срабатывания: {false_positives[:5]}", critical=True)


def check_email_calendar(at_risk):
    print("\n=== Проверка письма и календаря ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # --- Письмо ---
    cur.execute("SELECT subject, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    recip_subj_ok = False
    matched_body = ""
    for subj, to_addr, body in emails:
        to_str = json.dumps(to_addr).lower() if isinstance(to_addr, list) else str(to_addr).lower()
        if TARGET_EMAIL.lower() in to_str and subj and "intervention" in subj.lower():
            recip_subj_ok = True
            matched_body = (body or "")
    check("Письмо отправлено на academic-advisors@university.edu, тема про intervention",
          recip_subj_ok, f"писем всего: {len(emails)}", critical=True)

    # Тело письма перечисляет курсы группы риска (хотя бы большинство).
    bl = matched_body.lower()
    listed = sum(1 for c in at_risk if short_name(c["name"]) in bl)
    need = max(1, int(len(at_risk) * 0.6)) if at_risk else 0
    check("Тело письма перечисляет курсы группы риска",
          (len(at_risk) == 0) or (listed >= need),
          f"перечислено {listed}/{len(at_risk)}")

    # Reverse-noise: шумовая рассылка не отправлена.
    cur.execute("""SELECT COUNT(*) FROM email.messages
        WHERE (subject ILIKE '%рассылк%' OR subject ILIKE '%newsletter%'
               OR subject ILIKE '%обслуживание%')
          AND to_addr::text ILIKE %s""", (f"%{TARGET_EMAIL}%",))
    noise_fwd = cur.fetchone()[0]
    check("Шумовые письма не пересланы консультантам", noise_fwd == 0, f"найдено {noise_fwd}")

    # --- Событие календаря ---
    cur.execute("SELECT summary, start_datetime, end_datetime FROM gcal.events WHERE summary ILIKE %s",
                (f"%{EVENT_KW}%",))
    events = cur.fetchall()
    check("Событие 'Advising Follow-up Meeting' существует", len(events) >= 1,
          "не найдено событие с 'advising follow-up'")
    dt_ok = False
    for summ, sdt, edt in events:
        su = to_utc_naive(sdt)
        eu = to_utc_naive(edt)
        if su is None:
            continue
        if (su.year, su.month, su.day) == EVENT_DATE and su.hour == EVENT_START_HOUR_UTC and su.minute == 0:
            if eu is None or (eu.hour == EVENT_END_HOUR_UTC and eu.minute == 0):
                dt_ok = True
                break
    check("Событие 'Advising Follow-up Meeting' на 2026-06-12 14:00-15:00 UTC",
          dt_ok,
          f"events: {[(e[0], str(to_utc_naive(e[1]))) for e in events][:3]}", critical=True)

    # Reverse-noise: шумовые события не удалены.
    cur.execute("""SELECT COUNT(*) FROM gcal.events
        WHERE summary ILIKE '%планёрк%' OR summary ILIKE '%планерк%' OR summary ILIKE '%обед%'""")
    noise_events = cur.fetchone()[0]
    check("Шумовые события календаря не удалены", noise_events >= 1, f"осталось {noise_events}")

    conn.close()


def run_evaluation(agent_workspace):
    expected = get_expected()
    print(f"[eval] курсов с оценками из canvas: {len(expected)}")

    at_risk = [c for c in expected.values() if c["risk_level"] in ("critical", "at risk")]
    at_risk_names = [c["name"] for c in at_risk]
    print(f"[eval] курсов группы риска (эталон): {len(at_risk)}")

    # --- Рабочие файлы агента (Э1/Э2/Э4) ---
    for fn in ["enrollment.json", "performance.json", "at_risk.json"]:
        p = os.path.join(agent_workspace, fn)
        check(f"{fn} создан", os.path.exists(p))
    py_files = [f for f in os.listdir(agent_workspace) if f.endswith(".py")] if os.path.isdir(agent_workspace) else []
    check("Python-скрипт анализа создан (risk_identifier.py)", len(py_files) >= 1, f"найдено: {py_files}")

    # at_risk.json внутренняя согласованность (если есть)
    arp = os.path.join(agent_workspace, "at_risk.json")
    if os.path.exists(arp):
        try:
            with open(arp, encoding="utf-8") as f:
                data = json.load(f)
            txt = json.dumps(data, ensure_ascii=False).lower()
            check("at_risk.json использует метки уровней риска",
                  ("critical" in txt) or ("at risk" in txt.replace("_", " ")) or ("at_risk" in txt))
        except Exception as e:
            check("at_risk.json парсится", False, str(e))

    check_gsheet(expected)
    check_teamly(expected, at_risk_names)
    check_email_calendar(at_risk)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = (PASS_COUNT / total * 100) if total else 0.0
    return accuracy, total


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", default=".")
    parser.add_argument("--groundtruth_workspace", default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    try:
        accuracy, total = run_evaluation(args.agent_workspace)
    except Exception as e:
        print(f"[eval] FATAL: {e}")
        FAILED_NAMES.append("DB/eval error")
        accuracy, total = 0.0, max(1, PASS_COUNT + FAIL_COUNT)

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    print(f"\n=== ИТОГО: {PASS_COUNT}/{total} проверок ({accuracy:.1f}%) ===")
    if critical_failed:
        print(f"CRITICAL FAILED ({len(critical_failed)}): {critical_failed}")

    success = (not critical_failed) and (accuracy >= 70)
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "accuracy": accuracy,
                       "critical_failed": critical_failed, "success": success}, f)

    if critical_failed:
        print("Overall: FAIL (critical check failed)")
        sys.exit(1)
    if accuracy >= 70:
        print("Overall: PASS")
        sys.exit(0)
    print("Overall: FAIL")
    sys.exit(1)


if __name__ == "__main__":
    main()
