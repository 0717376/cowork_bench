"""
Скрипт оценки для задачи canvas-enrollment-gsheet.

Проверки:
1. Таблица Google Sheet "Fall 2014 Enrollment Tracker" существует с корректными
   данными о зачислении (читаются вживую из Canvas).
2. Письмо отправлено на planning@university.edu о недобирающих (under-enrolled) курсах.

Модель оценки: CRITICAL_CHECKS — любой провал критической проверки => общий FAIL
независимо от accuracy. Иначе PASS требует accuracy >= 70%.
Ожидаемые значения берутся динамически из canvas.courses/enrollments, чтобы
оставаться честными к «живым» данным; жёстко зашитых значений нет.
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

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []

# Критические проверки: провал любой => общий FAIL независимо от accuracy.
CRITICAL_CHECKS = {
    "CRITICAL: GSheet содержит код топ-курса (макс. зачисление)",
    "CRITICAL: GSheet содержит количество для топ-курса",
    "CRITICAL: Письмо отправлено на planning@university.edu",
    "CRITICAL: Все недобирающие курсы (code) перечислены в теле письма",
    "CRITICAL: Все количества недобирающих курсов есть в теле письма",
    "CRITICAL: Ни один курс с зачислением >= 1000 не попал в письмо",
}


def load_expected():
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT c.name, c.course_code, COUNT(e.id) as enrollment_count
        FROM canvas.courses c
        LEFT JOIN canvas.enrollments e ON c.id = e.course_id
        WHERE c.name LIKE '%%Fall 2014%%'
        GROUP BY c.name, c.course_code
        ORDER BY COUNT(e.id) DESC
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return [{"name": r[0], "code": r[1], "count": int(r[2])} for r in rows]


def check(name, condition, detail="", critical=False):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        detail_str = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def check_gsheet(expected):
    print("\n=== Проверка Google Sheet ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE LOWER(title) LIKE '%%fall 2014%%' AND LOWER(title) LIKE '%%enrollment%%'
    """)
    sheets = cur.fetchall()
    # Структурная проверка (не критическая).
    check("Таблица Google Sheet 'Fall 2014 Enrollment Tracker' существует",
          len(sheets) >= 1,
          f"Найдено {len(sheets)} подходящих таблиц")

    all_values = ""
    if sheets:
        ss_id = sheets[0][0]
        cur.execute("""
            SELECT c.value FROM gsheet.cells c
            JOIN gsheet.sheets s ON c.spreadsheet_id = s.spreadsheet_id AND c.sheet_id = s.id
            WHERE c.spreadsheet_id = %s
        """, (ss_id,))
        cells = cur.fetchall()
        all_values = " ".join(str(c[0]) for c in cells if c[0])

    av_lower = all_values.lower()

    # Структурные (не критические): название листа и заголовки столбцов.
    check("Лист 'Enrollment' присутствует (по заголовку/ячейкам)",
          "enrollment" in av_lower or len(sheets) >= 1,
          "")
    check("Заголовок столбца 'Course Name' присутствует",
          "course name" in av_lower, "")
    check("Заголовок столбца 'Course Code' присутствует",
          "course code" in av_lower, "")
    check("Заголовок столбца 'Enrollment Count' присутствует",
          "enrollment count" in av_lower, "")

    # КРИТИЧЕСКИЕ: топ-курс (макс. зачисление) — доказательство реальных данных.
    if expected:
        top = expected[0]  # ORDER BY count DESC
        check("CRITICAL: GSheet содержит код топ-курса (макс. зачисление)",
              top["code"] in all_values,
              f"Код '{top['code']}' не найден", critical=True)
        check("CRITICAL: GSheet содержит количество для топ-курса",
              str(top["count"]) in all_values,
              f"Количество {top['count']} ({top['code']}) не найдено", critical=True)

    # Остальные курсы в таблице — не критические структурные проверки.
    for course in expected[1:]:
        check(f"GSheet содержит код курса '{course['code']}'",
              course["code"] in all_values,
              "Не найдено в ячейках")
        check(f"GSheet содержит количество {course['count']} для {course['code']}",
              str(course["count"]) in all_values,
              "Значение не найдено")

    cur.close()
    conn.close()


def check_email(expected):
    print("\n=== Проверка письма ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_emails = cur.fetchall()

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

    target_email = "planning@university.edu"
    found = None
    for subj, from_addr, to_addr, body in all_emails:
        recipients = parse_recipients(to_addr)
        if target_email in recipients:
            found = (subj, from_addr, to_addr, body)
            break

    # КРИТИЧЕСКАЯ: письмо ушло точному получателю.
    check("CRITICAL: Письмо отправлено на planning@university.edu",
          found is not None,
          f"Всего писем: {len(all_emails)}", critical=True)

    if found:
        subj, from_addr, to_addr, body = found
        subj_lower = (subj or "").lower()
        body_lower = (body or "").lower()

        # Структурные (не критические): маркеры темы. EN-литерал сохранён в task.md,
        # но допускаем и RU-варианты темы.
        check("Тема письма содержит маркер недобора (EN/RU)",
              ("under" in subj_lower and "enroll" in subj_lower)
              or "недобор" in subj_lower or "недонабор" in subj_lower,
              f"Тема: {(subj or '')[:100]}")

        check("Тема письма содержит 'Fall 2014' (EN/RU)",
              "fall 2014" in subj_lower or "осень 2014" in subj_lower,
              f"Тема: {(subj or '')[:100]}")

        under_enrolled = [c for c in expected if c["count"] < 1000]
        over_enrolled = [c for c in expected if c["count"] >= 1000]

        # КРИТИЧЕСКАЯ: все недобирающие курсы (по EN-коду или имени) есть в теле.
        all_codes_ok = True
        all_counts_ok = True
        for course in under_enrolled:
            name_parts = course["name"].lower().split("(")[0].strip()
            code_ok = course["code"].lower() in body_lower or name_parts in body_lower
            count_ok = str(course["count"]) in (body or "")
            if not code_ok:
                all_codes_ok = False
            if not count_ok:
                all_counts_ok = False
        check("CRITICAL: Все недобирающие курсы (code) перечислены в теле письма",
              all_codes_ok,
              f"Не все из {len(under_enrolled)} курсов найдены", critical=True)
        check("CRITICAL: Все количества недобирающих курсов есть в теле письма",
              all_counts_ok,
              f"Не все количества найдены", critical=True)

        # КРИТИЧЕСКАЯ: ни один курс с >=1000 не должен быть упомянут как недобирающий.
        # Проверяем, что код перенаборных курсов не фигурирует в теле письма.
        leaked = []
        for course in over_enrolled:
            if course["code"].lower() in body_lower:
                leaked.append(course["code"])
        check("CRITICAL: Ни один курс с зачислением >= 1000 не попал в письмо",
              len(leaked) == 0,
              f"Просочились: {leaked}", critical=True)
    else:
        # Помечаем зависимые критические проверки как проваленные.
        check("CRITICAL: Все недобирающие курсы (code) перечислены в теле письма",
              False, "Письмо не найдено", critical=True)
        check("CRITICAL: Все количества недобирающих курсов есть в теле письма",
              False, "Письмо не найдено", critical=True)
        check("CRITICAL: Ни один курс с зачислением >= 1000 не попал в письмо",
              False, "Письмо не найдено", critical=True)

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected = load_expected()

    check_gsheet(expected)
    check_email(expected)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total else 0
    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]

    print(f"\n=== Итог: {PASS_COUNT}/{total} проверок пройдено ({accuracy:.1f}%) ===")
    if critical_failed:
        print(f"КРИТИЧЕСКИЕ ПРОВАЛЫ: {len(critical_failed)}")
        for n in critical_failed:
            print(f"  - {n}")

    if args.res_log_file:
        try:
            with open(args.res_log_file, "w") as f:
                json.dump({
                    "total_passed": PASS_COUNT, "total_checks": total,
                    "accuracy": accuracy, "critical_failed": critical_failed,
                }, f, indent=2)
        except Exception:
            pass

    success = (not critical_failed) and accuracy >= 70
    print(f"  Результат: {'PASS' if success else 'FAIL'}")
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
