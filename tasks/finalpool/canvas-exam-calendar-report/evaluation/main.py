"""
Скрипт оценки для задачи canvas-exam-calendar-report.

Проверки:
1. Excel-файл (exam_review_plan.xlsx) — корректные данные курсов из Canvas
2. События Google-календаря — учебные сессии созданы правильно
3. Письмо со сводкой по всем экзаменам

КРИТИЧЕСКИЕ ЧЕКИ (CRITICAL_CHECKS): семантические проверки сути задачи. Любой
их провал => вся задача FAIL (sys.exit(1)) ещё до порога по accuracy, независимо
от общего процента. Структурные проверки (наличие листа, ровно/не-менее событий)
остаются НЕ критическими.

Порог: accuracy >= 70% И отсутствие критических провалов => PASS.

Usage:
    python evaluation/main.py \
        --agent_workspace /path/to/workspace \
        --groundtruth_workspace /path/to/groundtruth \
        --res_log_file /path/to/result.json \
        --launch_time "2026-03-06 10:00:00"
"""

import argparse
import json
import sys

from .check_local import check_local, check_exam_dates, check_instructors
from .check_gcal import check_gcal
from .check_email import check_email


PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []

# Семантические критические чеки (по строке name, как в record()).
CRITICAL_CHECKS = {
    "xlsx: даты экзаменов и производные учебные сессии (date-math)",
    "xlsx: TBD/N-A ветвление для курсов без срока сдачи",
    "xlsx: имя/email преподавателя из Canvas (первый по алфавиту)",
    "gcal: 8 учебных сессий с верными названиями и датами",
    "email: тело письма содержит счётчики (4 с датой, 2 TBD) и сводку",
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAILED
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILED.append(name)


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    """Запустить все проверки."""

    # 1. Excel — полное сравнение с эталоном (структурный/широкий чек).
    print("\n=== Проверка Excel-вывода ===")
    local_pass, local_err = check_local(agent_workspace, groundtruth_workspace)
    record("Excel exam_review_plan.xlsx (полное сравнение с эталоном)", local_pass, local_err or "")

    # 1a/1b/1c. Семантические критические чеки по xlsx.
    dates_pass, dates_err = check_exam_dates(agent_workspace)
    record("xlsx: даты экзаменов и производные учебные сессии (date-math)", dates_pass, dates_err or "")

    tbd_pass = True
    tbd_err = ""
    # TBD/N-A ветвление: AAA-2013J и BBB-2013J => Exam_Date=TBD, обе сессии N/A.
    # check_exam_dates уже включает их ожидаемые значения; выделяем как отдельный
    # критический семантический чек на саму логику ветвления.
    if not dates_pass and ("AAA-2013J" in (dates_err or "") or "BBB-2013J" in (dates_err or "")):
        tbd_pass = False
        tbd_err = dates_err
    record("xlsx: TBD/N-A ветвление для курсов без срока сдачи", tbd_pass, tbd_err or "")

    instr_pass, instr_err = check_instructors(agent_workspace)
    record("xlsx: имя/email преподавателя из Canvas (первый по алфавиту)", instr_pass, instr_err or "")

    # 2. Google-календарь — ровно ожидаемые 8 сессий с верными датами/названиями.
    print("\n=== Проверка Google-календаря ===")
    gcal_pass, gcal_err = check_gcal()
    record("gcal: 8 учебных сессий с верными названиями и датами", gcal_pass, gcal_err or "")

    # 3. Письмо — сводка со счётчиками и кодами курсов.
    print("\n=== Проверка письма ===")
    email_pass, email_err = check_email()
    record("email: тело письма содержит счётчики (4 с датой, 2 TBD) и сводку", email_pass, email_err or "")

    total = PASS_COUNT + FAIL_COUNT
    pct = 100.0 * PASS_COUNT / total if total else 0.0
    critical_ok = not CRITICAL_FAILED
    all_passed = critical_ok and pct >= 70.0

    print(f"\n=== ИТОГ ===")
    print(f"  Пройдено: {PASS_COUNT}/{total} ({pct:.1f}%)")
    if CRITICAL_FAILED:
        print(f"  КРИТИЧЕСКИЙ ПРОВАЛ: {CRITICAL_FAILED}")
    print(f"  Результат: {'PASS' if all_passed else 'FAIL'}")

    if res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "accuracy": pct,
            "critical_failed": CRITICAL_FAILED,
            "success": all_passed,
            "details": {
                "excel_full": local_pass,
                "excel_dates": dates_pass,
                "excel_tbd": tbd_pass,
                "excel_instructors": instr_pass,
                "gcal": gcal_pass,
                "email": email_pass,
            },
        }
        with open(res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    return all_passed, f"Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}, Accuracy: {pct:.1f}%"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    success, message = run_evaluation(
        args.agent_workspace,
        args.groundtruth_workspace,
        args.launch_time,
        args.res_log_file,
    )
    print(message)

    if CRITICAL_FAILED:
        print(f"CRITICAL FAIL: {CRITICAL_FAILED}")
        sys.exit(1)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
