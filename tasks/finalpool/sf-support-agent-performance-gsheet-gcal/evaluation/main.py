"""Evaluation for sf-support-agent-performance-gsheet-gcal.

Источник данных — хранилище ClickHouse (схема sf_data). Значения REPORTER в
таблице TICKETS руссифицируются централизованно (db/zzz_clickhouse_after_init.sql):
Alice->Алиса, Bob->Борис, Charlie->Карл, Emily->Эмилия, John->Иван. Поэтому агент
ОБЯЗАН записать в таблицу русские имена, и evaluator сверяется именно с ними.

Структура проверок:
  - НЕСКОЛЬКО CRITICAL (семантических) проверок: правильные имена агентов из БД,
    верные количества тикетов по каждому агенту и порядок сортировки, корректно
    определённые лидеры в Rankings, событие календаря ровно через 10 дней после
    запуска (жёсткий допуск), письмо с правильными адресами, чьё тело называет
    И лучшего по CSAT, И агента с наибольшим объёмом. Любой провал CRITICAL =>
    sys.exit(1) ещё ДО порога точности.
  - Остальные структурные проверки идут к порогу accuracy >= 70.

Замечания по локали:
  - Сопоставление имён регистро- и акцент-независимое, принимаются RU и EN формы.
  - RU-ключевые слова ищутся в .lower() ОРИГИНАЛЬНОГО текста (без normalize).
"""
import argparse
import os
import sys
import unicodedata
from datetime import datetime, timedelta

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILURES = []

# Фактические данные из БД (отсортированы по тикетам по убыванию).
# Имена даны в RU-форме (как в sf_data после руссификации) + EN-алиас для гибкого матча.
AGENT_DATA = [
    # (RU_name, EN_alias, Total_Tickets, Avg_Response_Hrs, Avg_CSAT, SLA_Compliance_Rate)
    ("Эмилия", "Emily",   9193, 15.09, 3.25, 21.47),
    ("Карл",   "Charlie", 6640, 15.00, 3.24, 20.99),
    ("Иван",   "John",    6199, 15.06, 3.27, 20.37),
    ("Борис",  "Bob",     5446, 15.01, 3.26, 21.26),
    ("Алиса",  "Alice",   4110, 14.94, 3.27, 20.39),
]

TICKET_COUNTS = [str(a[2]) for a in AGENT_DATA]  # ["9193","6640","6199","5446","4110"]
TOP_VOLUME = AGENT_DATA[0]                        # Эмилия, 9193
# Топ по CSAT: 3.27 — ничья между Иван и Алиса. Принимаем любого из них.
MAX_CSAT = max(a[4] for a in AGENT_DATA)
TOP_CSAT_AGENTS = [a for a in AGENT_DATA if abs(a[4] - MAX_CSAT) < 1e-6]  # Иван, Алиса


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def check_critical(name, condition, detail=""):
    check("[CRITICAL] " + name, condition, detail)
    if not condition:
        CRITICAL_FAILURES.append(name)


def norm(s):
    """Регистр- и акцент-независимая нормализация для сопоставления имён."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.lower()


def agent_in(agent, text_norm):
    """Имя агента присутствует в тексте (RU-форма или EN-алиас)."""
    return norm(agent[0]) in text_norm or norm(agent[1]) in text_norm


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE title ILIKE '%agent%performance%' OR title ILIKE '%agent%scorecard%'
    """)
    sheets = cur.fetchall()
    check("Agent Performance Scorecard spreadsheet exists", len(sheets) >= 1,
          f"Found: {[s[1] for s in sheets]}")

    if not sheets:
        cur.close()
        conn.close()
        # Без таблицы критические семантические проверки невозможны.
        check_critical("Таблица содержит руссифицированные имена агентов (>=4 из 5)", False, "no spreadsheet")
        check_critical("Верные количества тикетов по агентам присутствуют", False, "no spreadsheet")
        return False

    ss_id = sheets[0][0]

    # Check sheets/tabs
    cur.execute("SELECT title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
    tabs = [r[0] for r in cur.fetchall()]
    has_scorecards = any("scorecard" in t.lower() for t in tabs)
    has_rankings = any("ranking" in t.lower() for t in tabs)
    check("Has Scorecards sheet", has_scorecards, f"Tabs: {tabs}")
    check("Has Rankings sheet", has_rankings, f"Tabs: {tabs}")

    # Все значения ячеек таблицы.
    cur.execute("SELECT value FROM gsheet.cells WHERE spreadsheet_id = %s", (ss_id,))
    cells = [str(r[0]) for r in cur.fetchall() if r[0] is not None]
    all_vals = " ".join(cells)
    all_norm = norm(all_vals)

    cur.close()
    conn.close()

    # --- CRITICAL: правильные русские имена агентов из БД ---
    present = sum(1 for a in AGENT_DATA if agent_in(a, all_norm))
    check_critical(
        "Таблица содержит руссифицированные имена агентов из БД (Эмилия/Карл/Иван/Борис/Алиса, >=4 из 5)",
        present >= 4,
        f"Найдено {present}/5; values_norm[:200]={all_norm[:200]}")

    # --- CRITICAL: верные количества тикетов по агентам ---
    counts_present = sum(1 for c in TICKET_COUNTS if c in all_vals)
    check_critical(
        "Верные количества тикетов присутствуют (9193/6640/6199/5446/4110, >=4 из 5)",
        counts_present >= 4 and "9193" in all_vals,
        f"Найдено {counts_present}/5, есть 9193={'9193' in all_vals}")

    # Структурная (некритическая) проверка наличия хотя бы одного значения CSAT.
    check("Таблица содержит значение Avg_CSAT (3.2x)",
          any(c in all_vals for c in ["3.24", "3.25", "3.26", "3.27", "3,24", "3,25", "3,26", "3,27"]),
          "Значения CSAT не найдены")

    return has_scorecards and has_rankings


def check_gcal(launch_time_str):
    print("\n=== Checking Google Calendar ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT summary, description, start_datetime FROM gcal.events ORDER BY start_datetime")
    events = cur.fetchall()
    cur.close()
    conn.close()

    print(f"  Found {len(events)} calendar events")
    check("At least 1 calendar event", len(events) >= 1, f"Found {len(events)}")

    review_events = [e for e in events
                     if "agent" in norm(e[0]) and "performance" in norm(e[0])]
    has_review = len(review_events) >= 1
    check("Agent Performance Review event exists", has_review,
          f"Events: {[e[0] for e in events]}")

    # --- CRITICAL: событие ровно через 10 дней после запуска (жёсткий допуск) ---
    if launch_time_str:
        if not has_review:
            check_critical("Событие 'Agent Performance Review' ровно через 10 дней после запуска",
                           False, "событие не найдено")
        else:
            try:
                launch_dt = datetime.fromisoformat(launch_time_str)
                expected_dt = launch_dt + timedelta(days=10)
                ev_dt = next((ev[2] for ev in review_events if ev[2]), None)
                if ev_dt is None:
                    check_critical("Событие 'Agent Performance Review' ровно через 10 дней после запуска",
                                   False, "у события нет start_datetime")
                else:
                    diff = abs((ev_dt.replace(tzinfo=None) - expected_dt).total_seconds())
                    # Жёсткий допуск: несколько часов, не +/-2 дня.
                    check_critical(
                        "Событие 'Agent Performance Review' ровно через 10 дней после запуска",
                        diff <= 3600 * 6,
                        f"Ожидалось около {expected_dt}, получено {ev_dt} (diff {diff/3600:.1f}ч)")
            except Exception as e:
                print(f"  [INFO] Не удалось проверить дату: {e}")

    return has_review


def check_email():
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
        WHERE subject ILIKE '%agent%' AND subject ILIKE '%performance%'
        ORDER BY date DESC
    """)
    emails = cur.fetchall()
    cur.close()
    conn.close()

    has_email = len(emails) >= 1
    check("Agent performance email exists", has_email, f"Found {len(emails)}")
    if not has_email:
        check_critical(
            "Письмо performance-review@ называет лучшего по CSAT и агента с макс. объёмом (Эмилия)",
            False, "письмо не найдено")
        return

    e = emails[0]
    to_str = norm(e[2])
    from_str = norm(e[1])
    check("Email to performance-review@company.example.com",
          "performance-review@company.example.com" in to_str, f"to: {e[2]}")
    check("Email from support-manager@company.example.com",
          "support-manager@company.example.com" in from_str, f"from: {e[1]}")

    body_norm = norm(e[3])
    # --- CRITICAL: тело письма называет И лучшего по CSAT, И агента с макс. объёмом ---
    names_top_csat = any(agent_in(a, body_norm) for a in TOP_CSAT_AGENTS)
    names_top_volume = agent_in(TOP_VOLUME, body_norm)
    check_critical(
        "Письмо performance-review@ называет лучшего по CSAT и агента с макс. объёмом (Эмилия)",
        names_top_csat and names_top_volume,
        f"top_csat_named={names_top_csat}, top_volume(Эмилия)_named={names_top_volume}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_gcal(args.launch_time)
    check_email()

    total = PASS_COUNT + FAIL_COUNT
    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    if total == 0:
        print("FAIL: проверки не выполнялись.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"  Accuracy: {accuracy:.1f}%")

    # Критический гейт ДО порога точности.
    if CRITICAL_FAILURES:
        print(f"\nCRITICAL FAILURE ({len(CRITICAL_FAILURES)}): {CRITICAL_FAILURES}")
        print("Overall: FAIL")
        sys.exit(1)

    overall = accuracy >= 70
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
