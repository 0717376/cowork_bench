"""Preprocess script for terminal-canvas-pdf-excel-teamly-gcal.

- Clears teamly pages and gcal events left from previous runs (idempotency).
- Ensures an 'ACCRED' Teamly space exists for the agent to create the
  "Accreditation Action Items" page in.
- Injects RU noise pages (unrelated tracker) + RU noise calendar events.

We intentionally do NOT pre-create the answer page nor the xlsx — the agent
must produce them itself so the evaluation actually tests the agent.
"""
import os
import argparse, json, os
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)


def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    # Clear teamly noise/answer pages from previous runs (idempotent).
    # Drop the noise project tracker space (and its pages via cascade) plus any
    # accreditation page the agent may have created on a previous run.
    cur.execute("DELETE FROM teamly.spaces WHERE key IN ('PROJTRACK', 'ACCRED')")
    cur.execute("DELETE FROM teamly.pages WHERE title ILIKE '%%accreditation action items%%'")
    # Clear gcal
    cur.execute("DELETE FROM gcal.events")
    conn.commit()
    cur.close()
    conn.close()
    print("[preprocess] Cleared teamly noise/answer pages and gcal schema")


def inject_noise_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")

    # ── Noise Teamly space + pages (unrelated project tracker) ───────────────
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('PROJTRACK', 'Вехи проекта', 'Отслеживание сроков по проектам')
        ON CONFLICT (key) DO NOTHING
        RETURNING id
    """)
    row = cur.fetchone()
    if row:
        proj_space_id = row[0]
    else:
        cur.execute("SELECT id FROM teamly.spaces WHERE key='PROJTRACK'")
        proj_space_id = cur.fetchone()[0]

    for title, status in [
        ("Обзор бюджета Q1", "Открыто"),
        ("Миграция серверов", "Закрыто"),
        ("Запуск маркетинговой кампании", "Открыто"),
    ]:
        cur.execute(
            "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
            (proj_space_id, title,
             f"# {title}\n\nСтатус: {status}\n\nНесвязанная с аккредитацией запись трекера проектов.",
             "admin"))

    # ── Empty space for the agent's accreditation page ───────────────────────
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('ACCRED', 'Аккредитация', 'Материалы самообследования к аккредитации программы.')
        ON CONFLICT (key) DO NOTHING
    """)

    # ── Noise calendar events ────────────────────────────────────────────────
    cur.execute("""INSERT INTO gcal.events (summary, description, start_datetime, end_datetime, status)
        VALUES (%s, %s, %s, %s, %s)""",
        ("Квартальный обзор бюджета",
         "Ежеквартальный обзор бюджета с финансовым отделом",
         launch_dt + timedelta(days=5),
         launch_dt + timedelta(days=5, hours=1),
         "confirmed"))

    cur.execute("""INSERT INTO gcal.events (summary, description, start_datetime, end_datetime, status)
        VALUES (%s, %s, %s, %s, %s)""",
        ("Планирование выезда преподавателей",
         "Планирование мероприятий предстоящего выезда преподавателей",
         launch_dt + timedelta(days=12),
         launch_dt + timedelta(days=12, hours=2),
         "confirmed"))

    # An existing accreditation-related event (noise but relevant topic)
    cur.execute("""INSERT INTO gcal.events (summary, description, start_datetime, end_datetime, status)
        VALUES (%s, %s, %s, %s, %s)""",
        ("Совещание комитета по аккредитации",
         "Первичное обсуждение сроков предстоящей аккредитации",
         launch_dt - timedelta(days=10),
         launch_dt - timedelta(days=10) + timedelta(hours=1),
         "confirmed"))

    conn.commit()
    cur.close()
    conn.close()
    print("[preprocess] Noise data injected (teamly: 1 space + 3 pages, gcal: 3 events); 'ACCRED' space ready")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_noise_data(args.launch_time)


if __name__ == "__main__":
    main()
