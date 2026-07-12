"""Preprocess for terminal-canvas-teamly-study-gcal-excel.

- Clears Teamly study-planner pages (idempotency) and ensures a dedicated
  Teamly space the agent can drop the "Student Study Planner" page into.
- Injects a little unrelated Teamly noise so the eval must find the right page.
- Clears gcal events and injects two RU noise events on Thu Mar 5 / Fri Mar 6
  (outside the Mar 9-13 study week, no weekend conflict).
- Canvas is read-only (live LMS data); we do not touch it.

We deliberately do NOT pre-create the study-planner page, the calendar study
sessions, the Excel report nor the script — the agent must produce them itself.
"""
import argparse
import glob as globmod
import os
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}


def setup_teamly(cur):
    """Ensure the advisor space exists and clear prior study-planner pages."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; "
              "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return

    # Idempotency: drop any study-planner pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%student study planner%'
            OR title ILIKE '%study planner%'
            OR title ILIKE '%учебн%план%'
    """)

    # Dedicated space for the agent to drop the study-planner page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('ADVISOR', 'Учебный консультант',
                'Планы учебной нагрузки и расписания студентов, '
                'подготовленные академическим консультантом университета.')
        ON CONFLICT (key) DO NOTHING
    """)

    # A bit of unrelated noise so the eval must locate the right page.
    cur.execute("SELECT id FROM teamly.spaces WHERE key='ADVISOR'")
    space_id = cur.fetchone()[0]
    cur.execute("""
        INSERT INTO teamly.pages (space_id, title, body, author)
        SELECT %s, 'Заметки с планёрки',
               E'# Заметки с планёрки\n\nОбсудили закупку канцелярии и расписание дежурств.',
               'office'
        WHERE NOT EXISTS (
            SELECT 1 FROM teamly.pages WHERE title='Заметки с планёрки'
        )
    """, (space_id,))
    print("[preprocess] Teamly ready: 'ADVISOR' space ensured, "
          "prior study-planner pages cleared, noise page inserted.")


def setup_gcal(cur):
    """Clear gcal events and inject two RU noise events outside the study week."""
    cur.execute("DELETE FROM gcal.events")
    cur.execute(
        "INSERT INTO gcal.events (id, summary, description, start_datetime, "
        "end_datetime, status) VALUES (%s, %s, %s, %s, %s, 'confirmed')",
        (str(uuid.uuid4()), "Планёрка кафедры", "Еженедельная планёрка",
         "2026-03-05 09:00:00", "2026-03-05 09:30:00"))
    cur.execute(
        "INSERT INTO gcal.events (id, summary, description, start_datetime, "
        "end_datetime, status) VALUES (%s, %s, %s, %s, %s, 'confirmed')",
        (str(uuid.uuid4()), "Обеденная йога", "Оздоровительная активность",
         "2026-03-06 12:00:00", "2026-03-06 12:45:00"))
    print("[preprocess] Cleared gcal and injected RU noise events.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        setup_teamly(cur)
        setup_gcal(cur)
        conn.commit()
        print("[preprocess] Committed teamly + gcal setup.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["Study_Plan_Report.xlsx", "study_planner.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
