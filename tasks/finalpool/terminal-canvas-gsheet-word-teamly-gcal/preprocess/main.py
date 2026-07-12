"""Preprocess for terminal-canvas-gsheet-word-teamly-gcal.
Clears gsheet, teamly tracker pages, gcal. Injects RU noise. Canvas is read-only global seed."""
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Clear gsheet
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.permissions")
        cur.execute("DELETE FROM gsheet.spreadsheets")
        conn.commit()
        print("[preprocess] Cleared gsheet data.")

        # Clear teamly tracker pages left over from previous runs (idempotency).
        # We do NOT pre-create the tracker — the agent must build it in Teamly.
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                DELETE FROM teamly.pages
                 WHERE title ILIKE '%%advising%%'
                    OR title ILIKE '%%tracker%%'
                    OR title ILIKE '%%сопровожд%%'
                    OR title ILIKE '%%консультац%%'
            """)
            # Inject one RU noise space (must NOT be confused with the tracker).
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('ARCHIVE', 'Архив протоколов',
                        'Архив протоколов заседаний и служебных записок кафедры.')
                ON CONFLICT (key) DO NOTHING
            """)
            # Provide a generic space the agent can place the tracker into.
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('ADVISING', 'Академическое сопровождение',
                        'Материалы по академическому сопровождению студентов программы Global Governance.')
                ON CONFLICT (key) DO NOTHING
            """)
            conn.commit()
            print("[preprocess] Cleared teamly tracker pages; ensured spaces 'ARCHIVE' (noise) and 'ADVISING'.")
        else:
            print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql.")

        # Clear gcal
        cur.execute("DELETE FROM gcal.events")
        conn.commit()
        print("[preprocess] Cleared gcal events.")

        # Inject RU noise gcal events (should NOT match advising checks)
        noise_events = [
            ("Совещание кафедры", "2026-03-16 08:00:00", "2026-03-16 09:00:00", "Еженедельное совещание кафедры"),
            ("Заседание учебной комиссии", "2026-03-18 16:00:00", "2026-03-18 17:00:00", "Квартальный пересмотр учебного плана"),
        ]
        for summary, start, end, desc in noise_events:
            cur.execute("""
                INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status)
                VALUES (%s, %s, %s, %s, %s, 'confirmed')
            """, (str(uuid.uuid4()), summary, desc, start, end))

        # Inject RU noise gsheet (should NOT match the analytics check)
        noise_ss_id = str(uuid.uuid4())
        cur.execute("""
            INSERT INTO gsheet.spreadsheets (id, title, created_at, updated_at)
            VALUES (%s, %s, NOW(), NOW())
        """, (noise_ss_id, "Бюджет 2025"))

        conn.commit()
        print("[preprocess] Injected RU noise data.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Clean up agent workspace (do not leak groundtruth deliverables)
    if args.agent_workspace:
        for pattern in ["Academic_Advising_Report.xlsx", "Advising_Recommendations.docx", "advising_analyzer.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
