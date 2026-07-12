"""Preprocess for canvas-ta-workload-excel-email.

- Ensures a Teamly space exists for the agent to create the "TA Staffing
  Overview" page in, and clears any prior staffing pages (idempotency).
- Clears email data.
- Canvas is the read-only source of truth and is left untouched.

We intentionally do NOT pre-create the staffing page, the Excel file, nor the
email — the agent must produce them itself so the evaluation actually tests the
agent.
"""
import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def setup_teamly(cur):
    """Ensure a department Teamly space exists and clear prior staffing pages."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return
    # Dedicated space for the agent to drop the staffing overview page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('DEPT', 'Кафедра',
                'Кадровые сводки и учёт нагрузки ассистентов преподавателей (TA) по курсам кафедры.')
        ON CONFLICT (key) DO NOTHING
    """)
    # Idempotency: drop any staffing pages left from previous runs (EN title is
    # preserved per task.md; cover a couple of lenient variants too).
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%ta staffing%'
            OR title ILIKE '%staffing overview%'
            OR title ILIKE '%кадров%ассистент%'
    """)
    print("[preprocess] Teamly ready: 'DEPT' space ensured, prior staffing pages cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        setup_teamly(cur)

        # Seed leak: foreign course 9991 (CHR-RU-101) breaks the 22-course scope.
        cur.execute("DELETE FROM canvas.courses WHERE id = 9991")

        # Clear email data
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

        conn.commit()
        print("[preprocess] Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
