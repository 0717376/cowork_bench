"""Preprocess for canvas-enrollment-teamly.

- Ensures a Teamly space exists for the agent to create the enrollment tracker
  page in, and clears any prior tracker pages (idempotency).
- Clears email data.
- Removes stray canvas courses/enrollments left over from other tasks
  (course_id > 22) so the canonical 22-course dataset is the single source of
  truth that the evaluation recomputes from.

We intentionally do NOT pre-create the tracker page nor the email — the agent
must produce them itself so the evaluation actually tests the agent.
"""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def setup_teamly(cur):
    """Ensure the registrar Teamly space exists and clear prior tracker pages."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return
    # Dedicated space for the agent to drop the enrollment tracker page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('REGISTRAR', 'Учебный отдел',
                'Сводки по набору студентов и учёт курсов учебного отдела университета.')
        ON CONFLICT (key) DO NOTHING
    """)
    # Idempotency: drop any tracker pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%course enrollment tracker%'
            OR title ILIKE '%enrollment%'
            OR title ILIKE '%трекер%набор%'
    """)
    print("[preprocess] Teamly ready: 'REGISTRAR' space ensured, prior tracker pages cleared.")


def clean_stray_canvas(cur):
    """Remove canvas courses/enrollments leaked from other tasks so that the
    canonical 22-course dataset stays the single source of truth. We only touch
    rows with course_id > 22 (the seeded enrollment dataset is ids 1..22)."""
    cur.execute("DELETE FROM canvas.enrollments WHERE course_id > 22")
    cur.execute("DELETE FROM canvas.courses WHERE id > 22")
    print("[preprocess] Cleaned stray canvas courses/enrollments (id > 22).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        setup_teamly(cur)
        clean_stray_canvas(cur)
        # Clear emails
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared email data. Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
