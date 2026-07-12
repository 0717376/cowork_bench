"""Preprocess for canvas-submission-stats.

- Ensures a Teamly space exists for the agent to create the "Canvas Submission
  Analysis" page in, and clears any prior analysis pages (idempotency).
- Removes stray canvas courses/enrollments/submissions left over from other
  tasks (course_id > 22) so the canonical 22-course dataset stays the single
  source of truth that the static groundtruth was computed from.

We intentionally do NOT pre-create the analysis page nor the xlsx — the agent
must produce them itself so the evaluation actually tests the agent.
"""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def setup_teamly(cur):
    """Ensure the registrar Teamly space exists and clear prior analysis pages."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return
    # Dedicated space for the agent to drop the analysis page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('REGISTRAR', 'Учебный отдел',
                'Аналитика по сдаче работ и учёт курсов учебного отдела университета.')
        ON CONFLICT (key) DO NOTHING
    """)
    # Idempotency: drop any analysis pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%canvas submission analysis%'
            OR title ILIKE '%submission analysis%'
            OR title ILIKE '%анализ%сдач%'
    """)
    print("[preprocess] Teamly ready: 'REGISTRAR' space ensured, prior analysis pages cleared.")


def clean_stray_canvas(cur):
    """Remove canvas rows leaked from other tasks so the canonical 22-course
    dataset stays the single source of truth. The seeded dataset is ids 1..22;
    only rows with course_id > 22 are pruned. Each delete is guarded so a
    missing optional table does not abort preprocess."""
    for tbl, col in [
        ("canvas.submissions", "course_id"),
        ("canvas.enrollments", "course_id"),
        ("canvas.assignments", "course_id"),
        ("canvas.courses", "id"),
    ]:
        try:
            cur.execute(f"SELECT to_regclass('{tbl}')")
            if cur.fetchone()[0] is None:
                continue
            cur.execute(f"DELETE FROM {tbl} WHERE {col} > 22")
        except Exception as e:
            print(f"[preprocess] skip {tbl}: {e}")
    print("[preprocess] Cleaned stray canvas rows (id > 22).")


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
