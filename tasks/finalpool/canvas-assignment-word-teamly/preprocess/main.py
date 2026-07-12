"""
Preprocess for canvas-assignment-word-notion task.

Canvas is read-only (live MCP). Teamly knowledge base: ensure a target space
exists and clear any leftover "Assignment Overview" pages from prior runs.

We do NOT pre-create the "CCC-2014J Assignment Overview" page — that is the
agent's deliverable.
"""

import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def setup_teamly(conn):
    """Ensure a teamly space exists and clear leftover deliverable pages."""
    print("[preprocess] Setting up Teamly knowledge base...")
    with conn.cursor() as cur:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is None:
            print("[preprocess] WARNING: teamly schema not found; "
                  "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
            return
        # Dedicated space for the agent to drop the assignment overview into.
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('COURSES', 'Учебные курсы',
                    'Методички и обзоры заданий учебных курсов кафедры.')
            ON CONFLICT (key) DO NOTHING
        """)
        # Idempotency: drop any overview pages left from previous runs.
        cur.execute("""
            DELETE FROM teamly.pages
             WHERE title ILIKE '%CCC-2014J%'
                OR title ILIKE '%assignment overview%'
                OR title ILIKE '%обзор заданий%'
        """)
    conn.commit()
    print("[preprocess] Teamly ready: 'COURSES' space ensured, prior overview pages cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)

    try:
        setup_teamly(conn)
        conn.commit()
        print("[preprocess] Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
