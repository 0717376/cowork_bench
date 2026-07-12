"""
Preprocess for canvas-course-notion-wiki task (Teamly knowledge base).

Canvas is read-only and is the source of truth (read live by the agent).
This script only clears leftover user-created Teamly pages for idempotency.
It does NOT pre-create any course pages (no answer pre-seeding).
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


def clear_teamly(cur):
    """Удаляет пользовательские страницы Teamly (сидовые id<=3 — это пример формата)."""
    print("[preprocess] Clearing user-created Teamly pages...")
    try:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    except Exception as e:
        print(f"[preprocess] teamly.pages cleanup skipped: {e}")
    print("[preprocess] Teamly user pages cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_teamly(cur)
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
