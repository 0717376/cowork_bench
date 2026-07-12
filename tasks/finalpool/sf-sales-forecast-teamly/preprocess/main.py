"""
Preprocess script for sf-sales-forecast-notion task.

- Ensures a Teamly space exists for the agent to create the dashboard page in.
- Clears any dashboard pages left over from previous runs (idempotency).

We intentionally do NOT pre-create the dashboard page nor the xlsx — the agent
must produce them itself so the evaluation actually tests the agent.

ClickHouse (sf_data) data is read-only and seeded centrally.
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is None:
            print("[preprocess] WARNING: teamly schema not found; "
                  "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        else:
            # Dedicated space for the agent to drop the dashboard page into.
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('SALES', 'Продажи',
                        'Аналитические панели и отчёты отдела продаж.')
                ON CONFLICT (key) DO NOTHING
            """)
            # Idempotency: drop any dashboard pages left from previous runs.
            cur.execute("""
                DELETE FROM teamly.pages
                 WHERE title ILIKE '%sales performance dashboard%'
                    OR title ILIKE '%monthly revenue%'
                    OR title ILIKE '%regional performance%'
                    OR title ILIKE '%панель продаж%'
            """)
        conn.commit()
        print("[preprocess] Teamly ready: 'SALES' space ensured, prior dashboard pages cleared.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
