"""
Preprocess for kulinar-weekly-plan-gsheet-gcal task (russified -> kulinar).
- Clears gsheet, gcal, and email schemas so agent starts fresh (idempotent).
- kulinar is a standalone MCP server (50 RU recipes seeded in JSON); no DB
  setup or answer pre-seeding is performed here.
"""
import os
import argparse
import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_schemas(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.spreadsheets")
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.messages")
    conn.commit()
    print("[preprocess] Cleared gsheet, gcal, and email schemas")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        clear_schemas(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
