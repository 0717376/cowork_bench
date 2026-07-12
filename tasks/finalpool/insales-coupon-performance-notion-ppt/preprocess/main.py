"""Preprocess: clear writable schema data for clean state."""
import os
import argparse
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Clear teamly pages created by the agent (keep seeded noise pages id <= 3)
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0]:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")

    # Clear email schema
    cur.execute('DELETE FROM "email"."sent_log"')
    cur.execute('DELETE FROM "email"."messages"')

    conn.commit()
    cur.close()
    conn.close()
    print("Data cleared for schemas: teamly, email")


if __name__ == "__main__":
    main()
