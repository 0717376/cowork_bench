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

    # Clear gsheet schema
    cur.execute('DELETE FROM "gsheet"."cells"')
    cur.execute('DELETE FROM "gsheet"."sheets"')
    cur.execute('DELETE FROM "gsheet"."permissions"')
    cur.execute('DELETE FROM "gsheet"."spreadsheets"')
    cur.execute('DELETE FROM "gsheet"."folders"')

    # Clear email schema (sent messages)
    cur.execute('DELETE FROM "email"."sent_log"')
    cur.execute('DELETE FROM "email"."messages"')

    # Clear teamly user pages (seed pages have id <= 3) and ensure a space
    # exists for the loyalty program. Do NOT pre-seed the answer page.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("SELECT id FROM teamly.spaces WHERE key = 'TEAM'")
            if cur.fetchone() is None:
                cur.execute(
                    "INSERT INTO teamly.spaces (key, name, description) "
                    "VALUES ('TEAM', 'Команда', 'Рабочее пространство команды.')"
                )
    except Exception as e:
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")

    conn.commit()
    cur.close()
    conn.close()
    print("Data cleared for schemas: gsheet, email, teamly")


if __name__ == "__main__":
    main()
