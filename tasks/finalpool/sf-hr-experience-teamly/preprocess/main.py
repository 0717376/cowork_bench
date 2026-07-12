"""Preprocess: clear writable schema data for clean state (Teamly + email)."""
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

    # Clear user-created Teamly pages (seed pages have id <= 3); ensure an HR space.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('HR', 'HR-аналитика',
                        'Отчёты и дашборды по персоналу: стаж, текучесть, эффективность.')
                ON CONFLICT (key) DO NOTHING
            """)
    except Exception as e:
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")

    # Clear email tables (FK order)
    for table in ["attachments", "sent_log", "drafts", "messages", "folders", "account_config"]:
        cur.execute(f'DELETE FROM email."{table}"')

    conn.commit()
    cur.close()
    conn.close()
    print("Data cleared for schemas: teamly (user pages), email")


if __name__ == "__main__":
    main()
