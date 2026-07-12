"""Preprocess for yf-financial-health-excel-notion (russified: moex + teamly).

Idempotently clears leftover Teamly pages (keeps seeded pages id<=3), ensures an
empty FINANCE space for the agent to write the dashboard into, and clears email
data. Does NOT pre-seed the answer page or the Excel file.
"""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        # Teamly: drop user-created pages (seed pages have id <= 3); ensure an
        # empty FINANCE space exists for the dashboard. Idempotent.
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('FINANCE', 'Финансовая аналитика',
                        'Отчёты и дашборды по финансовому здоровью компаний MOEX.')
                ON CONFLICT (key) DO NOTHING
            """)

        # Email: clear leftovers so the eval sees only this run's message.
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared teamly (id>3) and email data; ensured FINANCE space.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
