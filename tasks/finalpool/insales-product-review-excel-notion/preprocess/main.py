"""
Preprocess script for insales-product-review-excel-notion task (InSales + Teamly).
Clears Teamly user pages and email data; ensures the PRODUCTS space exists.
The InSales store data (wc.* schema) is read-only and russified centrally.
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        # Clear Teamly user-created data idempotently; seed pages have id <= 3.
        # Ensure the PRODUCTS space exists (empty — do NOT pre-seed the answer).
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.page_labels WHERE page_id > 3")
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('PRODUCTS', 'Продукты',
                        'Аналитика по товарам и клиентским отзывам интернет-магазина.')
                ON CONFLICT (key) DO NOTHING
                """
            )
        print("[preprocess] Cleared Teamly user pages; ensured PRODUCTS space.")

        # Clear email data
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

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
