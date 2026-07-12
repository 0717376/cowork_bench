"""
Preprocess script for sf-sales-product-ranking-ppt task.

ClickHouse (sf_data schema) is read-only source data. This script only
clears Google Sheet state idempotently for a clean evaluation.
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


def clear_gsheet(cur):
    print("[preprocess] Clearing Google Sheet data...")
    for t in ["gsheet.cells", "gsheet.sheets", "gsheet.permissions", "gsheet.spreadsheets", "gsheet.folders"]:
        cur.execute(f"DELETE FROM {t}")
        print(f"[preprocess] Cleared {t}")
    print("[preprocess] Google Sheet data cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gsheet(cur)
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
