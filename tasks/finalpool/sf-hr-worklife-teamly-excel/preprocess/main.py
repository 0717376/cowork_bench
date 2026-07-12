"""
Preprocess script for sf-hr-worklife task (ClickHouse + Teamly).
ClickHouse HR DWH is read-only. Clear agent-created Teamly pages and email data.
"""
import argparse
import glob
import os

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
        print("[preprocess] Clearing agent-created Teamly pages...")
        # Keep seed pages (id <= 3); remove anything the agent may have created.
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")

        print("[preprocess] Clearing email data...")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages WHERE folder_id != 0")
        cur.execute("DELETE FROM email.drafts")

        conn.commit()
        print("[preprocess] DB cleanup done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["WL_Balance_Report.xlsx"]:
            for f in glob.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
