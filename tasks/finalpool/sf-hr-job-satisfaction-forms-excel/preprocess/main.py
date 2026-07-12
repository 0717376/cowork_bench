"""Preprocess: clear forms, email, and teamly leftovers for sf-hr-job-satisfaction-gform-excel."""
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

    # Clear forms (gform schema) data
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")

    # Clear email data
    for t in ["email.attachments", "email.sent_log", "email.messages"]:
        cur.execute(f"DELETE FROM {t}")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        conn.rollback()

    # Clear teamly leftovers: keep the 3 seeded sample pages (id <= 3),
    # remove anything the agent (or a previous run) may have created.
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")

    conn.commit()
    cur.close()
    conn.close()
    print("Cleared forms, email, and teamly leftover data")


if __name__ == "__main__":
    main()
