"""Preprocess for canvas-quiz-remediation-forms-gcal.

Idempotently clears forms (PG-backed RU-fork, schema gform.*), gcal, and email
data. Canvas (course 16 quizzes/submissions) is seeded globally and read live,
so it is NOT touched here. No answers are pre-seeded.
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
        cur.execute("DELETE FROM gform.responses")
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared gform, gcal, and email data.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
