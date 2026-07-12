"""Preprocess for sf-hr-mentorship-excel-gcal-email (ClickHouse fork).

Idempotently clears the writable schemas this task touches (email, gcal) and copies
the initial_workspace files (Russian Mentorship_Guidelines.pdf) into the agent
workspace. The russified sf_data HR data is seeded centrally
(db/zzz_clickhouse_after_init.sql) -- nothing is pre-seeded as an answer here.
"""
import argparse
import os
import shutil
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def clear_writable_schemas(conn):
    cur = conn.cursor()
    # Only clear tables that actually exist (guarded, idempotent).
    targets = [
        "email.sent_log",
        "email.messages",
        "gcal.events",
    ]
    for t in targets:
        cur.execute("SELECT to_regclass(%s)", (t,))
        if cur.fetchone()[0] is not None:
            cur.execute(f"DELETE FROM {t}")
    conn.commit()
    print("[preprocess] Writable schemas cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    try:
        clear_writable_schemas(conn)
    finally:
        conn.close()

    if args.agent_workspace and os.path.exists(args.agent_workspace):
        initial_ws = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "initial_workspace")
        if os.path.exists(initial_ws):
            for f in os.listdir(initial_ws):
                src = os.path.join(initial_ws, f)
                if os.path.isfile(src) and not f.endswith(".py"):
                    shutil.copy2(src, os.path.join(args.agent_workspace, f))
                    print(f"[preprocess] Copied {f} to agent workspace.")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
