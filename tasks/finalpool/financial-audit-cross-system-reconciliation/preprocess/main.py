#!/usr/bin/env python3
"""Preprocess: prepare a clean state for the financial-audit reconciliation task.

The reconciliation source data lives in initial_workspace files (the harness copies
them into the agent workspace). Here we only clear the writable gcal/email schemas
so the calendar-event and email checks start from a clean slate. We do NOT pre-seed
any answer (no audit spreadsheet, no report, no email, no event).
"""
from argparse import ArgumentParser
from pathlib import Path
import os

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        Path(args.agent_workspace).mkdir(parents=True, exist_ok=True)

    # Idempotently clear writable schemas (calendar + email) for a clean state.
    try:
        import psycopg2
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute('DELETE FROM "gcal"."events"')
        cur.execute('DELETE FROM "email"."sent_log"')
        cur.execute('DELETE FROM "email"."messages"')
        conn.commit()
        cur.close()
        conn.close()
        print("Cleared schemas: gcal, email")
    except Exception as e:
        # Non-fatal: DB may be unavailable in some local runs.
        print(f"Warning: could not clear gcal/email schemas: {e}")

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
