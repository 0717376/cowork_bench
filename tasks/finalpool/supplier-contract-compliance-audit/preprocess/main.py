#!/usr/bin/env python3
"""Preprocess script for supplier-contract-compliance-audit task setup.

Source files (supplier_list.csv, contract_template.pdf, data.csv, config.json)
live in initial_workspace/ and are copied to the agent workspace by the harness.
ClickHouse (sf_data) data is read-only and globally seeded; nothing to seed here.

This script idempotently clears the email outbox and Google Calendar so the
phase-6 deliverables (audit emails + scheduled meetings) are verified honestly.
It does NOT pre-seed any answer files.
"""

from argparse import ArgumentParser
from pathlib import Path
import os


def clear_outbox():
    import psycopg2
    cfg = {
        "host": os.environ.get("PGHOST", "localhost"),
        "port": int(os.environ.get("PGPORT", "5432")),
        "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
        "user": "eigent",
        "password": "camel",
    }
    try:
        conn = psycopg2.connect(**cfg)
        conn.autocommit = True
        cur = conn.cursor()
        for stmt in (
            "DELETE FROM email.attachments",
            "DELETE FROM email.sent_log",
            "DELETE FROM email.messages",
            "DELETE FROM email.drafts",
            "DELETE FROM gcal.events",
        ):
            try:
                cur.execute(stmt)
            except Exception as e:
                print(f"[preprocess] skip ({e})")
        cur.close()
        conn.close()
        print("[preprocess] Email outbox and Google Calendar cleared.")
    except Exception as e:
        print(f"[preprocess] DB cleanup skipped: {e}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        Path(args.agent_workspace).mkdir(parents=True, exist_ok=True)

    clear_outbox()
    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
