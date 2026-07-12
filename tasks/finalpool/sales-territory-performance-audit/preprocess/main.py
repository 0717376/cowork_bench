#!/usr/bin/env python3
"""Preprocess script for task setup.

Source sales data lives in the central ClickHouse fork (schema sf_data,
russified centrally) plus the read-only seed files in initial_workspace
(rep_assignments.csv, quota_targets.xlsx). Nothing to inject here.

We idempotently clear any leftover sent emails / calendar events from a
previous run so they cannot pre-satisfy the phase-6 deliverable checks.
This is best-effort: if the DB or tables are absent it is a no-op.
"""

from argparse import ArgumentParser
from pathlib import Path
import os


def _clear_leftovers():
    try:
        import psycopg2
    except ImportError:
        return
    cfg = {
        "host": os.environ.get("PGHOST", "localhost"),
        "port": int(os.environ.get("PGPORT", "5432")),
        "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
        "user": os.environ.get("PGUSER", "eigent"),
        "password": os.environ.get("PGPASSWORD", "camel"),
    }
    try:
        conn = psycopg2.connect(**cfg)
        conn.autocommit = True
        cur = conn.cursor()
        for stmt in (
            "DELETE FROM email.messages",
            "DELETE FROM gcal.events",
        ):
            try:
                cur.execute(stmt)
            except Exception:
                conn.rollback()
        cur.close()
        conn.close()
    except Exception:
        pass


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    if args.agent_workspace:
        Path(args.agent_workspace).mkdir(parents=True, exist_ok=True)

    _clear_leftovers()

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
