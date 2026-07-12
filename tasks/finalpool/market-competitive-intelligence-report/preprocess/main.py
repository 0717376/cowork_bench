#!/usr/bin/env python3
"""Preprocess for market-competitive-intelligence-report.

- Seeds the RU competitor source CSV into the agent workspace (the agent
  consumes this instead of crawling live websites).
- Clears writable leftovers idempotently: the agent's ClickHouse table in
  sf_data, gcal.events and email tables.
- Does NOT pre-create any deliverable (Excel/Word/table rows/email/event).
"""

from argparse import ArgumentParser
import os
import shutil
from pathlib import Path

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

TASK_ROOT = Path(__file__).resolve().parent.parent
SOURCE_CSV = TASK_ROOT / "initial_workspace" / "competitors.csv"

# Agent's consolidation target. Dropped on each run so reruns are idempotent
# and the agent's rows are never pre-seeded.
AGENT_TABLE = 'sf_data."MARKET_INTEL__PUBLIC__COMPETITORS"'

# Deliverables the agent must produce; removed if left over from a prior run.
DELIVERABLES = ["Competitive_Analysis.xlsx", "Competitive_Report.docx"]


def clear_writable(conn):
    cur = conn.cursor()
    cur.execute(f"DROP TABLE IF EXISTS {AGENT_TABLE}")
    try:
        cur.execute("DELETE FROM gcal.events")
    except Exception:
        conn.rollback()
    try:
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.messages")
    except Exception:
        conn.rollback()
    for tbl in ("email.sent_log", "email.drafts"):
        try:
            cur.execute(f"DELETE FROM {tbl}")
        except Exception:
            conn.rollback()
    conn.commit()
    cur.close()


def seed_workspace(agent_workspace):
    """Copy the source CSV into the agent workspace; remove stale deliverables."""
    agent_ws = Path(agent_workspace)
    agent_ws.mkdir(parents=True, exist_ok=True)
    if SOURCE_CSV.exists():
        shutil.copy2(SOURCE_CSV, agent_ws / SOURCE_CSV.name)
    for name in DELIVERABLES:
        f = agent_ws / name
        if f.exists():
            f.unlink()


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_writable(conn)
    finally:
        conn.close()

    if args.agent_workspace:
        seed_workspace(args.agent_workspace)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
