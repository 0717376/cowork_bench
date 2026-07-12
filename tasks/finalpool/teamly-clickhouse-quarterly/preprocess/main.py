"""
Preprocess script for the Q4 sales review task (Teamly + ClickHouse fork).

1. Clears Google Sheets data (cells, sheets, spreadsheets) idempotently.
2. Ensures a Teamly space exists and seeds ONE knowledge-base page
   "Q4 2024 Sales Targets" with the regional revenue targets (in Russian)
   that the agent must read. This page is SOURCE DATA, not the answer.
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

SPACE_KEY = "SALES"
SPACE_NAME = "Продажи"
TARGETS_PAGE_TITLE = "Q4 2024 Sales Targets"

# Russified region -> target. Region strings copied verbatim from
# db/zzz_clickhouse_after_init.sql so seed <-> eval <-> groundtruth stay in sync.
TARGETS = [
    ("Азиатско-Тихоокеанский регион", 80000),
    ("Европа", 85000),
    ("Латинская Америка", 75000),
    ("Ближний Восток", 85000),
    ("Северная Америка", 80000),
]


def clear_gsheet(cur):
    """Clear all Google Sheets data."""
    print("[preprocess] Clearing Google Sheets data...")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.sheets")
    # Need to clear permissions before spreadsheets due to FK constraint
    try:
        cur.execute("DELETE FROM gsheet.permissions")
    except Exception:
        pass
    cur.execute("DELETE FROM gsheet.spreadsheets")
    print("[preprocess] Google Sheets data cleared.")


def seed_teamly_targets(cur):
    """Ensure a Teamly space + the source 'Q4 2024 Sales Targets' page exist.

    The page lists the 5 russified regions with their targets. Numeric values
    (80000/85000/75000/85000/80000) stay parseable. This is the data the agent
    must read; it is NOT the deliverable.
    """
    print("[preprocess] Seeding Teamly targets page...")

    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; skipping seed.")
        return

    # Ensure space
    cur.execute(
        """INSERT INTO teamly.spaces (key, name, description)
           VALUES (%s, %s, %s) ON CONFLICT (key) DO NOTHING""",
        (SPACE_KEY, SPACE_NAME, "Планы продаж, целевые показатели, квартальные обзоры."),
    )
    cur.execute("SELECT id FROM teamly.spaces WHERE key = %s", (SPACE_KEY,))
    space_id = cur.fetchone()[0]

    # Idempotent: remove any prior copy of this page in this space
    cur.execute(
        "DELETE FROM teamly.pages WHERE space_id = %s AND title = %s",
        (space_id, TARGETS_PAGE_TITLE),
    )

    lines = ["# Целевые показатели выручки по регионам на Q4 2024",
             "",
             "Руководство установило следующие квартальные цели по выручке:"]
    for region, target in TARGETS:
        lines.append(f"- {region}: {target}")
    body = "\n".join(lines)

    cur.execute(
        """INSERT INTO teamly.pages (space_id, title, body, author)
           VALUES (%s, %s, %s, %s)""",
        (space_id, TARGETS_PAGE_TITLE, body, "management"),
    )
    print(f"[preprocess] Seeded page '{TARGETS_PAGE_TITLE}' in space '{SPACE_KEY}'.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gsheet(cur)
        seed_teamly_targets(cur)
        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
