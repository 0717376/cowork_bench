"""Preprocess: MOEX finance is read-only.

Idempotency for teamly: remove any leftover "Earnings Analysis" pages from a
previous run, but KEEP the seeded TEAM/TRIPS spaces and their seed pages.
We do NOT pre-create the "Earnings Analysis Key Insights" page (that is the
agent's deliverable).
"""
import os
import argparse
import psycopg2

DB = dict(
    host=os.environ.get("PGHOST", "localhost"),
    port=5432,
    dbname=os.environ.get("PGDATABASE", "cowork_gym"),
    user="eigent",
    password="camel",
)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    # Remove only leftover deliverable pages (idempotency), keep seeded content.
    cur.execute(
        "DELETE FROM teamly.pages "
        "WHERE lower(title) LIKE '%earnings%' OR lower(title) LIKE '%insights%'"
    )
    conn.commit()
    cur.close()
    conn.close()
    print("[preprocess] Cleared leftover Earnings teamly pages; seeds preserved.")


if __name__ == "__main__":
    main()
