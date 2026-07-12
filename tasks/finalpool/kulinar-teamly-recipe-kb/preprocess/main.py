"""
Preprocess for kulinar-teamly-recipe-kb task.
- Idempotently clears any Teamly space/pages left over from previous runs of
  this task (the recipe-collection space the agent must create).
- Ensures memory directory and file exist in agent workspace (empty scaffold,
  NOT an answer).

We intentionally do NOT pre-create the recipe space, recipe pages, the xlsx, or
any memory entities — the agent must produce them itself.

Prerequisites:
  - PostgreSQL cowork_gym database running on localhost:5432
  - Teamly schema seeded (db/zzz_teamly_after_init.sql)
"""
import argparse
import json
import os

import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

# Marker substrings identifying this task's space so cleanup is idempotent and
# does not touch the globally-seeded demo spaces.
SPACE_MARKERS = ("коллекция рецептов", "recipe collection")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Idempotently remove this task's recipe space (and its pages via cascade).
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute("SELECT id, name FROM teamly.spaces")
        for sid, name in cur.fetchall():
            nl = (name or "").lower()
            if any(m in nl for m in SPACE_MARKERS):
                cur.execute("DELETE FROM teamly.pages WHERE space_id = %s", (sid,))
                cur.execute("DELETE FROM teamly.spaces WHERE id = %s", (sid,))
        conn.commit()
        print("Cleared leftover Teamly recipe-collection space(s)")
    cur.close()
    conn.close()

    # Ensure memory directory and file exist (empty scaffold).
    if args.agent_workspace:
        mem_dir = os.path.join(args.agent_workspace, "memory")
        os.makedirs(mem_dir, exist_ok=True)
        mem_file = os.path.join(mem_dir, "memory.json")
        if not os.path.exists(mem_file):
            with open(mem_file, "w") as f:
                json.dump({"entities": [], "relations": []}, f)
        print(f"Memory file ensured at {mem_file}")

    print("Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
