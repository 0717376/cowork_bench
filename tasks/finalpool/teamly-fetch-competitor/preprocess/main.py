"""
Preprocess script for the competitor-analysis task (Teamly variant).

This script:
1. Clears user-created Teamly pages and ensures a target space exists.
2. Extracts files/mock_api.tar.gz and starts an HTTP server on port 30156.
3. Removes the stale competitor_data.json from the workspace and ensures the
   memory dir exists in agent_workspace.
"""

import argparse
import asyncio
import os
import shutil
import tarfile

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_teamly(cur):
    """Clear user-created Teamly pages and ensure a target space exists.

    Seed pages have id <= 3 (global seed); only user-created pages are removed.
    No page is pre-created, so the agent's deliverable is not pre-seeded.
    """
    print("[preprocess] Clearing Teamly data...")
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute(
            """
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('COMPETITOR', 'Конкурентный анализ',
                    'Пространство для отчётов отдела продуктовой стратегии о конкурентах.')
            ON CONFLICT (key) DO NOTHING
            """
        )
    print("[preprocess] Teamly data cleared.")


async def setup_mock_server():
    """Extract mock_api.tar.gz and start HTTP server on port 30156."""
    print("[preprocess] Setting up mock competitor API...")

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_api.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")

    serve_dir = tmp_dir
    port = 30156

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --directory \"{serve_dir}\" "
        f"> \"{serve_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock competitor API running at http://localhost:{port}")


def prepare_workspace(agent_workspace):
    """Remove the stale competitor_data.json and ensure the memory dir exists.

    The legacy competitor_data.json describes a different (defunct) task and
    would mislead the agent, so it is removed if present.
    """
    if not agent_workspace:
        return

    stale = os.path.join(agent_workspace, "competitor_data.json")
    if os.path.exists(stale):
        os.remove(stale)
        print(f"[preprocess] Removed stale file {stale}")

    mem_dir = os.path.join(agent_workspace, "memory")
    os.makedirs(mem_dir, exist_ok=True)
    mem_file = os.path.join(mem_dir, "memory.json")
    if not os.path.exists(mem_file):
        import json
        with open(mem_file, "w") as f:
            json.dump({"notes": []}, f)
    print(f"[preprocess] Memory directory ensured at {mem_dir}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_teamly(cur)
        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()

    if args.agent_workspace:
        prepare_workspace(args.agent_workspace)

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
