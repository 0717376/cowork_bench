"""
Preprocess for fetch-sf-sales-territory-ppt-email-gcal task.

Clears gcal events and email data. Extracts mock_pages.tar.gz and starts HTTP server on port 30215.
ClickHouse SALES_DW (PG schema sf_data, russified centrally) is read-only - no injection needed.
The CRM JSON in mock_pages.tar.gz uses the same russified region names as the DB so the join key matches.
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
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_db(cur):
    """Clear gcal events and email data."""
    print("[preprocess] Clearing gcal events and email data...")
    cur.execute("DELETE FROM gcal.events")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] DB cleared.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30215."""
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    port = 30215

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --bind 0.0.0.0 --directory \"{mock_dir}\" "
        f"> \"{mock_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)

    # Health check: poll the mock server until it serves the CRM data, fail
    # preprocess if it never comes up so the agent never starts against a dead
    # endpoint. Bound to 0.0.0.0 above so localhost:30215 is reachable from the
    # agent container's fetch MCP.
    health_url = f"http://localhost:{port}/api/territory_data.json"
    ready = False
    for attempt in range(20):
        check = await asyncio.create_subprocess_shell(
            f"curl -fsS -o /dev/null -w '%{{http_code}}' {health_url}",
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        out, _ = await check.communicate()
        if check.returncode == 0 and out.decode().strip() == "200":
            ready = True
            break
        await asyncio.sleep(0.5)

    if not ready:
        raise RuntimeError(
            f"[preprocess] Mock server health check failed: {health_url} "
            f"did not return 200 after polling."
        )
    print(f"[preprocess] Mock server running and serving CRM data at http://localhost:{port}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        clear_db(cur)
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] DB error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
