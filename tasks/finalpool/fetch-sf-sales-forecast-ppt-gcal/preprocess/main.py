"""
Preprocess script for fetch-sf-sales-forecast-ppt-gcal task.

This script:
1. Clears gcal data
2. Injects noise calendar event
3. Extracts mock_pages.tar.gz and starts HTTP server on port 30209
NOTE: Does NOT modify ClickHouse (sf_data) warehouse data (read-only).
"""

import argparse
import asyncio
import json
import os
import shutil
import sys
import tarfile
import urllib.request
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_gcal(cur):
    """Clear all Google Calendar events."""
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] Google Calendar events cleared.")


def inject_noise_gcal(cur):
    """Inject noise calendar event."""
    print("[preprocess] Injecting noise calendar event...")
    cur.execute(
        "INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime) "
        "VALUES (%s, %s, %s, %s, %s)",
        (
            f"noise-{uuid.uuid4()}",
            "Еженедельная планёрка команды",
            "Регулярная еженедельная планёрка команды",
            "2026-03-10T09:00:00+00:00",
            "2026-03-10T09:30:00+00:00",
        ),
    )
    print("[preprocess] Noise calendar event injected.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30209."""
    print("[preprocess] Setting up mock market research API...")

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    port = 30209

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup {sys.executable} -m http.server {port} --directory \"{mock_dir}\" "
        f"> \"{mock_dir}/server.log\" 2>&1 &"
    )

    # Readiness gate: poll the projections endpoint until it returns 200.
    # The agent depends on this API being reachable; fail loudly if it never
    # comes up rather than letting the agent loop on a dead endpoint.
    health_url = f"http://localhost:{port}/api/projections.json"
    deadline = asyncio.get_event_loop().time() + 15.0
    last_err = None
    while asyncio.get_event_loop().time() < deadline:
        try:
            with urllib.request.urlopen(health_url, timeout=2) as resp:
                if resp.status == 200:
                    print(f"[preprocess] Mock API server running at http://localhost:{port}")
                    break
                last_err = f"HTTP {resp.status}"
        except Exception as e:  # noqa: BLE001 - any failure means not ready yet
            last_err = e
        await asyncio.sleep(0.5)
    else:
        raise RuntimeError(
            f"[preprocess] Mock API server did not become ready at {health_url} "
            f"within 15s (last error: {last_err})"
        )


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gcal(cur)
        inject_noise_gcal(cur)
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
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
