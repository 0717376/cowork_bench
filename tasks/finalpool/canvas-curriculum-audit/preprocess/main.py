"""
Preprocess script for canvas-curriculum-audit task.
Starts mock HTTP server on port 30217.
Canvas is read-only (кроме удаления утёкшего в сид курса 9991).
"""
import argparse
import asyncio
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

PORT = 30217


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Seed leak: foreign course 9991 (CHR-RU-101) breaks the 22-course scope.
        cur.execute("DELETE FROM canvas.courses WHERE id = 9991")
        conn.commit()
        print("[preprocess] Removed leaked course 9991 (CHR-RU-101).")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Start mock HTTP server
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    serve_dir = os.path.join(task_root, "tmp", "mock_pages")

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {PORT} --directory \"{serve_dir}\" "
        f"> \"{serve_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{PORT}")
    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
