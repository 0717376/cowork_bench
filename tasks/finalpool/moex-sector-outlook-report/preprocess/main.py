"""Preprocess for yf-sector-outlook-report (RU: moex-finance + teamly).

- Seeds an empty teamly space 'RESEARCH' with a parent page 'Research Library'
  (idempotent). Does NOT pre-create the deliverable 'Sector Outlook Report' page.
- moex-finance is read-only (live MOEX series come from the globally seeded
  moex.* schema), so nothing to inject there.
- Starts the mock research-portal API server on port 30226 serving
  files/mock_pages (sector_outlook.json with RU sector labels).
"""
import argparse
import asyncio
import os
import shutil

import psycopg2

PORT = 30226
DB = dict(
    host=os.environ.get("PGHOST", "localhost"),
    port=5432,
    dbname=os.environ.get("PGDATABASE", "cowork_gym"),
    user="eigent",
    password="camel",
)


def clear_and_setup_teamly():
    conn = psycopg2.connect(**DB)
    conn.autocommit = True
    cur = conn.cursor()

    # Ensure a RESEARCH space exists (idempotent).
    cur.execute(
        "INSERT INTO teamly.spaces (key, name, description) VALUES (%s, %s, %s) "
        "ON CONFLICT (key) DO NOTHING",
        ("RESEARCH", "Библиотека исследований",
         "Аналитические отчёты, секторальные обзоры, исследовательские материалы."),
    )
    cur.execute("SELECT id FROM teamly.spaces WHERE key = 'RESEARCH'")
    space_id = cur.fetchone()[0]

    # Remove only leftover deliverable pages (idempotency); keep seeded content.
    cur.execute(
        "DELETE FROM teamly.pages "
        "WHERE space_id = %s AND lower(title) LIKE '%%sector outlook report%%'",
        (space_id,),
    )

    # Seed an empty parent page 'Research Library' if not present.
    cur.execute(
        "SELECT id FROM teamly.pages WHERE space_id = %s AND title = %s",
        (space_id, "Research Library"),
    )
    if cur.fetchone() is None:
        cur.execute(
            "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
            (space_id, "Research Library",
             "# Research Library\n\nКаталог исследовательских отчётов управляющей компании.",
             "admin"),
        )

    cur.close()
    conn.close()
    print("[preprocess] Teamly: RESEARCH space + 'Research Library' page ensured; "
          "leftover report pages cleared.")


async def run_command(cmd):
    proc = await asyncio.create_subprocess_shell(
        cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    await proc.wait()


async def setup_mock_server():
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    mock_dir = os.path.join(task_root, "files", "mock_pages")
    tmp_dir = os.path.join(task_root, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    await run_command(f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null")
    await asyncio.sleep(0.5)
    log_path = os.path.join(tmp_dir, "server.log")
    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {PORT} --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock API server running at http://localhost:{PORT}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clear_and_setup_teamly()
    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
