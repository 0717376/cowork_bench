"""
Preprocess script for canvas-learning-analytics task.

Canvas is read-only (live English data — course names, grades, submissions).
This script:
- Ensures a Teamly space exists for the agent to create the dashboard page in.
- Clears prior dashboard pages (idempotency) — does NOT pre-create the answer.
- Injects RU noise pages (decoys, not the answer).
- Starts the mock analytics-benchmark HTTP server on port 30219.
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

PORT = 30219


def setup_teamly(cur):
    """Ensure an analytics space exists and clear prior dashboard pages."""
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is None:
        print("[preprocess] WARNING: teamly schema not found; "
              "run db/zzz_teamly_after_init.sql. Skipping teamly setup.")
        return

    # Dedicated space for the agent to drop the dashboard page into.
    cur.execute("""
        INSERT INTO teamly.spaces (key, name, description)
        VALUES ('ANALYTICS', 'Учебная аналитика',
                'Панели KPI, сравнение курсов с бенчмарками, отчёты проректора.')
        ON CONFLICT (key) DO NOTHING
    """)
    cur.execute("SELECT id FROM teamly.spaces WHERE key='ANALYTICS'")
    space_id = cur.fetchone()[0]

    # Idempotency: drop any dashboard pages left from previous runs.
    cur.execute("""
        DELETE FROM teamly.pages
         WHERE title ILIKE '%learning analytics%'
            OR title ILIKE '%analytics dashboard%'
            OR title ILIKE '%панель%аналитик%'
            OR title ILIKE '%учебн%аналитик%'
    """)

    # Inject RU noise pages (decoys, NOT the answer) into the ANALYTICS space.
    noise_pages = [
        ("Протокол заседания учебно-методической комиссии",
         "# Протокол\n\nОбсудили нагрузку преподавателей и расписание на весенний семестр.",
         "Канцелярия"),
        ("Регламент сбора отчётности по курсам",
         "# Регламент\n\nОтчёты по успеваемости сдаются в конце каждого семестра в деканат.",
         "Деканат"),
    ]
    for title, body, author in noise_pages:
        cur.execute("""
            INSERT INTO teamly.pages (space_id, title, body, author)
            VALUES (%s, %s, %s, %s)
        """, (space_id, title, body, author))
    print("[preprocess] Teamly ready: 'ANALYTICS' space ensured, prior dashboard pages cleared, RU noise injected.")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        setup_teamly(cur)
        # Seed leak: foreign course 9991 (CHR-RU-101) breaks the 22-course scope.
        cur.execute("DELETE FROM canvas.courses WHERE id = 9991")
        conn.commit()
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
