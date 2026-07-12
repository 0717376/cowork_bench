"""Preprocess for insales-review-competitor-study. Sets up mock server and clears teamly."""
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
PORT = 30203


def clear_teamly(cur):
    """Clear user-created Teamly pages, ensure a space, inject noise pages."""
    print("[preprocess] Clearing Teamly data...")
    # Drop user-created pages (seed pages have id <= 3); ensure a space.
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute("""
            INSERT INTO teamly.spaces (key, name, description)
            VALUES ('REVIEWS', 'Анализ отзывов',
                    'База знаний по качеству товаров и отзывам.')
            ON CONFLICT (key) DO NOTHING
        """)

    # Resolve target space id for noise pages.
    cur.execute("SELECT id FROM teamly.spaces WHERE key = 'REVIEWS'")
    row = cur.fetchone()
    if row is None:
        cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
        row = cur.fetchone()
    space_id = row[0] if row else None

    # Inject noise pages (RU titles) so the agent has leftover content to ignore;
    # these must NOT satisfy the analysis-page check.
    if space_id is not None:
        cur.execute(
            "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
            (space_id, "База знаний",
             "Корневая страница базы знаний по качеству товаров.", "team"),
        )
        cur.execute(
            "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
            (space_id, "Старый анализ за Q4 2025",
             "Архивные заметки за прошлый квартал. Не относится к текущей задаче.", "team"),
        )
    print("[preprocess] Teamly data cleared and noise pages injected.")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server."""
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)

    serve_dir = os.path.join(tmp_dir, "mock_pages")

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {PORT} --directory \"{serve_dir}\" "
        f"> \"{tmp_dir}/server.log\" 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{PORT}")


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
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
