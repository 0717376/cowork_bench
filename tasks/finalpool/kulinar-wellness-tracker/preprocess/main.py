"""
Preprocess for kulinar-wellness-tracker task.
- Clears agent-created Teamly pages and ensures a knowledge-base space.
- Starts mock HTTP server on port 30235.
"""
import argparse
import asyncio
import os
import shutil
import tarfile

import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

MOCK_PORT = 30235


WELLNESS_SPACE_KEY = "WELLNESS"


def clear_and_setup_teamly(conn):
    """Clear agent-created Teamly pages and ensure a wellness knowledge-base space.

    Seed pages have id <= 3 (см. zzz_teamly_after_init.sql) — их не трогаем.
    Не пред-засеваем страницу 'Client Wellness Dashboard': её создаёт агент.
    """
    with conn.cursor() as cur:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute(
                """
                INSERT INTO teamly.spaces (key, name, description)
                VALUES (%s, 'База знаний коуча по здоровью',
                        'Велнес-программы клиентов: рецепты, планы питания, отслеживание прогресса.')
                ON CONFLICT (key) DO NOTHING
                """,
                (WELLNESS_SPACE_KEY,),
            )
    conn.commit()
    print("[preprocess] Cleared agent Teamly pages and ensured wellness space")


async def setup_mock_server():
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    mock_src = os.path.join(files_dir, "mock_pages")
    if not os.path.exists(tar_path) and os.path.exists(mock_src):
        with tarfile.open(tar_path, "w:gz") as tar:
            tar.add(mock_src, arcname="mock_pages")

    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)
        serve_dir = os.path.join(tmp_dir, "mock_pages")
    else:
        serve_dir = tmp_dir
        if os.path.exists(mock_src):
            shutil.copytree(mock_src, os.path.join(tmp_dir, "mock_pages"))
            serve_dir = os.path.join(tmp_dir, "mock_pages")

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{MOCK_PORT}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {MOCK_PORT} --directory \"{serve_dir}\" "
        f"> \"{serve_dir}/server.log\" 2>&1 &")
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server at http://localhost:{MOCK_PORT}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        clear_and_setup_teamly(conn)
    finally:
        conn.close()

    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
