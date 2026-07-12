"""Preprocess script for fetch-sf-sales-competitor-excel-teamly."""
import os
import argparse, json, os, sys, shutil, tarfile, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Teamly space dedicated to this task. Created idempotently; the agent's
# deliverable page is NOT pre-seeded here.
SPACE_KEY = "SFCOMP"
SPACE_NAME = "Анализ конкурентов по продажам"
SPACE_DESC = "Пространство для отчётов и дашбордов по конкурентному анализу продаж."


def clear_writable_schemas():
    """Clear user-created Teamly pages and ensure the target space exists.

    Global seed pages have id <= 3 (zzz_teamly_after_init.sql); only
    user-created pages (id > 3) are removed so reruns are idempotent and the
    agent's deliverable is never pre-seeded.
    """
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT to_regclass('teamly.pages')")
    if cur.fetchone()[0] is not None:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    cur.execute("SELECT to_regclass('teamly.spaces')")
    if cur.fetchone()[0] is not None:
        cur.execute(
            """
            INSERT INTO teamly.spaces (key, name, description)
            VALUES (%s, %s, %s)
            ON CONFLICT (key) DO NOTHING
            """,
            (SPACE_KEY, SPACE_NAME, SPACE_DESC),
        )
    conn.commit()
    cur.close()
    conn.close()


def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)


def inject_data(launch_time):
    """Inject a RU noise page so the workspace is not empty.

    This is plausible-but-irrelevant content; it does NOT contain the
    deliverable analysis and must not match the dashboard check.
    """
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO teamly.pages (space_id, title, body, author)
        VALUES (
            (SELECT id FROM teamly.spaces WHERE key=%s),
            'Старые заметки проекта',
            E'# Старые заметки проекта\n\nАрхивные черновики прошлого квартала. Не относится к текущему анализу конкурентов.',
            'admin'
        )
        """,
        (SPACE_KEY,),
    )
    conn.commit()
    cur.close()
    conn.close()


def setup_mock_server(port=30335):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill existing process on port
    try:
        subprocess.run(f"kill -9 $(lsof -ti:{port}) 2>/dev/null", shell=True, timeout=5)
    except Exception:
        pass
    time.sleep(0.5)

    # Extract mock pages
    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)

    # Start HTTP server
    mock_dir = os.path.join(tmp_dir, "mock_pages")
    if os.path.exists(mock_dir):
        log_path = os.path.join(mock_dir, "server.log")
        subprocess.Popen(
            f"nohup python3 -m http.server {port} --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"Mock server started on port {port}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)
    setup_mock_server(30335)

if __name__ == "__main__":
    main()
