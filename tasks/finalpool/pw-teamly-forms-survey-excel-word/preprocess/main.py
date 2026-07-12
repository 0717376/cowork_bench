"""Preprocess script for pw-notion-gform-survey-excel-word."""
import os
import argparse, json, os, sys, shutil, tarfile, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()

    # Teamly: чистим только пользовательские страницы/пространства (сидовых 3 страницы / 2 пространства)
    try:
        cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    except Exception:
        pass
    try:
        cur.execute("DELETE FROM teamly.spaces WHERE id > 2")
    except Exception:
        pass

    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")

    # Шумовая страница в teamly (НЕ ответ задачи) — даём контекст рабочего пространства.
    cur.execute("""
        INSERT INTO teamly.pages (space_id, title, body, author)
        VALUES ((SELECT id FROM teamly.spaces WHERE key='TEAM'),
                'Старые заметки по проекту',
                'Архивные рабочие заметки. К текущему анализу опроса отношения не имеют.',
                'team-archive')
    """)
    conn.commit()
    cur.close()
    conn.close()


def setup_mock_server(port=30330):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill existing process on port
    try:
        subprocess.run(f"kill -9 $(lsof -ti:30330) 2>/dev/null", shell=True, timeout=5)
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
            f"nohup python3 -m http.server 30330 --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"Mock server started on port 30330")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)
    setup_mock_server(30330)

if __name__ == "__main__":
    main()
