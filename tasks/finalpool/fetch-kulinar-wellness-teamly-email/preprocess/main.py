"""Preprocess script for fetch-kulinar-wellness-teamly-email."""
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

    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")

    # Teamly (RU-аналог Confluence): удаляем пользовательские страницы
    # (сид-страницы имеют id <= 3), гарантируем наличие пространства.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('WELLNESS', 'Здоровое питание',
                        'Пространство команды по анализу кулинарного оздоровительного отчёта.')
                ON CONFLICT (key) DO NOTHING
            """)
    except Exception as e:
        print(f"[preprocess] WARNING: teamly cleanup skipped: {e}")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")

    # Шумовое письмо (RU прозой) — агент должен его проигнорировать.
    inbox_id = 1
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if row: inbox_id = row[0]
    cur.execute("""INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
        VALUES (%s, %s, %s, %s, %s, %s, %s, true)""",
        (inbox_id, '<noise-cook_wellness-001@co.com>', 'Еженедельная рассылка', 'newsletter@company.com',
         json.dumps(['all@company.com']), launch_dt - timedelta(hours=6),
         'Новости компании за неделю: в комнате отдыха установили новую кофемашину.'))

    # Шумовая страница Teamly (RU-заголовок) — пользовательский остаток,
    # который НЕ должен удовлетворять проверку дашборда.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("SELECT id FROM teamly.spaces WHERE key = 'WELLNESS'")
            row = cur.fetchone()
            if row is None:
                cur.execute("SELECT id FROM teamly.spaces ORDER BY id LIMIT 1")
                row = cur.fetchone()
            space_id = row[0] if row else None
            if space_id is not None:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) "
                    "VALUES (%s, %s, %s, %s)",
                    (space_id, "Старые заметки проекта",
                     "Архивные заметки команды. Не относится к текущей задаче.",
                     "team"),
                )
    except Exception as e:
        print(f"[preprocess] WARNING: noise teamly page skipped: {e}")
    conn.commit()
    cur.close()
    conn.close()


def setup_mock_server(port=30321):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill existing process on port
    try:
        subprocess.run(f"kill -9 $(lsof -ti:30321) 2>/dev/null", shell=True, timeout=5)
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
            f"nohup python3 -m http.server 30321 --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"Mock server started on port 30321")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)
    setup_mock_server(30321)

if __name__ == "__main__":
    main()
