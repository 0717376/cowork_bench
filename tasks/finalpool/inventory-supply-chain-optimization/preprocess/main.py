"""Preprocess script for inventory-supply-chain-optimization.

Sources read honestly by the agent:
  - InSales store (wc.* schema, globally seeded) -> product stock / total_sales
  - ClickHouse SALES_DW (sf_data."SALES_DW__PUBLIC__ORDERS", globally seeded) -> demand history
  - mock supplier catalog HTTP server on :30306 -> lead times / MOQ / reliability (RUB)

This script seeds NO answers. It only:
  - clears writable output schemas (gcal / email / gsheet) idempotently,
  - injects a couple of noise rows so the agent must filter,
  - (re)starts the mock supplier-catalog HTTP server.
The wc.* and sf_data seeds are global and intentionally left untouched.
"""
import os
import argparse, json, shutil, tarfile, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MOCK_PORT = 30306


def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)


def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    # Calendar / email / google-sheet are the writable deliverable schemas.
    cur.execute("DELETE FROM gcal.events")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.sheets")
    cur.execute("DELETE FROM gsheet.spreadsheets")
    conn.commit()
    cur.close()
    conn.close()


def inject_noise(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")

    # Noise calendar event (must NOT be confused with the review meeting).
    cur.execute("""INSERT INTO gcal.events (summary, start_datetime, end_datetime, status)
        VALUES ('Синхронизация команды', %s, %s, 'confirmed')""",
        (launch_dt.replace(hour=14), launch_dt.replace(hour=14, minute=30)))

    # Noise inbox email.
    inbox_id = 1
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if row:
        inbox_id = row[0]
    cur.execute("""INSERT INTO email.messages
        (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
        VALUES (%s, '<noise-inv-001@co.com>', 'Корпоратив в эту пятницу',
                'events@company.com', %s, %s,
                'Приглашаем на пятничную встречу команды!', true)""",
        (inbox_id, json.dumps(['all@company.com']), launch_dt - timedelta(hours=4)))

    conn.commit()
    cur.close()
    conn.close()


def setup_mock_server(port=MOCK_PORT):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill existing process on the port.
    try:
        subprocess.run(f"kill -9 $(lsof -ti:{port}) 2>/dev/null", shell=True, timeout=5)
    except Exception:
        pass
    time.sleep(0.5)

    # Extract mock pages.
    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)

    # Start the HTTP server serving the supplier catalog.
    mock_dir = os.path.join(tmp_dir, "mock_pages")
    if os.path.exists(mock_dir):
        log_path = os.path.join(mock_dir, "server.log")
        subprocess.Popen(
            f"nohup python3 -m http.server {port} --directory \"{mock_dir}\" > \"{log_path}\" 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"Mock supplier-catalog server started on port {port}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-01 09:00:00")
    args = parser.parse_args()

    if args.agent_workspace:
        from pathlib import Path
        Path(args.agent_workspace).mkdir(parents=True, exist_ok=True)

    try:
        clear_writable_schemas()
        inject_noise(args.launch_time)
    except Exception as e:
        print(f"[WARN] DB setup skipped: {e}")

    setup_mock_server(MOCK_PORT)
    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
