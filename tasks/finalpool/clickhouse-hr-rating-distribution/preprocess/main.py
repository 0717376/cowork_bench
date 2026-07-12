"""Preprocess: очистка данных email для чистого состояния перед проверкой.

Данные ClickHouse (схема sf_data, HR_ANALYTICS) русифицируются централизованно
в db/zzz_clickhouse_after_init.sql — здесь правки данных MCP не требуются.
"""
import os
import argparse
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "cowork_gym", "user": "eigent", "password": "camel"}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    for t in ["email.attachments", "email.sent_log", "email.messages"]:
        cur.execute(f"DELETE FROM {t}")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    conn.commit()
    cur.close()
    conn.close()
    print("Email data cleared")

if __name__ == "__main__":
    main()
