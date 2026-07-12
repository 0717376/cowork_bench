"""Предобработка: очистка записываемой схемы gsheet для чистого старта.

Данные продаж sf_data (SALES_DW) сидируются глобально и русифицируются
централизованно через db/zzz_clickhouse_after_init.sql (CATEGORY
'tv, audio & cameras' -> 'ТВ, аудио и камеры'), поэтому здесь их не трогаем.
Очищаем только схему gsheet (идемпотентно), чтобы предыдущие запуски не оставляли
таблицу 'Product Category Report'. Ответы НЕ предсоздаём.
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
    cur.execute('DELETE FROM gsheet.cells')
    cur.execute('DELETE FROM gsheet.sheets')
    cur.execute('DELETE FROM gsheet.permissions')
    cur.execute('DELETE FROM gsheet.spreadsheets')
    cur.execute('DELETE FROM gsheet.folders')
    conn.commit()
    cur.close()
    conn.close()
    print("Данные очищены для схем: gsheet")

if __name__ == "__main__":
    main()
