"""Preprocess script for sf-support-coaching-plan-excel-clickhouse-teamly-gcal.
Clears writable schemas (teamly user pages, gcal) and injects noise data.
Copies initial_workspace files to agent workspace.
"""
import argparse
import os
import shutil
from datetime import datetime

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="cowork_gym", user="eigent", password="camel")


def clear_schemas(conn):
    cur = conn.cursor()
    # Teamly: seed pages have id <= 3. Remove any user-created pages idempotently.
    try:
        cur.execute("SELECT to_regclass('teamly.pages')")
        if cur.fetchone()[0] is not None:
            cur.execute("DELETE FROM teamly.pages WHERE id > 3")
    except Exception as e:
        print(f"[preprocess] WARNING: teamly pages cleanup skipped: {e}")
    # GCal: clear all events (idempotent fresh slate).
    cur.execute("DELETE FROM gcal.events")
    conn.commit()
    cur.close()
    print("[preprocess] Cleared teamly user pages and gcal schema.")


def inject_noise(conn):
    cur = conn.cursor()

    # Ensure a dedicated Teamly space exists for support operations context.
    # This is an EMPTY space (no agent pages) — the agent must create the
    # 'Agent Coaching Tracker' space and the per-agent pages itself.
    try:
        cur.execute("SELECT to_regclass('teamly.spaces')")
        if cur.fetchone()[0] is not None:
            cur.execute("""
                INSERT INTO teamly.spaces (key, name, description)
                VALUES ('SUPPORT', 'Поддержка клиентов',
                        'Регламенты, метрики и материалы службы поддержки.')
                ON CONFLICT (key) DO NOTHING
            """)
            # Noise pages in the SUPPORT space (leftovers, NOT the deliverable).
            cur.execute("SELECT id FROM teamly.spaces WHERE key='SUPPORT'")
            sid = cur.fetchone()[0]
            noise_pages = [
                ("Заметки с ежедневной планёрки",
                 "# Заметки с планёрки\n\nОбсудили очередь тикетов и дежурства на неделю."),
                ("Трекер OKR на 1 квартал",
                 "# OKR Q1\n\nЦели команды поддержки на первый квартал."),
                ("График праздников 2026",
                 "# Праздники 2026\n\nНерабочие дни и дежурства в праздники."),
            ]
            for title, body in noise_pages:
                cur.execute(
                    "INSERT INTO teamly.pages (space_id, title, body, author) VALUES (%s, %s, %s, %s)",
                    (sid, title, body, "Координатор поддержки"),
                )
            print("[preprocess] Injected SUPPORT space + 3 noise Teamly pages.")
    except Exception as e:
        print(f"[preprocess] WARNING: teamly noise injection skipped: {e}")

    # Noise GCal events (RU titles, deliberately NOT containing 'coaching'/'коучинг').
    noise_events = [
        ("Еженедельная синхронизация команды", "2026-03-10 09:00:00", "2026-03-10 09:30:00"),
        ("Демонстрация продукта", "2026-03-12 14:00:00", "2026-03-12 15:00:00"),
        ("Общее собрание", "2026-03-13 16:00:00", "2026-03-13 17:00:00"),
    ]
    for summary, start, end in noise_events:
        cur.execute("""
            INSERT INTO gcal.events (summary, description, start_datetime, end_datetime, status)
            VALUES (%s, %s, %s, %s, %s)
        """, (summary, "Регулярная встреча команды", start, end, "confirmed"))
    print("[preprocess] Injected 3 noise GCal events.")

    conn.commit()
    cur.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    try:
        clear_schemas(conn)
        inject_noise(conn)
    finally:
        conn.close()

    # Copy initial_workspace files to agent workspace
    if args.agent_workspace and os.path.exists(args.agent_workspace):
        initial_ws = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "initial_workspace")
        if os.path.exists(initial_ws):
            for f in os.listdir(initial_ws):
                src = os.path.join(initial_ws, f)
                if os.path.isfile(src):
                    shutil.copy2(src, os.path.join(args.agent_workspace, f))
                    print(f"[preprocess] Copied {f} to agent workspace.")

    # Remove any previously created output files from agent workspace
    if args.agent_workspace:
        for fname in ["Agent_Scorecard.xlsx"]:
            path = os.path.join(args.agent_workspace, fname)
            if os.path.exists(path):
                os.remove(path)
                print(f"[preprocess] Removed old {fname}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
