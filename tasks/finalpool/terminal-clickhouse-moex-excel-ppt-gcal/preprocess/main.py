"""Preprocess for terminal-clickhouse-moex-excel-ppt-gcal.
Clears gcal, injects conflicting calendar events for the scheduling step.
ClickHouse (sf_data) and MOEX (moex) data are read-only."""
import argparse
import glob as globmod
import json
import os
import uuid
from datetime import datetime, timedelta

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent", "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    launch = datetime.strptime(args.launch_time or "2026-03-07 10:00:00", "%Y-%m-%d %H:%M:%S")

    def dt(days, hours, minutes=0):
        """Return datetime string as offset from launch_time."""
        return (launch + timedelta(days=days, hours=hours, minutes=minutes)).strftime("%Y-%m-%d %H:%M:%S")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Clear gcal
        cur.execute("DELETE FROM gcal.events")
        conn.commit()
        print("[preprocess] Cleared gcal events.")

        # Inject conflicting events for the week starting launch+2 days
        # to make scheduling non-trivial
        conflicts = [
            # launch+1d23h = Monday Mar 9 09:00 (morning blocked)
            ("Совет директоров", dt(1, 23), dt(2, 1),
             "Ежеквартальное заседание совета директоров", "confirmed"),
            ("Обед с инвесторами", dt(2, 2), dt(2, 3, 30),
             "Обед по связям с инвесторами", "confirmed"),
            ("Обзор продуктов", dt(2, 4), dt(2, 6),
             "Обзор дорожной карты продуктов", "confirmed"),

            # Tuesday Mar 10: mostly free (9-11 available = first 2hr slot)
            ("Ежедневный статус", dt(2, 22, 30), dt(2, 23),
             "Ежедневная планёрка команды", "confirmed"),
            ("Звонок с клиентом", dt(3, 4), dt(3, 5),
             "Звонок с крупным клиентом", "confirmed"),

            # Wednesday Mar 11: heavily booked
            ("Стратегическая сессия", dt(3, 23), dt(4, 7),
             "Полнодневная стратегическая сессия", "confirmed"),

            # Thursday Mar 12: afternoon packed
            ("HR-обзор", dt(4, 23), dt(5, 0),
             "Ежеквартальный HR-обзор", "confirmed"),
            ("Планирование бюджета", dt(5, 0, 30), dt(5, 2, 30),
             "Планирование бюджета на 2027 финансовый год", "confirmed"),
            ("Общее собрание", dt(5, 4), dt(5, 6),
             "Общее собрание компании", "confirmed"),

            # Friday Mar 13: morning free
            ("Пятничный обед", dt(6, 2), dt(6, 3),
             "Командный обед", "confirmed"),
            ("Ретроспектива спринта", dt(6, 5), dt(6, 6, 30),
             "Ретроспектива спринта", "confirmed"),
        ]

        # Also inject some noise events outside the target week
        noise = [
            ("Еженедельный 1:1", dt(9, 0), dt(9, 0, 30),
             "Встреча с руководителем", "confirmed"),
            ("Занятие йогой", dt(6, 21), dt(6, 22),
             "Субботняя йога", "confirmed"),
        ]

        for summary, start, end, desc, status in conflicts + noise:
            cur.execute("""
                INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (str(uuid.uuid4()), summary, desc, start, end, status))

        conn.commit()
        print(f"[preprocess] Injected {len(conflicts)} conflict events + {len(noise)} noise events.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Clean agent workspace
    if args.agent_workspace:
        for pattern in ["Investment_Committee_Briefing.xlsx", "Committee_Briefing.pptx",
                        "briefing_notes.txt", "compute_growth.py", "market_comparison.py",
                        "market_comparison.json"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
