"""Preprocess for moex-market-news-teamly.

- moex.news has NO seeded rows in the global seed, yet the whole task is a news
  digest. We inject RU market-news rows that the agent consumes via the
  moex-finance MCP (get_moex_finance_news). The JSON shape matches what the MCP
  reads (content.contentType == "STORY", content.title/summary/description/
  canonicalUrl.url) AND what the eval reads
  (content.provider.displayName, content.title, content.pubDate).
  This is SOURCE data the agent reads, NOT the agent's deliverable — we do not
  pre-create the Excel file, the Teamly page or the email.
- Idempotent: DELETE FROM moex.news then re-INSERT, so repeated runs stay clean.
- Teamly: remove only leftover deliverable pages (idempotency); keep seeded
  TEAM/TRIPS spaces and their seed pages. Do NOT pre-create the digest page.
- Email: clear writable email tables for a clean state.
"""
import os
import json
import argparse
import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

# RU market news: (symbol, title, publisher, pubDate). Publishers are real RU
# financial media. MGNT.ME (Магнит) is intentionally the top publisher driver.
NEWS = [
    ("SBER.ME", "Сбербанк отчитался о росте чистой прибыли за квартал", "РБК", "2026-05-26T08:15:00Z"),
    ("SBER.ME", "Набсовет Сбербанка рекомендовал дивиденды по итогам года", "Ведомости", "2026-05-25T11:40:00Z"),
    ("SBER.ME", "Сбер расширяет линейку ИИ-сервисов для бизнеса", "Интерфакс", "2026-05-24T09:05:00Z"),
    ("GAZP.ME", "Газпром нарастил экспорт газа в дружественные страны", "ТАСС", "2026-05-27T07:30:00Z"),
    ("GAZP.ME", "Совет директоров Газпрома обсудит инвестпрограмму", "Прайм", "2026-05-25T14:10:00Z"),
    ("GAZP.ME", "Аналитики оценили перспективы дивидендов Газпрома", "РБК", "2026-05-23T10:20:00Z"),
    ("LKOH.ME", "ЛУКОЙЛ увеличил добычу на новых месторождениях", "Коммерсантъ", "2026-05-26T12:00:00Z"),
    ("LKOH.ME", "ЛУКОЙЛ объявил программу обратного выкупа акций", "Ведомости", "2026-05-24T15:45:00Z"),
    ("TCSG.ME", "Группа Т-Технологии (TCS) показала рост числа клиентов", "Интерфакс", "2026-05-27T09:50:00Z"),
    ("TCSG.ME", "Т-Банк запустил новый инвестиционный продукт", "Прайм", "2026-05-25T16:25:00Z"),
    ("TCSG.ME", "Аналитики повысили прогноз по бумагам TCS Group", "РБК", "2026-05-22T08:40:00Z"),
    ("MGNT.ME", "Магнит открыл тысячный магазин нового формата", "ТАСС", "2026-05-28T07:00:00Z"),
    ("MGNT.ME", "Магнит отчитался о росте выручки в первом квартале", "Коммерсантъ", "2026-05-26T13:30:00Z"),
    ("MGNT.ME", "Совет директоров Магнита рассмотрит дивидендную политику", "Ведомости", "2026-05-25T10:15:00Z"),
    ("MGNT.ME", "Магнит расширяет сеть дискаунтеров в регионах", "Прайм", "2026-05-23T11:05:00Z"),
    ("MTSS.ME", "МТС представила обновлённую экосистему сервисов", "Интерфакс", "2026-05-27T14:35:00Z"),
    ("MTSS.ME", "МТС увеличила инвестиции в развитие сети 5G", "РБК", "2026-05-24T09:55:00Z"),
    ("MTSS.ME", "Аналитики отметили устойчивость дивидендов МТС", "ТАСС", "2026-05-22T12:50:00Z"),
]


def build_data(symbol, title, publisher, pub_date):
    """JSON shape matching the moex-finance MCP and the evaluation reads."""
    slug = title.lower().replace(" ", "-")[:40]
    return {
        "id": f"{symbol}-{pub_date}",
        "content": {
            "contentType": "STORY",
            "title": title,
            "summary": title,
            "description": title,
            "pubDate": pub_date,
            "provider": {"displayName": publisher},
            "canonicalUrl": {"url": f"https://news.example.ru/{symbol}/{slug}"},
        },
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Inject RU source news (idempotent).
    cur.execute("DELETE FROM moex.news")
    for symbol, title, publisher, pub_date in NEWS:
        data = build_data(symbol, title, publisher, pub_date)
        cur.execute(
            "INSERT INTO moex.news (symbol, data) VALUES (%s, %s)",
            (symbol, json.dumps(data, ensure_ascii=False)),
        )

    # Teamly idempotency: drop leftover deliverable pages, keep seeds.
    cur.execute(
        "DELETE FROM teamly.pages "
        "WHERE lower(title) LIKE '%рыночн%' OR lower(title) LIKE '%market%' "
        "OR lower(title) LIKE '%аналитик%' OR lower(title) LIKE '%дайджест%'"
    )

    # Clear writable email tables (FK order).
    for table in ["attachments", "sent_log", "drafts", "messages", "folders", "account_config"]:
        cur.execute(f'DELETE FROM email."{table}"')

    conn.commit()
    cur.close()
    conn.close()
    print(f"[preprocess] Injected {len(NEWS)} RU news rows into moex.news; "
          "cleared leftover teamly digest pages and email tables.")


if __name__ == "__main__":
    main()
