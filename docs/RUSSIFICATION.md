# Русификация таска: рецепт

Пошагово, без воды. Сценарий — берём апстримный таск (например `12306-sf-hr-training-travel-excel-email-gcal`) и адаптируем под РФ-реалии (Москва→Казань, РЖД, 1С:ЗУП).

Состояние на 2026-05-25: 8 тасков переведено (`rzd-*`, `hr1c-*`). Список под перевод — в `memory/project_russification.md`.

---

## 0. Решить уровень русификации

| Уровень | Что меняем | Когда выбирать |
|---|---|---|
| **Soft** (sys_prompt + MCP list) | `docs/agent_system_prompt.md`, `task_config.json` (rail_12306→rzd, snowflake→hr1c) | Логика таска нейтральна; данные апстрима подойдут. |
| **Full** (геолокализация) | + `task.md`, `evaluation/main.py`, `preprocess/main.py`, `groundtruth_workspace/`, `initial_workspace/` | Таск завязан на CN/US-реалии (Пекин→Шанхай, snowflake-HR). |

Все нынешние `rzd-*` / `hr1c-*` — **full**.

---

## 1. Scaffold

```bash
./cowork_bench new --from 12306-sf-hr-training-travel-excel-email-gcal \
                   --to   rzd-hr1c-training-trip-kazan-excel-email-gcal
```

`scripts/scaffold_ru_task.py` копирует папку, патчит `agent_system_prompt.md` на русский шаблон, меняет MCP в `task_config.json` (`rail_12306`→`rzd`, `snowflake`→`hr1c`, drop `notion`). Дальше — руками.

---

## 2. Новый домен → форк MCP-сервера (только если нужно)

Если используются уже существующие `rzd`/`hr1c` — **пропустить**. Если домен новый (например, российская e-commerce):

1. Скопировать `local_servers/mcp-snowflake-server/` → `local_servers/<new>-mcp-server/`.
2. В `src/<new>_mcp_server/`:
   - `db_client.py`: `SnowflakeDB`→`<New>DB`, `PG_SCHEMA = os.environ.get("PG_SCHEMA", "<new>_data")`, заменить `'sf_data'` → `f'{PG_SCHEMA}'` во всех SQL-запросах.
   - `server.py`: bulk rename `SnowflakeDB`→`<New>DB`, `mcp_snowflake_server`→`<new>_mcp_server`, обновить display name.
   - `__init__.py`: убрать `import snowflake.connector`, упростить `parse_args` (оставить `--private_key_path` как no-op для CLI-совместимости).
3. `pyproject.toml`:
   - `name = "<new>_mcp_server"`.
   - `[project.scripts]` → `<new>_mcp_server = "<new>_mcp_server:main"`.
   - Убрать `snowflake-connector-python`, `snowpark`, `cryptography`, `pyOpenSSL`. **Не забыть `pyyaml>=6.0`** (легко пропустить → `ModuleNotFoundError` в рантайме).
4. Пересобрать образ: `docker compose build cowork`.

---

## 3. Миграция данных: `db/zzz_<name>_after_init.sql`

Префикс `zzz_` критичен: docker-entrypoint грузит файлы alphabetically под `en_US.utf8`, `_` < `.`, поэтому `zzz_*.sql` идёт **после** `init.sql.gz`. См. [`memory/feedback_db_init_order.md`](../memory/feedback_db_init_order.md).

```sql
CREATE SCHEMA IF NOT EXISTS hr1c_data;
ALTER SCHEMA hr1c_data OWNER TO eigent;

CREATE TABLE hr1c_data."HR__PUBLIC__EMPLOYEES" (
    "EMPLOYEE_ID"   numeric NOT NULL PRIMARY KEY,
    "EMPLOYEE_NAME" character varying(200) NOT NULL,
    ...
);
ALTER TABLE hr1c_data."HR__PUBLIC__EMPLOYEES" OWNER TO eigent;

INSERT INTO hr1c_data."HR__PUBLIC__EMPLOYEES" VALUES
(10001, 'Иванов Алексей Сергеевич', ...),
...;
```

Важно:
- **Имена БД/схемы/таблицы — английские** (требование SQL-транслятора three-part naming: `DB.SCHEMA.TABLE` → `schema_data."DB__SCHEMA__TABLE"`).
- **Значения колонок — русские** (ФИО, отделы, должности).
- **Идентификаторы в двойных кавычках** в DDL и INSERT — иначе Postgres lowercase'нет, и `SELECT "EMPLOYEE_ID"` начнёт мискаться с `employee_id`.

---

## 4. Регистрация MCP в `configs/mcp_servers/<name>.yaml`

```yaml
type: stdio
name: hr1c
params:
  command: uv
  args:
    - --directory
    - ${local_servers_paths}/hr1c-mcp-server
    - run
    - hr1c_mcp_server
    - --allow_write
    - --exclude-json-results
  env:
    PG_USER: eigent
    PG_PASSWORD: camel
    PG_HOST: cowork_pg
    PG_PORT: "5432"
    PG_DATABASE: cowork_gym
    PG_SCHEMA: hr1c_data
```

---

## 5. Mount миграции в Postgres-контейнеры

В **обоих** местах добавить mount `db/zzz_<name>_after_init.sql` → `/docker-entrypoint-initdb.d/`:

1. `docker-compose.yml` — для persistent `cowork_pg`.
2. `run_parallel.sh` — для эфемерных `pg-<port>-<task>` (per-task).

Иначе локальный `./cowork_bench run` работает, а `bench` — нет (или наоборот).

---

## 6. Перевести содержание таска

**`docs/task.md`** — переписать на русский с РФ-реалиями. Конкретные станции, поезда, даты, валюта.

**`docs/agent_system_prompt.md`** — scaffold уже подставил шаблон, дополнить под доменную специфику. **Никаких упоминаний `claim_done`** (см. [`memory/feedback_russification.md`](../memory/feedback_russification.md)).

**`task_config.json`** — `needed_mcp_servers` без лишнего (после scaffold проверить вручную). `needed_local_tools` — без `claim_done`.

**`initial_workspace/`** — перевести `.md`/`.txt`/`.json` файлы. Шаблоны `.xlsx` оставить с английскими заголовками колонок (eval сверяет посимвольно).

**`preprocess/main.py`** — переписать на новую MCP (`rzd`/`hr1c` вместо `rail_12306`/`snowflake`), русские email/gcal-стартовые данные.

**`evaluation/main.py`** — самое сложное. Если апстримный eval делает substring-проверки на английских ключах (`"Outbound"`, `"Total_Spent"`) — заменить на русские эквиваленты ИЛИ оставить английские заголовки в `groundtruth_workspace/`. Решение пер-таск.

**Гиды (`hr_query_guide.md` и т.п.)** — не давать слишком прямых подсказок (SQL/формулы/выбор поезда). Только бизнес-контекст + источники. См. [`memory/feedback_russification.md`](../memory/feedback_russification.md) про «не упрощать таск для модели».

**`groundtruth_workspace/`** — если eval строгий, сгенерировать руками через `_build_groundtruth.py` (паттерн см. в `rzd-hr1c-training-trip-kazan-strict-*`).

---

## 7. Журнал ревью

Добавить строку в `tasks_review.csv` (20 колонок, формат — см. [`memory/reference_tasks_review_csv.md`](../memory/reference_tasks_review_csv.md)). Найти позицию по индексу группы (`rzd-*` идут вместе).

---

## 8. Валидация

```bash
./cowork_bench test-db                                          # миграции применились?
./cowork_bench run <new-task>                                   # один прогон
./cowork_bench stability <new-task> 5                           # стабильность
```

Цель — **5/5 PASS** на стабильности. Если 4/5 — посмотреть `eval_res.json` упавшего прогона; обычно одна из:
- model wrote unquoted identifier в SQL → case-fold (известное ограничение);
- eval не нашёл подстроку → подправить eval или groundtruth;
- preprocess не очистил предыдущие email/gcal записи.

---

## Антипаттерны (из опыта)

- **Не патчить snowflake-MCP «параметризацией под русский».** Делать форк per-домен — snowflake грохнем целиком после миграции. Параметризация имеет смысл только когда форков ≥3.
- **Не оставлять `claim_done` в новых тасках.** В Strands семантика — `stop_reason="end_turn"` = SUCCESS. Упоминание `claim_done` сбивает модель.
- **Не упоминать «(как в Snowflake)»** в гидах — Snowflake скоро уйдёт.
- **`replace_all` в форке MCP** может создать `ff"..."` (двойной `f`-префикс). Перед коммитом — `grep -n 'ff"' src/`.
- **Забыть `pyyaml`** в `pyproject.toml` форка — рантайм-падение, не build-time.
- **Длинный `sleep` перед background-командой** — `tail -60` буфферится, вывод не виден. Использовать `> /tmp/run.log 2>&1` и `tail -f`.
