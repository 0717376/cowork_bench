# Environment Quirks

Известные особенности окружения Cowork-Bench, которые легко превращаются в потерянное время, если про них не знать. Дополняйте по мере находок — список растёт.

## Postgres

### init-файлы сортируются по `en_US.utf8`
Внутри `postgres:15` контейнера `LANG=en_US.utf8`. В этой локали символ `_` (0x5F) сортируется **раньше** `.` (0x2E). Поэтому glob `/docker-entrypoint-initdb.d/*` развернётся в порядке `init_rzd.sql, init.sql.gz` — то есть кастомный `init_*` будет применён **до** оригинального dump.

**Правило:** любой дополнительный SQL в `db/` называть `zzz_*` (или с префиксом `99_`+) — это гарантирует выполнение **после** `init.sql.gz` в любой локали.

`ON_ERROR_STOP=1` оборвёт init на первой ошибке, и healthcheck не выйдет в healthy — это то, что валит все 504 задачи в `run_parallel.sh` с `pg_fail`.

Перед любым изменением `db/` запускать `scripts/test_db_migration.sh` — проверяет, что свежий PG выходит в healthy ≤60s.

### Два разных Postgres в проекте
1. **Persistent** `cowork_pg` (docker-compose.yml, порт `15433:5432`, persistent volume). Используется `scripts/run_containerized.sh` и для отладки. Накапливает мутации между прогонами — то, что «работает здесь», может ломать свежий PG.
2. **Per-task свежий** `pg-${TASK_ID}` (`run_parallel.sh`). На каждую задачу — свой контейнер, своя сеть, init с нуля.

В preprocess/evaluator всегда: `port=int(os.environ.get("PGPORT", "5432"))` — поддерживает оба сценария (внутри docker network — 5432, с хоста — 15433).

## MCP-серверы

### Notion — placeholder-токен, не использовать
`configs/mcp_servers/notion.yaml` содержит `Bearer ntn-placeholder`. Реального API-токена нет. Любая операция возвращает `401 Unauthorized / "API token is invalid."`.

**Не добавлять** `notion` в `needed_mcp_servers` новых задач. Если в исходной таске нужна KB-страница — заменять на:
- word-документ (надёжно работает)
- дополнительный markdown в workspace через filesystem MCP
- email с отчётом

В evaluator не делать `DELETE FROM notion.pages` в preprocess — может задеть параллельные тесты, если они с notion взаимодействуют.

### Kulinar — русская кулинарная база
`local_servers/kulinar-mcp` — MCP с русскими рецептами (50 блюд классики РФ-кухни: салаты, закуски, супы, горячее, гарниры, выпечка, десерты, напитки). Источник правды — `scripts/generate_recipes.py`. Тулы: `getAllRecipes`, `getRecipesByCategory`, `getRecipeById`, `recommendMeals`, `whatToEat`.

Для русских тасок описывать кулинарные сценарии напрямую (меню, банкеты, закупки продуктов) — все названия блюд русские.

## Evaluator

### `normalize()` ≠ для русских слов
В evaluator русских задач типичная функция `normalize()` делает `NFKD + translit` (`а→a, о→o, е→e, р→p, с→c, у→y, к→k, х→x`).

Использовать её **только** для смешанных кириллица/латиница ID (`'752А' vs '752A'`). Русские keyword-слова искать в `lower()` оригинале — иначе матч провалится (в `norm` 'отправ' превращается в 'otpaв', и `"отправ" in norm` всегда False).

Шаблон:
```python
text_low = text.lower()
text_norm = normalize(text)
has_train_no = "752" in text_norm           # ID-матчинг — norm
has_ru_keyword = "отправ" in text_low       # русское слово — lower
```

## Tasks Review CSV

`tasks_review.csv` — 505 строк × 20 колонок. Шапка:
```
task,category,needed_mcps,task_md_words,eval_loc,eval_funcs_defined,eval_funcs_called,eval_unused_funcs,eval_has_chinese,gt_files,has_preprocess,last_result_status,last_result_pass,last_result_duration_s,eval_bug,task_quality,ru_effort_min,priority,status,notes
```

- `needed_mcps` — `;`-разделитель, не запятая.
- `notes` — последнее поле, может содержать запятые/кириллицу. Не заворачивать в лишние кавычки.
- После правки сразу валидировать shape:
  ```bash
  python3 -c "import csv; r=list(csv.reader(open('tasks_review.csv'))); print({k:len(v) for k,v in [(len(row),i) for i,row in enumerate(r)]})"
  ```
  Должно дать `{20: 505}`.
- Русифицированные строки (`category=rzd`) держать рядом, не разбрасывать.

## Русификация задач

При создании `tasks/finalpool/rzd-*` всё — на русском, включая `docs/agent_system_prompt.md`. Не оставлять английский шаблон.

Готовый русский шаблон system-prompt — в `tasks/finalpool/rzd-canvas-fieldtrip-novgorod-gcal-word-email/docs/agent_system_prompt.md`. Скелет для новой rzd-таски разворачивается через `scripts/scaffold_ru_task.py`.
