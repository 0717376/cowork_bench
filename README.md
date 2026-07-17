# Cowork Bench

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](./LICENSE)

**EN:** Cowork Bench is a multi-tool agent benchmark: **496 end-to-end tasks** that mirror real office workflows — pull data from a mock database, build Excel/Word/PPTX artifacts, schedule calendar events, send emails — executed through **MCP (Model Context Protocol) servers** (~120 tools visible per task). Evaluation is **fully deterministic** (programmatic checks + SQL side-effect verification, no LLM-as-judge) and runs entirely locally: one Docker Compose, a seeded PostgreSQL, zero external API calls at eval time. All task statements and system prompts are in **Russian** (identifiers stay Latin for exact matching), with Russia-specific domains (railways, MosBirzha, 1C HR, e-commerce). Requirements: Docker + bash; see the Quick Start below (in Russian). Built on top of [Toolathlon](https://github.com/hkust-nlp/Toolathlon) (Apache 2.0).

---

— мультитул-бенчмарк для LLM-агентов: задачи реальных рабочих процессов, локальная PostgreSQL, без внешних API на этапе прогона.

Обучать и оценивать LLM-агентов на реальном использовании инструментов сложно: существующие датасеты либо узки по покрытию инструментов, либо малы по масштабу, либо зависят от живых внешних API, которые меняются со временем. **Cowork Bench** — большой самодостаточный набор задач с богатой мок-базой, который целиком работает локально: на этапе оценки не требуется ни одного внешнего вызова. Все задания и системные промпты — на **русском языке**; идентификаторы (имена файлов, листов, колонок, email, тикеры, статус-слаги) сохранены на латинице для детерминированной проверки.

Каждая задача просит агента выполнить сквозную цель — например, вытащить данные из мок-базы, собрать Excel-отчёт, поставить событие в календарь и отправить письмо — пользуясь фиксированным набором MCP-серверов (Model Context Protocol) как инструментами.

Каждая задача полностью автоматизирована: `preprocess/main.py` готовит исходное состояние воркспейса, агент исполняет задачу доступными инструментами, а `evaluation/main.py` сверяет результат с эталоном (`groundtruth_workspace/`). Ни ручной разметки, ни живых внешних сервисов.

Бенчмарк нагружает способности, важные на практике: многошаговое планирование по разнородным инструментам, чтение и запись структурированных форматов, синхронизацию данных между системами и доведение длинных задач до конца под фиксированным бюджетом шагов.

## Раннер

В репозитории два раннера:

- **CAMEL** (`main.py`) — дефолт: единый проход агент+оценка, минимальная обвязка.
- **Strands** (`main_strands.py`) — phase-aware раннер с изоляцией эталона (агент не видит `groundtruth_workspace/` и `evaluation/`), устойчивым стримингом и ретраями; именно им получены solo-результаты из таблицы ниже. Включение: `AGENT_ENTRY=main_strands.py`, `AGENT_PHASE_AWARE=1` + переменные `LLM_*` — см. [`docs/STRANDS.md`](./docs/STRANDS.md).

Дефолтный раннер — **CAMEL-агент** (`main.py`). Завершение задачи определяется по тому, что модель закончила ход (перестала вызывать инструменты) — **служебный инструмент `claim_done` не требуется**: статус `SUCCESS` ставится, когда агент завершил ход без новых tool-call'ов (см. `utils/roles/task_agent.py`). Оценка делается эвалуатором задачи по сайд-эффектам, а не по факту вызова сигнального тула.

Модель и провайдер для CAMEL-раннера задаются (в порядке приоритета):
1. переменными окружения `MODEL_NAME` и `MODEL_PROVIDER` (их читает `main.py`);
2. аргументами `--model_name` / `--provider`;
3. файлом `scripts/eval_config.json` (fallback).

Ключ и эндпоинт берутся из `MODEL_API_KEY` / `MODEL_API_URL`. Допустимые значения `MODEL_PROVIDER`: `openai`, `anthropic`, `gemini`, `deepseek`, `aihubmix`, `openai_compatible`.

## Quick Start

Требуется только **Docker** и **bash**. Всё остальное (Python, MCP-серверы, миграции) — внутри контейнеров.

```bash
cp .env.example .env                  # вписать MODEL_* (см. ниже)
docker build -t cowork-pack:latest .  # собрать образ (один раз, ~10–15 мин)

./cowork_bench up                     # поднять Postgres (cowork_pg) + сеть cowork_net
./cowork_bench test-db                # smoke миграций db/* на свежем PG
./cowork_bench run <task-name>        # один прогон одной задачи (агент + оценка)
./cowork_bench stability <task> 5     # 5 прогонов подряд, сводка PASS/FAIL
./cowork_bench bench 3                # параллельный прогон всех задач (concurrency=3)
./cowork_bench tasks [pattern]        # список задач (опционально по подстроке)
./cowork_bench stop                   # аварийная зачистка контейнеров + lock
./cowork_bench help                   # полный список команд
```

`./cowork_bench up` запускает `docker compose up -d` (Postgres `cowork_pg`, сеть `cowork_net` и опциональный dev-shell-контейнер `cowork`). Для одиночного прогона задачи (`run`) достаточно, чтобы `cowork_pg` был healthy, а сеть `cowork_net` существовала — `scripts/run_containerized.sh` сам поднимает эфемерные контейнеры агента под каждую задачу. `cowork_bench` source-ит `.env` и по умолчанию использует CAMEL-раннер (`AGENT_FRAMEWORK=camel`, `IMAGE=cowork-pack:latest`); любой дефолт переопределяется через `.env`/env. Грабли окружения — в [`docs/ENVIRONMENT_QUIRKS.md`](./docs/ENVIRONMENT_QUIRKS.md).

### Минимальный `.env` для CAMEL-раннера

```bash
MODEL_PROVIDER=openai_compatible          # openai | anthropic | gemini | deepseek | aihubmix | openai_compatible
MODEL_NAME=your-model-name
MODEL_API_KEY=sk-REPLACE_ME
MODEL_API_URL=https://your-llm-endpoint/v1   # нужен для openai_compatible / aihubmix
```

### Оценка результата

С дефолтным CAMEL-раннером оценка выполняется **автоматически в том же прогоне**: `main.py` после завершения агента сам вызывает эвалуатор и пишет вердикт в `eval_res.json`. Отдельную команду запускать не нужно.

Отдельный standalone-грейдер `scripts/run_eval.py` (и отдельный eval-контейнер в `run_containerized.sh`) задействуется только для phase-aware раннеров с изоляцией эталона (см. «Подключение своего раннера»).

### Где смотреть результат

```
dumps/<task>/<timestamp>/.../SingleUserTurn-<task>/
├── traj_log.json     # config + status прогона
├── traj.json         # полная траектория (messages + tool_calls)
├── eval_res.json     # {"pass": true|false, "details": "..."}
└── workspace/        # файлы, которые сделал агент
```

Агрегат параллельного прогона — `benchmark_logs/fully_parallel_<timestamp>/summary.csv`.

## Подключение своего раннера

`scripts/run_containerized.sh` выбирает entrypoint агента так:
- `AGENT_ENTRY=<path>` — явный путь к вашему скрипту-раннеру (высший приоритет);
- иначе `AGENT_FRAMEWORK=strands` → `main_strands.py`;
- иначе дефолт → `main.py` (CAMEL).

Контракт движка:
- принимает `--task_dir <task>` и `--max_steps <N>`;
- пишет `traj_log.json` и каталог `workspace/` под `/workspace/dumps` — оттуда их находит грейдер `scripts/run_eval.py`.

**Изоляция эталона (опционально).** Чтобы агент гарантированно не видел `groundtruth_workspace/` и `evaluation/`, прогон разбивается на две фазы (агент в санитизированном контейнере → оценка отдельно). Это включается, если раннер — `main_strands.py`, либо если выставлено `AGENT_PHASE_AWARE=1`; тогда агент должен поддерживать флаг `--phase agent`. Без этого выполняется единый проход (агент + оценка вместе, полный mount, **без** изоляции).

```bash
# свой раннер с изоляцией эталона
AGENT_ENTRY=my_runner.py AGENT_PHASE_AWARE=1 \
  bash scripts/run_containerized.sh <task-name> 100
```

Код раннера и грейдера запекается в образ (`Dockerfile: COPY . .`), поэтому собранный из этого репозитория образ работает без host-mount'ов. Для быстрой итерации над кодом раннера — `DEV_MOUNTS=1` (оверлей host-кода).

---

## Архитектура: MCP-серверы

`configs/mcp_servers/*.yaml` регистрирует MCP-сервер, `local_servers/<name>/` содержит его код. Бэкенд для «SQL-подобных» MCP — общий PostgreSQL `cowork_pg`, схема per-MCP (например `canvas`, `moex`, `wc`, `hr1c_data`).

Данные грузятся детерминированно: **`db/init.sql.gz`** — полный дамп (все схемы, включая русские данные `rzd`, `teamly`, `moex`, `hr1c`, `gform`, релейблинг `canvas`, `wc`), автоматически применяется docker-entrypoint'ом Postgres при первом старте.

Добавить новый домен → форкнуть MCP в `local_servers/`, написать post-init миграцию `db/zzz_<name>_after_init.sql` (префикс `zzz_` гарантирует загрузку **после** `init.sql.gz` — entrypoint сортирует файлы по алфавиту), зарегистрировать её в `docker-compose.yml` (volume в `/docker-entrypoint-initdb.d/`) и сервер в `configs/mcp_servers/<name>.yaml`. Готовый рецепт — [`docs/RUSSIFICATION.md`](./docs/RUSSIFICATION.md).

---

## Структура задачи

Все задачи лежат в `tasks/finalpool/`. Каждая директория задачи имеет единый layout:

```
<task-name>/
├── task_config.json         # какие MCP-серверы и локальные тулы доступны агенту
├── docs/
│   ├── task.md              # описание задачи (показывается агенту)
│   └── agent_system_prompt.md
├── evaluation/main.py       # автоматический эвалуатор
├── preprocess/main.py       # подготовка состояния БД (запускается перед задачей)
├── initial_workspace/       # входные файлы в воркспейсе агента
└── groundtruth_workspace/   # эталонные выходы для проверки
```

Описания задач (`task.md`) написаны **без брендовых имён инструментов** — сервисы вроде базы знаний, общего календаря или LMS описаны обобщённо. Это та же конвенция обфускации, что и в исходном проекте Toolathlon: она мешает агенту срезать углы по узнаванию ключевых слов и заставляет рассуждать о реальном использовании инструментов.

## Мок-база

**Подключение**: БД `cowork_gym` на контейнере `cowork_pg` (внутри сети — `cowork_pg:5432`, с хоста — `localhost:15433`; пользователь `eigent`, пароль `camel`).

Все данные отдаёт локальный PostgreSQL, инициализируемый из сжатого дампа (`db/init.sql.gz`) плюс post-init миграции. Внешних вызовов в рантайме нет — окружение полностью контролируемо, без лимитов API, смены схем или дрейфа данных.

Данные выведены или симулированы по реальным источникам: **Kaggle OULAD** (Open University Learning Analytics) для LMS, **Kaggle HR Analytics** для корпоративного HR, биржевые ряды для финансов, комбинация **Kaggle Amazon** и **DummyJSON** для e-commerce, плюс русские форки (РЖД, база знаний, формы).

### Богатые данными схемы

| Домен | Описание | Схема |
|-------|----------|-------|
| **canvas** | LMS — курсы, пользователи, записи, задания, сдачи, тесты, рубрики, объявления (русифицированные имена/реалии) | `canvas` |
| **moex** | Биржа — котировки, финансовая отчётность, новости, опционы, держатели | `moex` |
| **woocommerce/insales** | E-commerce — товары, заказы, клиенты, купоны, отзывы | `wc` |
| **rzd** | Ж/д — станции, поезда, маршруты, места | `rzd` |
| **teamly** | База знаний (российский аналог Confluence) | `teamly` |
| **forms** | Опросы/формы на PG-бэкенде | `gform` |
| **hr1c** | Корпоративный HR | `hr1c_data` |
| **youtube** | Видеоплатформа — каналы, плейлисты, видео, транскрипты | `youtube` |

## Результаты

Прогоны июнь–июль 2026, все 496 задач, метрика **pass@1**. Метод **solo** — single-thread прогоны Strands-раннером (concurrency 1): чистый pass@1 при нуле инфра-столлов. Метод **parallel** — более ранние многопоточные прогоны с поправкой ex-stall (см. сноски).

| Модель | Окно | PASS@1 | Метод |
|---|---|---|---|
| **MiniMax-M3** (428B-A23B, FP8) | 400k | 323 / 496 — **65.1%** | parallel |
| **Qwen 3.5 397B-A17B** (BF16, MoE) | 262k | 312 / 496 — **62.9%** | solo |
| **Qwen 3.6 27B** (Dense, FP8) | 262k | 309 / 496 — **62.3%** | solo |
| **Qwen 3.6 35B-A3B** (FP8, MoE) | 262k | 300 / 496 — **60.5%** | solo |
| **Qwen 3.5 122B-A10B** (MoE) | 262k | 265 / 496 — **53.4%** | parallel |
| **Gemma 4 31B** (Dense, BF16) | 262k | 257 / 496 — **51.8%** | parallel + solo |
| **GigaCode Agentic v1.6** \* (80B-A3B, FP8) | 262k | 120 / 421 — **28.5%** | parallel, ex-stall |
| **GigaChat-3.5** (432B-A28B) \*\*\* | 130k | 119 / 496 — **24.0%** | solo |
| **GigaChat-3-Max** \*\* (90B Dense, FP8) | 130k | 41 / 375 — **10.9%** | parallel, ex-stall |
| **GigaChat-3-Pro** \*\*\* | 130k | 24 / 496 — **4.8%** | solo |

> \* **GigaCode Agentic v1.6** — исключены инфра-фейлы: 75 stream-stall. Метрика **28.5%** = 120/421 дошедших до эвала (сырой pass@1 = 120/496 = 24.2%).
>
> \*\* **GigaChat-3-Max** — исключены инфра-фейлы: 59 stream-stall + 22 × HTTP 500 + 32 task_timeout. Метрика **10.9%** = 41/375 дошедших до эвала (сырой pass@1 = 41/496 = 8.3%).
>
> \*\*\* **GigaChat-3.5 / 3-Pro** — прогнаны через legacy `functions`-API с транслирующим прокси (**1 инструмент за ход** — структурный гандикап); реальное окно ~130k. 3-Pro — без reasoning.

Бенч даёт большой разброс по силе моделей (65% → 5%); внутри семейства Qwen влияют и размерность, и поколение: 3.6-модели 27B/35B держат уровень флагмана 3.5-397B (60–63%), тогда как 3.5-122B — 53.4%.

## Что отличает Cowork Bench

- **Масштаб и разнообразие.** Сотни задач по десяткам MCP-серверов и нескольким доменам данных; задачи требуют настоящей кросс-системной координации, а не одиночных lookup'ов.
- **Полностью локально и воспроизводимо.** Всё окружение поднимается из одного Docker Compose. Ключи к внешним сервисам на этапе оценки не нужны. Дамп PostgreSQL версионируется и детерминирован — результаты воспроизводимы между машинами и во времени.
- **Реалистичная сложность.** Задачи списаны с реальных рабочих процессов: HR-аналитика в Excel, сверка сдач LMS с дедлайнами в календаре, генерация презентаций из собранных данных и подобные многошаговые цели. Большинству задач нужно 4–7 инструментов.

## Acknowledgements

Cowork Bench построен на инфраструктуре и исходных пайплайнах данных проекта:

> **Toolathlon: Benchmarking LLM Agents on Real-World Tool-Use Tasks**
> HKUST-NLP — https://github.com/hkust-nlp/Toolathlon

Дизайн схемы мок-базы, интерфейсы MCP-серверов и фреймворк оценки происходят из проекта Toolathlon (Apache License 2.0, см. [`LICENSE`](./LICENSE)). Cowork Bench локализует пул задач на русский язык, заменяет ряд внешних сервисов на российские форки и расширяет набор данных.

Каталог `local_servers/` содержит vendored-форки сторонних open-source MCP-серверов (filesystem, excel, word, powerpoint, playwright, fetch и др.) — авторство и лицензии указаны в их собственных манифестах внутри соответствующих подкаталогов.
