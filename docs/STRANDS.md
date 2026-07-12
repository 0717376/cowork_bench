# Strands-раннер

Альтернативный раннер бенча на [Strands Agents](https://github.com/strands-agents/sdk-python) + LiteLLM. Именно им получены solo-результаты из таблицы в README (single-thread прогоны). Отличия от дефолтного CAMEL-раннера (`main.py`):

- **Изоляция эталона (phase-aware).** Прогон разбит на две фазы: агент работает в санитизированном контейнере **без** `groundtruth_workspace/` и `evaluation/`, оценка выполняется отдельным контейнером. У CAMEL-раннера по умолчанию единый проход без изоляции.
- **Устойчивый стриминг.** Ретраи транзиентных ошибок (connection/5xx) до первого чанка, idle-timeout на каждый чанк (`LLM_STREAM_IDLE_TIMEOUT`), перехват context-overflow с автоподрезкой истории через `SummarizingConversationManager`.
- **Единый OpenAI-совместимый вход.** Любой эндпоинт с Chat Completions API; модель конфигурируется только env-переменными.

## Запуск

```bash
# .env
AGENT_ENTRY=main_strands.py
AGENT_PHASE_AWARE=1
LLM_BASE_URL=https://your-llm-endpoint/v1
LLM_API_KEY=sk-REPLACE_ME
LLM_MODEL=your-model-name
LLM_CONTEXT_WINDOW=131072
```

Дальше — те же команды, что и обычно:

```bash
./cowork_bench run <task-name>   # одна задача
./cowork_bench bench 1           # весь пул, single-thread (как в solo-прогонах)
```

Одну задачу можно запустить и напрямую, без обвязки:

```bash
python main_strands.py --task_dir <task-name> --max_steps 100 --phase all
```

`--phase agent` / `--phase eval` — раздельные фазы (так их вызывает `scripts/run_containerized.sh` при `AGENT_PHASE_AWARE=1`); `--phase all` — всё в одном процессе, удобно для отладки.

## Переменные окружения

| Переменная | Дефолт | Что делает |
|---|---|---|
| `LLM_BASE_URL` | — (обязательна) | OpenAI-совместимый эндпоинт |
| `LLM_API_KEY` | — (обязательна) | Ключ |
| `LLM_MODEL` | — (обязательна) | Id модели |
| `LLM_CONTEXT_WINDOW` | `262144` | Окно контекста; ставьте **реальный** лимит модели (лучше чуть ниже — история подрезается ДО отказа эндпоинта) |
| `LLM_MAX_TOKENS` | `32768` | Лимит генерации на ход |
| `LLM_PARAM_PROFILE` | `clean` | `clean` — чистый OpenAI-запрос (любой hosted API); `vllm` — добавляет vLLM-only параметры (top_k, min_p, repetition_penalty, chat_template_kwargs) для self-hosted эндпоинтов |
| `LLM_TEMPERATURE` | `1.0` (clean) / `0.6` (vllm) | Температура |
| `LLM_REASONING_EFFORT` | — | Прокидывается в `extra_body` (для моделей с управляемым reasoning) |
| `LLM_STREAM_IDLE_TIMEOUT` | `300` | Секунд тишины в SSE-стриме до ретрая/фейла |
| `LLM_SSL_VERIFY` | `true` | `false` — отключить проверку TLS (самоподписанные сертификаты) |
| `LLM_OR_PROVIDER` | — | OpenRouter provider routing (`provider.order`), например `minimax/fp8` |
| `LLM_OR_ALLOW_FALLBACKS` | — | `false` — запретить фолбэк на другие провайдеры OpenRouter |

## Зависимости

Docker-образ уже содержит всё нужное (`strands-agents`, `litellm` ставятся в `Dockerfile`). Для запуска вне Docker:

```bash
uv pip install "strands-agents==1.42.0" "litellm>=1.82.6" markdownify
```

(В `pyproject.toml` эти пакеты не объявлены намеренно: они пинят `mcp>=1.23`, что конфликтует с пинами task-runtime окружения.)

## Как воспроизвести solo-результаты из README

1. `LLM_CONTEXT_WINDOW` — реальное окно модели (для моделей с окном ~130k мы ставили 122–128k с запасом на дрейф токенизатора).
2. `./cowork_bench bench 1` — concurrency 1 исключает инфра-столлы, все фейлы принадлежат модели.
3. Агрегат — `benchmark_logs/fully_parallel_<timestamp>/summary.csv`; pass@1 = доля задач с `Pass: True`.
