# forms-mcp

MCP-сервер для управления формами через PostgreSQL (RU-локализация / mock-замена google-forms-mcp).

Схема данных: `gform.*` (forms, questions, responses).

## Инструменты

- `create_form(title, description?)` — создать форму
- `add_text_question(formId, questionTitle, required?)` — текстовый вопрос
- `add_multiple_choice_question(formId, questionTitle, options, required?)` — вопрос с выбором
- `get_form(formId)` — получить форму со списком вопросов
- `get_form_responses(formId)` — список ответов на форму

## ENV

`PG_HOST`, `PG_PORT`, `PG_DATABASE`, `PG_USER`, `PG_PASSWORD`. Значения по умолчанию подходят для cowork_pg (localhost/15433/cowork_gym/postgres/postgres).

## Сборка

```
npm install
npm run build
```
