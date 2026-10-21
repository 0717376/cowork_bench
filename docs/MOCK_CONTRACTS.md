# Local mock contracts and regression checks

These checks exercise the benchmark environment without calling a model. Run
from the repository root in the benchmark image, with the virtual environment
belonging to each server. Tests requiring PostgreSQL must point
`BENCH_TEST_PG_HOST` at a dedicated disposable database server. They create test
schemas in `cowork_query_test`; never use a retained benchmark run database.

## Memory

The `arxiv-lit-review-gsheet` grader also accepts the canonical MCP JSONL output; its reference graph uses `entityType` and `relationType`, matching the tools.

The memory MCP reads its canonical JSONL records and historical whole-graph JSON
fixtures, including pretty-printed and empty graphs. Reads do not rewrite files.
The next write uses canonical JSONL while retaining existing entities and
relations. Invalid JSON or foreign object shapes produce an error instead of
being silently replaced by an empty graph. Graders share the same two graph
formats through `utils.memory_graph`; the historical notes grader additionally
accepts its notes object.

```
python -m unittest runners.smoke.test_memory_graph_contract -v
python -m unittest runners.smoke.test_memory_mcp_roundtrip -v
```

The round-trip test starts the compiled memory MCP over stdio. Its default path
is `/opt/local_servers/servers/src/memory/dist/index.js`; set
`BENCH_MEMORY_SERVER_PATH` to test another build. The server's TypeScript tests and
build also run from `local_servers/servers/src/memory` with `npm test` and
`npm run build`.

## arXiv search

The PostgreSQL mock interprets `AND`, implicit AND, `OR`, `ANDNOT`, parentheses,
and `ti:`, `abs:`, `au:`, `all:` scopes. AND binds more tightly than OR. Quoted
phrases match contiguous literal text, ignoring case; ordinary terms use English
PostgreSQL full-text search. Query values are SQL parameters. Malformed syntax
is rejected rather than interpreted as a broader search. Limits are 4096
characters, 128 tokens and 12 nested groups/scopes.

Date/category filters apply to structured queries, exact IDs and the plain-text
substring fallback. Structured queries never use that fallback.

```
/opt/local_servers/arxiv-mcp-server/.venv/bin/python3 -m unittest runners.smoke.test_arxiv_query_contract -v
```

## SQL mocks

The Snowflake/ClickHouse compatibility layer uses PostgreSQL fixture tables.
Unquoted column names are resolved against the physical table catalog without
changing strings, comments, quoted identifiers or function names. Empty SELECTs
return an empty row list; successful statements without a result set return the
existing success marker. Unsupported metadata commands fail explicitly.

The layer translates the observed common ClickHouse count/date/conditional
aggregate functions and basic MergeTree table declarations. It is a limited
compatibility layer, not a complete implementation of either SQL dialect.
Unknown columns, missing tables and invalid joins remain errors.

```
/opt/local_servers/mcp-clickhouse-server/.venv/bin/python3 -m unittest runners.smoke.test_sql_mock_contract -v
```

## Terminal commands

Quoted or escaped shell punctuation is argument data. Validation and execution
use the same quote-aware operator splitting. Nested command substitution and
unsupported grouping/background/newline syntax are rejected. The configured
command allowlist remains authoritative; these fixes do not add executables.

```
python -m unittest runners.smoke.test_terminal_quoted_arguments -v
```

Historical scores must retain the source/image/configuration used to produce
them. Changing mocks, prompts, tools or graders requires a separate result set;
regression tests and SQL replays do not establish a new model score.
