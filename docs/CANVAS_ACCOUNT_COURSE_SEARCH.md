# Canvas account course search

The MCP schema and CanvasClient already forwarded search_term for account
courses, but PgCanvasRouter selected only by account_id. It now adds a bound
ILIKE condition on name OR course_code, under that same account restriction.
An absent/empty search preserves listing behavior. The pattern is %term%, as
in the existing account-user route; SQL LIKE wildcard characters retain their
pattern meaning. Account selection is never guessed or corrected by the router.
Other account-course flags and sorting are outside this patch.

Node >=22.13 is required for the offline real-source loader:

```sh
node --test runners/smoke/test_canvas_account_search.cjs
```

Four tests check the real TS router with a narrow pool seam. For actual SQL and
MCP stdio, install/build the server and use an empty dedicated PostgreSQL DB:

```sh
(cd local_servers/mcp-canvas-lms && npm ci && npm run build)
BENCH_MOCK_TEST_PG_HOST=127.0.0.1 \
BENCH_MOCK_TEST_PG_DATABASE=cowork_mock_contract_test \
BENCH_MOCK_TEST_PG_USER=postgres \
node --test runners/smoke/test_canvas_account_search_pg.cjs
```

The optional suite resets canvas in the dedicated DB named exactly
cowork_mock_contract_test. Port/password: BENCH_MOCK_TEST_PG_PORT/PASSWORD.
Without HOST it is SKIP. Eight cases use the real compiled server/client/router
and PostgreSQL: unfiltered/empty search, name/code case-insensitive search,
no match, account scoping, wrong account and parameterized hostile input.
BENCH_MOCK_TEST_SERVER_ROOT can point to an existing compiled server root.

The fix does not substitute the correct account_id for a model's incorrect one,
or establish that a previously failed task now passes. No benchmark/model
rerun or historical score change was performed. QA used compiled servers in an
existing Linux runtime and PostgreSQL15, not a cold image build.
