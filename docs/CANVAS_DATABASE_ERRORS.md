# Canvas PostgreSQL error propagation

PgCanvasRouter.get/post/put/delete previously logged DB failures and resolved
with empty data and status200/201. They now log and reject the original error.
The existing Canvas MCP tool handler turns that rejection into isError:true,
while successful empty reads and normal CRUD keep their existing responses.
No retry is added: CanvasClient selects the PG router before Axios interceptors
are installed; request() dispatch also propagates rejection.

Node >=22.13 offline router tests (five tests, four verbs plus normal controls):

```sh
node --test runners/smoke/test_canvas_error_channel.cjs
```

Actual SQL/MCP integration uses the compiled server and an empty dedicated DB:

```sh
(cd local_servers/mcp-canvas-lms && npm ci && npm run build)
BENCH_MOCK_TEST_PG_HOST=127.0.0.1 \
BENCH_MOCK_TEST_PG_DATABASE=cowork_mock_contract_test \
BENCH_MOCK_TEST_PG_USER=postgres \
node --test runners/smoke/test_canvas_error_channel_pg.cjs
```

The suite resets canvas only in the guarded DB cowork_mock_contract_test.
Port/password: BENCH_MOCK_TEST_PG_PORT/PASSWORD. Without HOST it is SKIP.
BENCH_MOCK_TEST_SERVER_ROOT optionally selects an existing compiled server root.
Seven cases cover successful empty read, persisted normal course CRUD, all four
MCP error channels under an actual missing-table error (42P01), retained-row
identity after failed calls, recovery and one router diagnostic per failed call.

The missing-table failure occurs before a write. This patch does not introduce
transactions or promise rollback for routes with multiple statements. The MCP
integration covers course CRUD; existing higher-level quiz recovery/catches
are outside this patch. Injected errors prove the mechanism, not how many
historical empty replies were DB failures. No model/API/retained-run DB is used,
and historical scores remain unchanged. QA used PostgreSQL15 and newly compiled
servers in an existing Linux runtime, not a cold image build.
