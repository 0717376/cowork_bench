# Product stock filter

`woo_products_list.stockStatus` is mapped to `stock_status` by ProductService.
PgRestRouter previously ignored it and returned products from other stock states.
The filter now adds a bound equality only for wc.products, before LIMIT/OFFSET.
Allowed values are instock/outofstock/onbackorder; invalid values reject before
any SQL. Undefined/null/empty string retain the unfiltered behavior. Unrelated
resources do not acquire a stock_status column. Other filters remain unchanged.

Node >=22.13 is required for the offline source loader:

```sh
node --test runners/smoke/test_woo_stock_filter.cjs
```

Seven behavioral tests run the real TS router with a narrow pool seam. They
check binding, resource scope, invalid input, filter combinations and empty/error
controls. This alone does not execute SQL or the service/MCP stack.

For the optional integration, install the existing server dependencies and build:

```sh
(cd local_servers/woocommerce-mcp && npm ci && npm run build)
BENCH_MOCK_TEST_PG_HOST=127.0.0.1 \
BENCH_MOCK_TEST_PG_DATABASE=cowork_mock_contract_test \
BENCH_MOCK_TEST_PG_USER=postgres \
node --test runners/smoke/test_woo_stock_filter_pg.cjs
```

Use an empty dedicated database named exactly cowork_mock_contract_test; the
suite resets its wc schema. Port/password are BENCH_MOCK_TEST_PG_PORT/PASSWORD.
Without HOST it is SKIP. Nine cases exercise the compiled stdio MCP server,
camelCase parameter mapping, all stock states, category/search combined with
pagination, invalid values and genuine emptiness. An unrelated orders control
uses the compiled router directly on the same real DB. No model or HTTP API is
called. Optional BENCH_MOCK_TEST_SERVER_ROOT selects an already-built server root.

These tests reproduce the filter violation. They do not prove a future PASS for
any full benchmark task or attribute every observed FAIL to this filter. The
historical series and raw scores are unchanged. QA used PostgreSQL15 and compiled
servers in an existing Linux runtime, not a cold image build.
