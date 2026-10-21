# Sheets write ranges

The update tools accept a separate `sheet` and an **unqualified** A1 range.
For example, use `sheet="Sheet1", range="A3:D3"`. A qualified range supplied
there is rejected, even if its prefix matches the sheet. The tool quotes and
escapes the actual sheet name when building the backend request.

The PG backend also accepts a single qualified range directly. Previously
`Sheet1!Sheet1!A3:D3` produced undefined coordinates that were defaulted to row0,
col0. Invalid syntax now raises ValueError before acquiring a DB connection.
Intentional open row/column ranges retain their open coordinates. Whole-sheet
names can be quoted (`'Sheet1'`) to distinguish names from cell addresses.
The explicit backend whole-sheet form `Sheet1!` is preserved. Single cells and
quoted names with doubled apostrophes are supported.

Matrices must contain lists of rows with finite JSON scalar cells. Nested 3D
arrays, objects and malformed rows are rejected. All entries of values.batchUpdate
are checked before the first connection/write, so a later malformed range or
matrix cannot leave partial writes. This is input prevalidation, not a new
transaction guarantee for runtime database errors. Existing write/storage and
valueInputOption behavior otherwise remains unchanged.

With the Sheets server dependencies installed, run from the repository root:

```sh
python -m unittest discover -s local_servers/mcp-google-sheets/tests -v
```

Tests import the production parser, handlers and PG service. An in-memory storage
seam records writes; a real MCP server/client verifies isError and subsequent
recovery. No external Google API, PostgreSQL, model, or benchmark run is used.
The test verifies preserved headers/neighbors and valid/invalid controls; no
benchmark score change is claimed.
