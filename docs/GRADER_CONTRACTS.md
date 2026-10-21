# Mandatory Office deliverables

These task statements require named Office deliverables. An absent, unreadable,
or structurally incomplete mandatory deliverable must fail a prerequisite; it
cannot be offset by partial credit in other components. Existing semantic
checks and task-specific 70%/80% thresholds remain unchanged.

Six Word graders call `require_readable_docx` immediately after argument parsing,
before database access. Controls include a genuine DOCX with unchanged bytes,
missing files, plain text, corrupt ZIPs and a ZIP without Word payload.

Spreadsheet checks require the headers and sheets that activate existing
semantic checks in `arxiv-federated-learning-ru-teamly-excel` and
`moex-financial-health-gsheet`. `sf-support-resolution-gsheet` requires its
mandatory readable local XLSX before database access. The inventory grader
reports missing required stock/schedule headers as critical failures instead
of raising `KeyError`.

Run from the repository root with `python-docx`, `openpyxl`, and
`psycopg2-binary` installed:

```sh
python -m unittest discover -s runners/smoke -p 'test_required_*.py' -v
```

The tests use temporary files and mocked database entrypoints. They do not
connect to a database, invoke a model or modify historical results. Passing
these structural controls does not certify the full semantic correctness of
every output or establish a new benchmark score.
