# Canonical abstract length checks

The federated-learning task asks for character counts and their rounded mean.
The grader previously checked only that Abstract_Length existed and that the
reported mean exceeded 100; arbitrary or blank per-paper lengths could pass.

Two new mandatory checks run from main even if the ordinary Excel checker
returns early. They join workbook rows by Paper_ID to arxiv.papers.summary,
require each expected target ID once and verify every supplied ID resolves to
canonical text. Lengths equal Python len() of the decoded Unicode string, not
UTF-8 bytes or JSON escapes. The mean is computed from the canonical lengths
for those unique IDs, using Python round() (ties to even); the Summary must
contain exactly one Avg_Abstract_Length value. This rounding convention is
explicit so authors can review it.

Nonnegative integral numbers and numeric text are supported. Blank, negative,
boolean, fractional, nonfinite and incorrect counts fail. Missing columns,
sheets, required IDs, duplicate/unknown IDs or unavailable canonical data fail
closed; agent-provided counts cannot replace canonical data. A malformed mean
now produces a diagnostic FAIL rather than an unhandled float conversion error.

The existing 80% threshold and other partial checks remain. Two substantive
checks are added to the score denominator and to CRITICAL_CHECKS. Historical
raw outcomes and files are not rewritten. This patch is independent of the
already prepared Office schema/header gates; their combined source was also
tested. It does not duplicate that structural patch.

Seven offline tests exercise the real main and Office reader with a narrow
read-only DB seam, including numeric boundary values and ties-to-even rounding:

```sh
python -m unittest runners.smoke.test_abstract_length_contract -v
```

Optional real PostgreSQL integration with openpyxl and psycopg2-binary:

```sh
BENCH_ABSTRACT_TEST_PG_HOST=127.0.0.1 \
BENCH_ABSTRACT_TEST_PG_DATABASE=cowork_abstract_length_test \
BENCH_ABSTRACT_TEST_PG_USER=postgres \
python -m unittest runners.smoke.test_abstract_length_pg -v
```

Use an empty dedicated database named exactly cowork_abstract_length_test.
The suite creates/resets arxiv and teamly schemas and uses the public preprocess
seed, not a model reply. Port/password: BENCH_ABSTRACT_TEST_PG_PORT/PASSWORD.
Without HOST the suite is SKIP. Sixteen scenarios execute the full real main,
Teamly and Excel functions, using real XLSX files and real SQL; the connect seam
only redirects the hardcoded production DB settings to the guarded test DB.

Controls include canonical positive output, negative/blank/observed-wrong
values, off-by-one, wrong/self-consistent mean, duplicate/unknown IDs, missing
canonical data, actual SQL failure, Unicode versus UTF-8 length, numeric text
and reordered rows/columns. Workbook bytes remain unchanged. QA used an
existing Linux runtime with a source overlay, not a cold image build. No model
benchmark rerun, external arXiv request, retained-run DB or rescore is involved.
