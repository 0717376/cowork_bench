# Required email evidence and partial-credit policy

Three travel graders used only accuracy >= 70%. Even with no matching email,
otherwise correct outputs passed (10/12, 9/11, 15/20 for the synthetic controls).
A messages-only fallback also counted unsent records; attachment references in
body text were not checked against the actual uploaded files.

The affected tasks are:

- rzd-canvas-fieldtrip-novgorod-gcal-word-email: Field_Trip_Notice.docx
- rzd-kulinar-team-trip-spb-catering-excel-gcal: Team_Trip_Plan.xlsx
- rzd-msk-spb-team-trip-gcal-word: Travel_Plan.docx

Policy proposed by this patch: retain the existing partial accuracy and 70%
threshold, and separately require a matching sent email and its named readable
Office attachment. The references to an attached document in the two Word task
statements are treated as requiring a real attachment, not only that sentence.
All other content checks retain their existing partial-credit behavior. PASS
still does not certify that every soft/content criterion succeeded.

`total_passed`, `total_checks` and `accuracy` keep their existing meaning. The
result JSON adds `partial_score` (same numeric accuracy), `mandatory_checks`,
`mandatory_pass` and `pass`. Exit0 now requires accuracy>=70 and both mandatory
flags; all failures and the numerical score remain visible.

The evidence must be in email.sent_log joined to a message addressed to the exact
recipient (case insensitive). This proves mock sending, not real SMTP delivery.
The required filename must be attached to one of those sent messages. Bytes are
read from email.attachments.content and parsed as DOCX/XLSX, including ZIP CRC;
a workspace file or content_path cannot substitute for the attachment. Standard
OOXML MIME and application/octet-stream (used by the PG backend) are accepted.
This checks package readability, not full attachment semantics or byte identity
to a possibly later-edited local artifact. Existing content checks remain as-is.

With python-docx, openpyxl and psycopg2-binary installed, from repository root:

```sh
python -m unittest runners.smoke.test_required_email_contract -v
```

13 tests cover all three real main functions and check_email implementations,
using generated Office files and a narrow service DB seam. Most scenarios run
all non-email checker functions; the low-score control explicitly stubs them.

Optional real SQL integration: create an empty dedicated PostgreSQL database
named exactly cowork_required_email_test. The suite creates/resets email/gcal/
canvas test schemas there. NEVER use a retained benchmark run database.

```sh
BENCH_EMAIL_TEST_PG_HOST=127.0.0.1 \
BENCH_EMAIL_TEST_PG_DATABASE=cowork_required_email_test \
BENCH_EMAIL_TEST_PG_USER=postgres \
python -m unittest runners.smoke.test_required_email_pg -v
```

Port and password use BENCH_EMAIL_TEST_PG_PORT / BENCH_EMAIL_TEST_PG_PASSWORD.
Without the host setting this integration is SKIP, not an integration success.
It exercises 30 scenarios with real PostgreSQL, all three main functions and
real non-email checks. No model or real SMTP is called. Historical scores and
artifacts are not rewritten. A full model benchmark must be reported separately.
