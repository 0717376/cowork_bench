# Supplier critical-check prerequisites

In the supplier visit grader, removing Products priority/stock/city headers or
Travel_Plan city/price headers used to skip the declared CRITICAL comparisons.
A workbook with wrong priorities/prices could change from FAIL to PASS by losing
headers. Missing/unreadable workbooks and absent prerequisite sheets could also
be offset by unrelated points because the critical checks were never reached.

The patch makes file/readability and Products/Travel_Plan sheet presence critical.
Empty required sheets and missing enabling columns now explicitly fail the existing priority/price checks
instead of silently bypassing them. The 70% threshold and other soft criteria
are unchanged; Calendar and Email implementations are unchanged.

```sh
python -m unittest runners.smoke.test_supplier_required_structure -v
```

Tests execute the real Excel checker and main using synthetic real XLSX files;
calendar/email are isolated positive controls (3 and 4 checks). No DB, SMTP,
model or historical outcome is changed. A valid workbook remains accepted;
wrong priorities/prices, missing headers/sheets, missing and corrupt packages
are rejected even when partial accuracy is above 70%.
