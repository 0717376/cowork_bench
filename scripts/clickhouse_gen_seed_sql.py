"""Generate db/zzz_clickhouse_after_init.sql: idempotent UPDATEs that russify the
sf_data realia values (ClickHouse fork). Reads the actual distinct values from the
baked seed (db/init.sql.gz COPY blocks) so we map exactly what exists. Numbers,
IDs, column identifiers, KEEP-enums and PRODUCT_NAME/BRAND are untouched.

UPDATE ... WHERE col=<en> is idempotent (re-running is a no-op once russified).
Runs on next Postgres (re)build via /docker-entrypoint-initdb.d/ (zzz_ ordering).
"""
import gzip
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import clickhouse_relabel_map as M  # noqa: E402

PROJECT = os.path.dirname(HERE)
SEED = os.path.join(PROJECT, "db", "init.sql.gz")
OUT = os.path.join(PROJECT, "db", "zzz_clickhouse_after_init.sql")


def parse_copy_blocks(text):
    """Yield (table, [columns], [data_rows]) for each COPY sf_data block."""
    lines = text.split("\n")
    i, n = 0, len(lines)
    while i < n:
        ln = lines[i]
        if ln.startswith('COPY sf_data."') and "FROM stdin;" in ln:
            table = ln.split('"')[1]
            cols = [c.strip().strip('"') for c in ln[ln.index("(") + 1:ln.rindex(")")].split(",")] \
                if False else [c.strip().strip('"') for c in ln[ln.index("(") + 1:ln.rindex(")")].split(",")]
            rows = []
            i += 1
            while i < n and lines[i] != "\\.":
                rows.append(lines[i])
                i += 1
            yield table, cols, rows
        i += 1


def sql_str(s):
    return "'" + s.replace("'", "''") + "'"


def main():
    with gzip.open(SEED, "rt", encoding="utf-8", errors="replace") as f:
        text = f.read()

    # collect (table, column, set-of-distinct-en-values) for realia columns
    out = []
    out.append("-- ClickHouse fork: russify sf_data realia values (generated).")
    out.append("-- Idempotent: UPDATE ... WHERE col=<en>. Numbers/IDs/enums/PRODUCT_NAME kept English.")
    out.append("-- Source of distinct values: db/init.sql.gz COPY blocks. Map: scripts/clickhouse_relabel_map.py")
    out.append("SET client_encoding = 'UTF8';")
    out.append("")
    total = 0
    unmapped = []
    for table, cols, rows in parse_copy_blocks(text):
        realia = M.TABLE_REALIA.get(table)
        if not realia:
            continue
        idx = {c: k for k, c in enumerate(cols)}
        out.append(f'-- {table}')
        for col, mapname in realia.items():
            ci = idx[col]
            distinct = {}
            for r in rows:
                cells = r.split("\t")
                if ci >= len(cells):
                    continue
                v = cells[ci]
                if v == "\\N" or v == "":
                    continue
                if v in distinct:
                    continue
                ru = M.map_value(mapname, v)
                distinct[v] = ru
            for en, ru in sorted(distinct.items()):
                if ru is None:
                    unmapped.append((table, col, en))
                    continue
                if ru == en:
                    continue
                out.append(f'UPDATE sf_data."{table}" SET "{col}"={sql_str(ru)} WHERE "{col}"={sql_str(en)};')
                total += 1
        out.append("")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")
    print(f"wrote {OUT}: {total} UPDATE statements")
    if unmapped:
        print(f"UNMAPPED realia values ({len(unmapped)}) — review (expected only PRODUCT_NAME-like or new):")
        for t, c, v in unmapped[:40]:
            print(f"  {t}.{c} = {v!r}")
    else:
        print("all realia values mapped (no gaps)")


if __name__ == "__main__":
    main()
