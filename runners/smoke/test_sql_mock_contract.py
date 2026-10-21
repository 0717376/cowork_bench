"""Catalog-backed SQL names, empty SELECT and literal preservation on real PG."""

import asyncio
import importlib.util
import os
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location(
    "cowork_db",
    ROOT / "local_servers/mcp-clickhouse-server/src/mcp_snowflake_server/db_client.py",
)


@unittest.skipUnless(
    os.environ.get("BENCH_TEST_PG_HOST"), "dedicated disposable PostgreSQL required"
)
class SQLMockContract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import psycopg2

        cls.module = importlib.util.module_from_spec(SPEC)
        SPEC.loader.exec_module(cls.module)
        cls.old = {
            k: os.environ.get(k)
            for k in ["PG_HOST", "PG_USER", "PG_PASSWORD", "PG_DATABASE"]
        }
        os.environ.update(
            PG_HOST=os.environ["BENCH_TEST_PG_HOST"],
            PG_USER="eigent",
            PG_PASSWORD="camel",
            PG_DATABASE="cowork_query_test",
        )
        cls.connection = psycopg2.connect(
            host=os.environ["PG_HOST"],
            user="eigent",
            password="camel",
            dbname="cowork_query_test",
        )
        cls.connection.autocommit = True
        with cls.connection.cursor() as cur:
            cur.execute("CREATE SCHEMA sf_data")
            cur.execute(
                'CREATE TABLE sf_data."SALES_DW__PUBLIC__CUSTOMERS" ("CUSTOMER_ID" int,"NAME" text,"MixedCase" text,"SEGMENT" text)'
            )
            cur.execute(
                "INSERT INTO sf_data.\"SALES_DW__PUBLIC__CUSTOMERS\" VALUES (1,'Alice','exact','regular')"
            )
        cls.db = cls.module.SnowflakeDB({})

    @classmethod
    def tearDownClass(cls):
        if cls.db._conn:
            cls.db._conn.close()
        cls.connection.close()
        for k, v in cls.old.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

    def query(self, query):
        return asyncio.run(self.db.execute_query(query))[0]

    def test_unquoted_column_casing_and_quoted_aliases(self):
        expected = [{"ID": 1, "NAME": "Alice", "MIXEDCASE": "exact"}]
        self.assertEqual(
            self.query(
                'SELECT c.CUSTOMER_ID AS id,c.NAME,c."MixedCase" FROM SALES_DW.PUBLIC.CUSTOMERS c'
            ),
            expected,
        )
        self.assertEqual(
            self.query(
                'select c.customer_id as id,c.name,c."MixedCase" from SALES_DW.PUBLIC.CUSTOMERS c'
            ),
            expected,
        )

    def test_empty_read_is_empty_and_dml_has_success_marker(self):
        # Isolate the empty-result defect from identifier case translation.
        self.assertEqual(self.query("SELECT 1 AS n WHERE FALSE"), [])
        self.assertEqual(
            self.query("SELECT * FROM SALES_DW.PUBLIC.CUSTOMERS WHERE CUSTOMER_ID=999"),
            [],
        )
        data = self.query(
            "UPDATE SALES_DW.PUBLIC.CUSTOMERS SET NAME='Alice' WHERE CUSTOMER_ID=999"
        )
        self.assertEqual(data[0]["status"], "success")
        self.assertEqual(
            self.query(
                "SELECT table_name FROM SALES_DW.INFORMATION_SCHEMA.TABLES WHERE table_schema='MISSING'"
            ),
            [],
        )

    def test_literals_comments_and_explicit_case_are_not_rewritten(self):
        # No physical table/column is needed to observe literal corruption.
        self.assertEqual(
            self.query("SELECT 'SALES_DW.PUBLIC.CUSTOMERS' AS note"),
            [{"NOTE": "SALES_DW.PUBLIC.CUSTOMERS"}],
        )
        q = "SELECT 'SALES_DW.PUBLIC.CUSTOMERS' AS note, name FROM SALES_DW.PUBLIC.CUSTOMERS -- DB.PUBLIC.OTHER\n WHERE name='Alice'"
        self.assertEqual(
            self.query(q), [{"NOTE": "SALES_DW.PUBLIC.CUSTOMERS", "NAME": "Alice"}]
        )
        self.assertEqual(
            self.query('SELECT "MixedCase" FROM SALES_DW.PUBLIC.CUSTOMERS'),
            [{"MIXEDCASE": "exact"}],
        )

    def test_clickhouse_date_count_and_basic_create_table(self):
        self.assertEqual(
            self.query("SELECT toYYYYMM(DATE '2026-03-01') AS month"),
            [{"MONTH": 202603}],
        )
        self.assertEqual(
            self.query("SELECT count() AS n FROM SALES_DW.PUBLIC.CUSTOMERS"), [{"N": 1}]
        )
        self.query(
            "CREATE TABLE TEST.PUBLIC.ITEMS (ID UInt32, LABEL String, PRICE Float64) ENGINE = MergeTree ORDER BY ID"
        )
        self.query(
            "INSERT INTO TEST.PUBLIC.ITEMS (id,label,price) VALUES (1,'String DB.PUBLIC.ITEMS',2.5)"
        )
        self.assertEqual(
            self.query("SELECT id,label,price FROM TEST.PUBLIC.ITEMS"),
            [{"ID": 1, "LABEL": "String DB.PUBLIC.ITEMS", "PRICE": 2.5}],
        )

    def test_keyword_column_and_clickhouse_functions(self):
        self.assertEqual(
            self.query(
                "WITH c AS (SELECT SEGMENT FROM SALES_DW.PUBLIC.CUSTOMERS) SELECT SEGMENT FROM c"
            ),
            [{"SEGMENT": "regular"}],
        )
        self.assertEqual(
            self.query(
                "SELECT countIf(CUSTOMER_ID>0) AS n FROM SALES_DW.PUBLIC.CUSTOMERS"
            ),
            [{"N": 1}],
        )
        self.assertEqual(
            self.query(
                "SELECT toYear(DATE '2026-03-02') AS y,toMonth(DATE '2026-03-02') AS m,toString(DATE '2026-03-02') AS s,toStartOfQuarter(DATE '2026-03-02') AS q"
            ),
            [
                {
                    "Y": 2026,
                    "M": 3,
                    "S": "2026-03-02",
                    "Q": __import__("datetime").date(2026, 1, 1),
                }
            ],
        )
        self.assertEqual(
            self.query(
                "SELECT avgIf(CUSTOMER_ID,CUSTOMER_ID>0) AS average FROM SALES_DW.PUBLIC.CUSTOMERS"
            ),
            [{"AVERAGE": __import__("decimal").Decimal(1)}],
        )
        self.assertEqual(
            self.query(
                "SELECT sumIf(CUSTOMER_ID,CUSTOMER_ID>0) AS total FROM SALES_DW.PUBLIC.CUSTOMERS"
            ),
            [{"TOTAL": 1}],
        )
        self.assertEqual(self.query("SHOW TABLES FROM MISSING.PUBLIC"), [])
        self.assertTrue(self.query("SHOW TABLES FROM SALES_DW.PUBLIC"))

    def test_unknown_column_still_fails_instead_of_guessing_data(self):
        with self.assertRaises(Exception):
            self.query("SELECT invented_column FROM SALES_DW.PUBLIC.CUSTOMERS")
        self.db._conn.rollback()


if __name__ == "__main__":
    unittest.main()
