"""The advertised arXiv query language must match actual PostgreSQL results.

Pure parser tests use stdlib. Integration tests use only a dedicated disposable
PostgreSQL (BENCH_TEST_PG_HOST). Never point them at a benchmark run database.
"""

import asyncio
import importlib.util
import json
import os
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "local_servers/arxiv-mcp-server/src"
spec = importlib.util.spec_from_file_location(
    "cowork_query_syntax", SRC / "arxiv_mcp_server/query_syntax.py"
)
syntax = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = syntax
spec.loader.exec_module(syntax)


class QueryValidation(unittest.TestCase):
    def test_boundaries_and_malformed_queries_fail_without_sql(self):
        for query in [
            '"unclosed',
            "()",
            "x OR",
            "AND x",
            "ti:",
            "x)",
            "(x",
            "bad:x",
            '""',
            "x " * 129,
            "(" * 14 + "x" + ")" * 14,
        ]:
            with self.subTest(query=query), self.assertRaises(ValueError):
                syntax.compile_query(query)

    def test_untrusted_literals_are_only_parameters(self):
        query = 'ti:"Robert\'s research; DROP TABLE papers"'
        compiled = syntax.compile_query(query)
        self.assertNotIn("DROP", compiled.predicate)
        self.assertIn("Robert's research; DROP TABLE papers", compiled.params)
        self.assertIsNone(compiled.plain_text)


@unittest.skipUnless(
    os.environ.get("BENCH_TEST_PG_HOST"), "dedicated PostgreSQL required"
)
class PostgreSQLSearch(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import psycopg2

        cls.old_env = {
            k: os.environ.get(k)
            for k in ["PG_HOST", "PG_PORT", "PG_USER", "PG_PASSWORD", "PG_DATABASE"]
        }
        os.environ.update(
            PG_HOST=os.environ["BENCH_TEST_PG_HOST"],
            PG_PORT="5432",
            PG_USER="eigent",
            PG_PASSWORD="camel",
            PG_DATABASE="cowork_query_test",
        )
        cls.conn = psycopg2.connect(
            host=os.environ["PG_HOST"],
            user="eigent",
            password="camel",
            dbname="cowork_query_test",
        )
        cls.conn.autocommit = True
        with cls.conn.cursor() as c:
            c.execute("CREATE SCHEMA arxiv")
            c.execute("""CREATE TABLE arxiv.papers (
                id text PRIMARY KEY,title text,authors jsonb,summary text,categories jsonb,
                primary_category text,published timestamptz,doi text,journal_ref text,
                comment text,pdf_url text,links jsonb)""")
            samples = [
                (
                    "1",
                    "Neural networks for images",
                    '["Alice Smith"]',
                    "A graph survey",
                    '["cs.CV"]',
                    "2024-01-01",
                ),
                (
                    "2",
                    "Graph learning for language",
                    '["Bob Jones"]',
                    "Neural networks and attention",
                    '["cs.CL"]',
                    "2023-01-01",
                ),
                (
                    "3",
                    "Neural efficient scalable networks",
                    '["Alice Smith"]',
                    "A study",
                    '["cs.LG"]',
                    "2022-01-01",
                ),
                (
                    "4",
                    "Sparse retrieval systems",
                    '["Eve"]',
                    "Retrieval augmented generation",
                    '["cs.CL"]',
                    "2025-01-01",
                ),
                (
                    "5",
                    "Retrieval-augmented generation",
                    '["Dan"]',
                    "Another method",
                    '["cs.CL"]',
                    "2020-01-01",
                ),
            ]
            c.executemany(
                "INSERT INTO arxiv.papers(id,title,authors,summary,categories,published) VALUES (%s,%s,%s,%s,%s,%s)",
                samples,
            )
        with cls.conn.cursor() as c:
            c.execute(
                "INSERT INTO arxiv.papers(id,title,authors,summary,categories,published) VALUES ('2501.00001','Special paper','[]','A different subject','[\"cs.AI\"]','2025-01-01')"
            )
        sys.path.insert(0, str(SRC))
        from arxiv_mcp_server import pg_adapter

        cls.adapter = pg_adapter

    @classmethod
    def tearDownClass(cls):
        cls.conn.close()
        for k, v in cls.old_env.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

    def search(self, query, **kwargs):
        content = asyncio.run(
            self.adapter.handle_search({"query": query, "max_results": 50, **kwargs})
        )
        result = json.loads(content[0].text)
        return {p["id"] for p in result["papers"]}

    def test_or_returns_union_and_phrase_does_not_match_separated_words(self):
        self.assertEqual(self.search('"neural networks"'), {"1", "2"})
        self.assertEqual(
            self.search('"neural networks" OR "retrieval augmented generation"'),
            {"1", "2", "4"},
        )
        self.assertEqual(
            self.search(
                '"retrieval augmented generation" OR "retrieval-augmented generation"'
            ),
            {"4", "5"},
        )

    def test_field_scope_exclusions_and_precedence(self):
        self.assertEqual(self.search('ti:"neural networks"'), {"1"})
        self.assertEqual(self.search('abs:"neural networks"'), {"2"})
        self.assertEqual(self.search('au:"Alice Smith" AND neural'), {"1", "3"})
        self.assertEqual(self.search("neural ANDNOT (survey OR ti:graph)"), {"3"})
        self.assertEqual(
            self.search("(ti:graph OR ti:sparse) AND abs:attention"), {"2"}
        )

    def test_filters_are_retained_for_boolean_search_and_plain_fallback(self):
        self.assertEqual(
            self.search(
                "neural OR retrieval", date_from="2024-01-01", categories=["cs.CL"]
            ),
            {"4"},
        )
        self.assertEqual(self.search("neural", date_from="2026-01-01"), set())
        self.assertEqual(self.search("neural", categories=["cs.RO"]), set())
        self.assertEqual(
            self.search("", date_from="2024-01-01", categories=["cs.CL"]), {"4"}
        )

    def test_exact_id_does_not_bypass_date_or_category_filters(self):
        self.assertEqual(self.search("arxiv:2501.00001"), {"2501.00001"})
        self.assertEqual(self.search("2501.00001", date_from="2026-01-01"), set())
        self.assertEqual(self.search("2501.00001", categories=["cs.RO"]), set())

    def test_invalid_syntax_is_error_not_broad_match(self):
        result = asyncio.run(self.adapter.handle_search({"query": "neural OR"}))
        self.assertTrue(result[0].text.startswith("Error:"))


if __name__ == "__main__":
    unittest.main()
