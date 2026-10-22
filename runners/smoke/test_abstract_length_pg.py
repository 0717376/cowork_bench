"""Optional full-main integration against an empty dedicated PostgreSQL database."""
from pathlib import Path
import contextlib
import importlib.util
import io
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch
import openpyxl
import psycopg2

ROOT = Path(__file__).resolve().parents[2]
TASK = ROOT / 'tasks/finalpool/arxiv-federated-learning-ru-teamly-excel'


def module(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


@unittest.skipUnless(os.environ.get('BENCH_ABSTRACT_TEST_PG_HOST'), 'requires dedicated PostgreSQL')
class AbstractLengthPGTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        name = os.environ.get('BENCH_ABSTRACT_TEST_PG_DATABASE', 'cowork_abstract_length_test')
        if name != 'cowork_abstract_length_test':
            raise RuntimeError('Refusing any database except cowork_abstract_length_test')
        cls.config = dict(host=os.environ['BENCH_ABSTRACT_TEST_PG_HOST'],
                          port=int(os.environ.get('BENCH_ABSTRACT_TEST_PG_PORT', '5432')),
                          dbname=name, user=os.environ.get('BENCH_ABSTRACT_TEST_PG_USER', 'postgres'),
                          password=os.environ.get('BENCH_ABSTRACT_TEST_PG_PASSWORD', ''))
        cls.real_connect = staticmethod(psycopg2.connect)
        cls.db = cls.real_connect(**cls.config)
        cls.db.autocommit = True
        cls.seed = module(TASK / 'preprocess/main.py', 'abstract_seed')
        with cls.db.cursor() as cur:
            cur.execute("SELECT pg_advisory_lock(hashtext('cowork_abstract_length_test'))")
            cur.execute('''DROP SCHEMA IF EXISTS arxiv CASCADE; DROP SCHEMA IF EXISTS teamly CASCADE;
CREATE SCHEMA arxiv; CREATE SCHEMA teamly;
CREATE TABLE arxiv.papers(id text PRIMARY KEY,title text,authors jsonb,summary text,published timestamptz,categories jsonb);
CREATE TABLE teamly.pages(id int PRIMARY KEY,title text,body text,parent_id int);
INSERT INTO teamly.pages VALUES(4,'База знаний по федеративному обучению','hub',NULL);''')
            for i, paper in enumerate(cls.seed.FL_PAPERS, 5):
                cur.execute('INSERT INTO teamly.pages VALUES(%s,%s,%s,4)',
                            (i, paper['title'], 'federated learning'))

    @classmethod
    def tearDownClass(cls):
        try:
            with cls.db.cursor() as cur:
                cur.execute('DROP SCHEMA arxiv CASCADE; DROP SCHEMA teamly CASCADE')
        finally:
            cls.db.close()

    def evaluate(self, scenario):
        with contextlib.redirect_stdout(io.StringIO()):
            self.seed.inject_arxiv_papers(self.db)
        canonical = {p['id']: p['summary'] for p in self.seed.FL_PAPERS}
        if scenario in ['unicode-valid', 'utf8-invalid']:
            for i, pid in enumerate(canonical):
                # Decoded Unicode code points, including non-BMP/combining chars.
                canonical[pid] = 'Ж😀e\u0301\n' * (40+i)
                with self.db.cursor() as cur:
                    cur.execute('UPDATE arxiv.papers SET summary=%s WHERE id=%s', (canonical[pid], pid))
        lengths = {pid: len(text) for pid, text in canonical.items()}
        mean = round(sum(lengths.values()) / len(lengths))
        if scenario == 'poisoned':
            lengths = dict.fromkeys(lengths, -999); mean = 101
        elif scenario == 'blank':
            lengths = dict.fromkeys(lengths, None); mean = 101
        elif scenario == 'observed-wrong':
            lengths = dict(zip(['1602.05629','1812.06127','1908.07873','1912.04977'], [367,345,312,328])); mean = 338
        elif scenario == 'off-by-one':
            lengths['1602.05629'] += 1
        elif scenario == 'wrong-mean':
            mean += 1
        elif scenario == 'nonnumeric-mean':
            mean = 'not a number'
        elif scenario == 'self-consistent-wrong':
            lengths = dict.fromkeys(lengths, 113); mean = 113
        elif scenario == 'utf8-invalid':
            lengths = {pid: len(text.encode('utf-8')) for pid, text in canonical.items()}
            mean = round(sum(lengths.values()) / len(lengths))
        if scenario == 'missing-canonical':
            with self.db.cursor() as cur:
                cur.execute('DELETE FROM arxiv.papers WHERE id=%s', ('1602.05629',))
        hidden = scenario == 'database-error'
        if hidden:
            with self.db.cursor() as cur:
                cur.execute('ALTER TABLE arxiv.papers RENAME TO hidden_papers')
        try:
            with tempfile.TemporaryDirectory() as tmp:
                folder = Path(tmp); path = folder / 'Federated_Learning_Papers.xlsx'
                book = openpyxl.Workbook(); details = book.active; details.title = 'Paper Details'
                header = ['Paper_ID','Title','Authors','Year','Abstract_Length']
                rows = [[p['id'], p['title'], 'Synthetic author', int(p['published'][:4]), lengths[p['id']]]
                        for p in self.seed.FL_PAPERS]
                if scenario == 'duplicate-id': rows[-1][0] = rows[0][0]
                elif scenario == 'unknown-id': rows[-1][0] = '9999.99999'
                elif scenario == 'numeric-text':
                    for row in rows: row[4] = str(row[4])
                    mean = str(mean)
                elif scenario == 'reordered':
                    order = [4,2,1,0,3]; header = [header[i] for i in order]
                    rows = [[row[i] for i in order] for row in reversed(rows)]
                details.append(header)
                for row in rows: details.append(row)
                summary = book.create_sheet('Summary'); summary.append(['Metric','Value'])
                for metric, value in [('Total_Papers',4),('Avg_Abstract_Length',mean),('Earliest_Year',2016),('Latest_Year',2019)]:
                    summary.append([metric,value])
                book.save(path); book.close(); before = path.read_bytes()
                grader = module(TASK / 'evaluation/main.py', 'abstract_pg_grader')
                result_path = folder / 'result.json'
                with patch.object(grader.psycopg2, 'connect', side_effect=lambda **kw: self.real_connect(**self.config)), \
                     patch.object(sys, 'argv', ['grader','--agent_workspace',str(folder),'--res_log_file',str(result_path)]), \
                     contextlib.redirect_stdout(io.StringIO()):
                    with self.assertRaises(SystemExit) as exited:
                        grader.main()
                self.assertEqual(path.read_bytes(), before, 'grader must not rewrite the workbook')
                return exited.exception.code, json.loads(result_path.read_text())
        finally:
            if hidden:
                with self.db.cursor() as cur:
                    cur.execute('ALTER TABLE arxiv.hidden_papers RENAME TO papers')

    def test_real_main_canonical_seed_and_workbook_controls(self):
        for scenario in ['valid','poisoned','blank','observed-wrong','off-by-one','wrong-mean',
                         'self-consistent-wrong','duplicate-id','unknown-id','missing-canonical',
                         'database-error','unicode-valid','utf8-invalid','numeric-text','reordered','nonnumeric-mean']:
            with self.subTest(scenario=scenario):
                code, result = self.evaluate(scenario)
                expected = scenario in ['valid','unicode-valid','numeric-text','reordered']
                self.assertGreaterEqual(result['accuracy'], 80)
                self.assertEqual(code == 0, expected, result)
                self.assertEqual(bool(result['critical_failed']), not expected)


if __name__ == '__main__':
    unittest.main()
