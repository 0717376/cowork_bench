"""Optional real PostgreSQL integration; NEVER point this at a benchmark run DB."""
import contextlib
import importlib.util
import io
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import psycopg2
from runners.smoke.test_required_email_contract import ROOT, SPECS, workspace


@unittest.skipUnless(os.environ.get('BENCH_EMAIL_TEST_PG_HOST'),
                     'requires dedicated cowork_required_email_test PostgreSQL')
class RequiredEmailPGTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        database = os.environ.get('BENCH_EMAIL_TEST_PG_DATABASE', 'cowork_required_email_test')
        if database != 'cowork_required_email_test':
            raise RuntimeError('Refusing to use any database except cowork_required_email_test')
        cls.config = dict(host=os.environ['BENCH_EMAIL_TEST_PG_HOST'],
                          port=int(os.environ.get('BENCH_EMAIL_TEST_PG_PORT', '5432')),
                          dbname=database,
                          user=os.environ.get('BENCH_EMAIL_TEST_PG_USER', 'postgres'),
                          password=os.environ.get('BENCH_EMAIL_TEST_PG_PASSWORD', ''))
        cls.real_connect = staticmethod(psycopg2.connect)
        cls.db = cls.real_connect(**cls.config)
        cls.db.autocommit = True
        with cls.db.cursor() as cur:
            cur.execute('''CREATE SCHEMA IF NOT EXISTS email;
CREATE SCHEMA IF NOT EXISTS gcal; CREATE SCHEMA IF NOT EXISTS canvas;
CREATE TABLE IF NOT EXISTS email.messages(id int PRIMARY KEY,subject text,body_text text,from_addr text,to_addr jsonb);
CREATE TABLE IF NOT EXISTS email.sent_log(message_id int);
CREATE TABLE IF NOT EXISTS email.attachments(message_id int,filename text,content_type text,content bytea);
CREATE TABLE IF NOT EXISTS gcal.events(summary text,description text,start_datetime timestamptz,end_datetime timestamptz);
CREATE TABLE IF NOT EXISTS canvas.announcements(course_id int,title text,message text);
TRUNCATE gcal.events,canvas.announcements;
INSERT INTO gcal.events VALUES
('Москва Санкт-Петербург отправление','поездка','2026-03-10 06:00Z','2026-03-10 07:00Z'),
('прибытие заселение','trip','2026-03-10 11:00Z','2026-03-10 13:30Z'),
('ужин','dinner','2026-03-10 18:00Z','2026-03-10 20:00Z'),
('экскурсия Новгород','поездка','2026-03-12 12:00Z','2026-03-12 21:00Z'),
('возвращение','поездка','2026-03-15 14:00Z','2026-03-15 21:00Z');
INSERT INTO canvas.announcements VALUES(9991,'Экскурсия Новгород','Поездка');''')

    @classmethod
    def tearDownClass(cls):
        cls.db.close()

    def test_real_sql_all_three_graders_and_office_fixtures(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            workspace(folder)
            for task, recipient, filename, subject, body in SPECS:
                for scenario in ['missing','unsent','missing-attachment','wrong-name','wrong-type',
                                 'corrupt','unrelated-attachment','wrong-recipient','valid','generic-mime']:
                    with self.subTest(task=task, scenario=scenario):
                        with self.db.cursor() as cur:
                            cur.execute('TRUNCATE email.messages,email.sent_log,email.attachments')
                            if scenario != 'missing':
                                to = 'x'+recipient if scenario == 'wrong-recipient' else recipient
                                cur.execute('INSERT INTO email.messages VALUES(1,%s,%s,%s,%s::jsonb)',
                                            (subject,body,'sender@example.invalid',json.dumps([to])))
                                if scenario != 'unsent':
                                    cur.execute('INSERT INTO email.sent_log VALUES(1)')
                                if scenario != 'missing-attachment':
                                    data = b'corrupt' if scenario == 'corrupt' else (folder / filename).read_bytes()
                                    suffix = 'spreadsheetml.sheet' if filename.endswith('xlsx') else 'wordprocessingml.document'
                                    mime = 'application/vnd.openxmlformats-officedocument.'+suffix
                                    if scenario == 'wrong-type':
                                        mime = 'text/plain'
                                    elif scenario == 'generic-mime':
                                        mime = 'application/octet-stream'
                                    cur.execute('INSERT INTO email.attachments VALUES(%s,%s,%s,%s)',
                                                (2 if scenario == 'unrelated-attachment' else 1,
                                                 'Other'+filename if scenario == 'wrong-name' else filename,
                                                 mime,psycopg2.Binary(data)))
                        path = ROOT / 'tasks/finalpool' / task / 'evaluation/main.py'
                        spec = importlib.util.spec_from_file_location('required_email_pg_grader',path)
                        grader = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(grader)
                        result_file = folder / 'result.json'
                        with patch.object(grader.psycopg2,'connect',side_effect=lambda **kw: self.real_connect(**self.config)), \
                             patch.object(sys,'argv',['grader','--agent_workspace',str(folder),
                                                       '--res_log_file',str(result_file)]), \
                             contextlib.redirect_stdout(io.StringIO()):
                            with self.assertRaises(SystemExit) as exited:
                                grader.main()
                        result = json.loads(result_file.read_text())
                        expected = scenario in ['valid','generic-mime']
                        self.assertEqual(exited.exception.code == 0,expected)
                        self.assertGreaterEqual(result['accuracy'],70)


if __name__ == '__main__':
    unittest.main()
