"""Real Sheets handlers/backend, isolated storage seam and real MCP transport."""
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from contextlib import asynccontextmanager
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from mcp_google_sheets import pg_services as pg, server
from mcp.server.fastmcp import FastMCP
from mcp.shared.memory import create_connected_server_and_client_session


class StorageSeam:
    def setUp(self):
        self.cells = {(0, 0): "header", (2, 1): "neighbor"}
        self.original = dict(self.cells)
        self.acquire = Mock(return_value=object())
        self.write = Mock(side_effect=self.write_cells)
        for name, value in [('_get_conn', self.acquire), ('_put_conn', Mock()),
                            ('_resolve_sheet_id', Mock(return_value=7)),
                            ('_resolve_default_sheet', Mock(return_value=(7, 'Sheet1'))),
                            ('_write_cells', self.write)]:
            p = patch.object(pg, name, value)
            p.start()
            self.addCleanup(p.stop)
        p = patch('psycopg2.connect', side_effect=AssertionError('real DB forbidden'))
        p.start()
        self.addCleanup(p.stop)
        self.ctx = SimpleNamespace(request_context=SimpleNamespace(
            lifespan_context=SimpleNamespace(sheets_service=pg.PgSheetsService(None))))

    def write_cells(self, conn, spreadsheet, sheet, sr, sc, values):
        for ri, row in enumerate(values):
            for ci, value in enumerate(row):
                self.cells[sr + ri, sc + ci] = value
        return sum(map(len, values))

    def request(self, address, values):
        return pg.ValuesUpdateRequest(None, 'synthetic', address, 'USER_ENTERED', {'values': values})

    def assert_untouched(self):
        self.assertEqual(self.cells, self.original)
        self.acquire.assert_not_called()
        self.write.assert_not_called()


class RangeTests(StorageSeam, unittest.TestCase):
    def test_valid_a3_keeps_header_and_neighbor(self):
        result = self.request('Sheet1!A3', [['row3']]).execute()
        self.assertEqual(result['updatedCells'], 1)
        self.assertEqual(self.cells, {(0, 0): 'header', (2, 0): 'row3', (2, 1): 'neighbor'})

    def test_duplicate_prefix_rejected_without_connection(self):
        with self.assertRaises(ValueError):
            self.request('Sheet1!Sheet1!A3:D3', [['row3']]).execute()
        self.assert_untouched()

    def test_malformed_ranges_reject_before_connection(self):
        for address in ['', 'Sheet1!A0', 'Sheet1!A-1', 'Sheet1!A1:B2:C3',
                        'Sheet1!D3:A1', 'Sheet1!:', "'unterminated!A1", 'Sheet1!A1!B2']:
            with self.subTest(address=address), self.assertRaises(ValueError):
                self.request(address, [['x']]).execute()
        self.assert_untouched()

    def test_non_scalar_and_non_2d_values_reject_before_connection(self):
        for values in [[[['nested']]], [[{'x': 1}]], ['not a row'], {'x': 1},
                       [[float('nan')]], [[float('inf')]]]:
            with self.subTest(values=values), self.assertRaises(ValueError):
                self.request('Sheet1!A3', values).execute()
        self.assert_untouched()

    def test_all_json_scalars_remain_supported(self):
        result = self.request('Sheet1!A3:F3', [['text', True, 1, 1.5, None, '']]).execute()
        self.assertEqual(result['updatedCells'], 6)
        self.assertEqual(self.cells[0, 0], 'header')

    def test_parser_valid_open_row_column_and_quoted_ranges(self):
        cases = {
            'A3': (None, 2, 0, 2, 0),
            'Sheet1!': ('Sheet1', None, None, None, None),
            'Sheet1!A3:D3': ('Sheet1', 2, 0, 2, 3),
            'Sheet1!A3:5': ('Sheet1', 2, 0, 4, None),
            'Sheet1!A3:': ('Sheet1', 2, 0, None, None),
            'Sheet1!A:C': ('Sheet1', None, 0, None, 2),
            'Sheet1!3:5': ('Sheet1', 2, None, 4, None),
            "'Sales Q1'!A3:C5": ('Sales Q1', 2, 0, 4, 2),
            "'O''Brien! Q1'!A3": ("O'Brien! Q1", 2, 0, 2, 0),
            "'Sheet1'": ('Sheet1', None, None, None, None),
        }
        for address, expected in cases.items():
            with self.subTest(address=address):
                self.assertEqual(pg._parse_a1_range(address), expected)

    def test_whole_sheet_and_open_writes_keep_documented_origins(self):
        self.request("'Sheet1'", [['whole']]).execute()
        self.assertEqual(self.cells[0, 0], 'whole')
        self.request('Sheet1!B3:', [['open']]).execute()
        self.assertEqual(self.cells[2, 1], 'open')

    def test_invalid_later_batch_range_does_not_partially_write(self):
        body = {'data': [{'range': 'Sheet1!A3', 'values': [['first']]},
                         {'range': 'Sheet1!Sheet1!A4', 'values': [['second']]}]}
        with self.assertRaises(ValueError):
            pg.ValuesBatchUpdateRequest(None, 'synthetic', body).execute()
        self.assert_untouched()

    def test_invalid_later_batch_matrix_does_not_partially_write(self):
        body = {'data': [{'range': 'Sheet1!A3', 'values': [['first']]},
                         {'range': 'Sheet1!A4', 'values': [[['nested']]]}]}
        with self.assertRaises(ValueError):
            pg.ValuesBatchUpdateRequest(None, 'synthetic', body).execute()
        self.assert_untouched()

    def test_invalid_request_envelopes_reject_before_connection(self):
        for body in [[], "invalid", 1, False]:
            with self.subTest(body=body), self.assertRaises(ValueError):
                pg.ValuesUpdateRequest(None, 'synthetic', 'Sheet1!A3', 'RAW', body).execute()
        for body in [[], {'data': {}}, {'data': [None]}]:
            with self.subTest(batch=body), self.assertRaises(ValueError):
                pg.ValuesBatchUpdateRequest(None, 'synthetic', body).execute()
        self.assert_untouched()

    def test_valid_batch_preserves_unaddressed_cells(self):
        result = pg.ValuesBatchUpdateRequest(None, 'synthetic', {'data': [
            {'range': 'Sheet1!A3', 'values': [['first']]},
            {'range': 'Sheet1!A4', 'values': [['second']]}]}).execute()
        self.assertEqual(result['totalUpdatedCells'], 2)
        self.assertEqual(self.cells[0, 0], 'header')
        self.assertEqual(self.cells[2, 1], 'neighbor')

    def test_tool_rejects_qualified_or_conflicting_ranges(self):
        for address in ['Sheet1!A3', 'Other!A3', 'Sheet1!Sheet1!A3']:
            with self.subTest(address=address), self.assertRaises(ValueError):
                server.update_cells('synthetic', 'Sheet1', address, [['x']], self.ctx)
        self.assert_untouched()

    def test_tool_quotes_actual_sheet_name(self):
        result = server.update_cells('synthetic', "O'Brien! Q1", 'A3', [['x']], self.ctx)
        self.assertEqual(result['updatedCells'], 1)
        self.assertEqual(self.cells[2, 0], 'x')
        pg._resolve_sheet_id.assert_called_once_with(self.acquire.return_value, 'synthetic', "O'Brien! Q1")

    def test_batch_tool_prevalidates_all_ranges(self):
        with self.assertRaises(ValueError):
            server.batch_update_cells('synthetic', 'Sheet1', {'A3': [['x']], 'Other!A4': [['y']]}, self.ctx)
        self.assert_untouched()


class ProtocolTests(StorageSeam, unittest.IsolatedAsyncioTestCase):
    async def test_wire_error_and_recovery_without_header_damage(self):
        @asynccontextmanager
        async def lifespan(mcp):
            yield SimpleNamespace(sheets_service=pg.PgSheetsService(None))
        mcp = FastMCP('sheets-contract-test', lifespan=lifespan)
        mcp.add_tool(server.update_cells)
        async with create_connected_server_and_client_session(mcp._mcp_server) as client:
            bad = await client.call_tool('update_cells', {'spreadsheet_id': 'synthetic', 'sheet': 'Sheet1',
                  'range': 'Sheet1!A3', 'data': [['wrong']]})
            self.assertTrue(bad.isError)
            self.assert_untouched()
            good = await client.call_tool('update_cells', {'spreadsheet_id': 'synthetic', 'sheet': 'Sheet1',
                  'range': 'A3', 'data': [['right']]})
            self.assertFalse(good.isError)
            self.assertEqual(self.cells, {(0, 0): 'header', (2, 0): 'right', (2, 1): 'neighbor'})


if __name__ == '__main__':
    unittest.main()
