"""Real MCP stdio read/write/reload of a legacy graph; no model requests."""

import json
import os
from pathlib import Path
import tempfile
import unittest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from utils.memory_graph import read_memory_graph

SERVER = os.environ.get(
    "BENCH_MEMORY_SERVER_PATH", "/opt/local_servers/servers/src/memory/dist/index.js"
)


class MemoryMCPRoundtrip(unittest.IsolatedAsyncioTestCase):
    async def test_legacy_graph_survives_mcp_read_and_write(self):
        with tempfile.TemporaryDirectory() as temp:
            p = Path(temp) / "memory.json"
            graph = {
                "entities": [
                    {
                        "name": "seed",
                        "entityType": "note",
                        "observations": ["benchmark seed"],
                    }
                ],
                "relations": [],
            }
            p.write_text(json.dumps(graph, indent=2))
            before = p.read_bytes()
            params = StdioServerParameters(
                command="node",
                args=[SERVER],
                env={**os.environ, "MEMORY_FILE_PATH": str(p)},
            )
            async with stdio_client(params) as (r, w):
                async with ClientSession(r, w) as session:
                    await session.initialize()
                    result = await session.call_tool("read_graph", {})
                    self.assertFalse(result.isError)
                    self.assertEqual(json.loads(result.content[0].text), graph)
                    self.assertEqual(p.read_bytes(), before)
                    result = await session.call_tool(
                        "create_entities",
                        {
                            "entities": [
                                {
                                    "name": "new",
                                    "entityType": "finding",
                                    "observations": ["analysis completed"],
                                }
                            ]
                        },
                    )
                    self.assertFalse(result.isError)
                    result = await session.call_tool(
                        "create_relations",
                        {
                            "relations": [
                                {"from": "new", "to": "seed", "relationType": "extends"}
                            ]
                        },
                    )
                    self.assertFalse(result.isError)
            stored = read_memory_graph(p)
            self.assertEqual({e["name"] for e in stored["entities"]}, {"seed", "new"})
            self.assertEqual(
                stored["relations"],
                [{"from": "new", "to": "seed", "relationType": "extends"}],
            )
            self.assertEqual(len(p.read_text().splitlines()), 3)
            async with stdio_client(params) as (r, w):
                async with ClientSession(r, w) as session:
                    await session.initialize()
                    result = await session.call_tool("read_graph", {})
                    self.assertFalse(result.isError)
                    self.assertEqual(json.loads(result.content[0].text), stored)


if __name__ == "__main__":
    unittest.main()
