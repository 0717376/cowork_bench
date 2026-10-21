# Email send error channel

`send_email` must raise a FastMCP tool error when the service raises or returns
false. Returning an error sentence is successful MCP content and hides failure
from clients using `CallToolResult.isError`. Successful send responses are unchanged.
Only the send tool is changed here; other email tools are outside this patch.

Run with the email server dependencies installed, from the repository root:

```sh
python -m unittest discover -s local_servers/emails-mcp/tests -p test_send_error_channel.py -v
```

Tests use the real registered tool, MCP serialization/client and EmailService,
with a synthetic backend. No SMTP, model or PostgreSQL connection is used. On the
base revision three negative controls fail; the success control remains valid.
This does not recalculate benchmark scores or add client-side string heuristics.
