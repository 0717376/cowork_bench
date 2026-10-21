# Email attachment upload root

The benchmark launches the email server from its source directory and passes
`--attachment_upload_path` separately. PG SMTP previously ignored that setting
and opened relative paths from cwd. A file in the agent workspace could therefore
be missing, or a same-name file in the server directory could be attached instead.

`create_services` now injects the configured root into PG SMTP. Relative paths
start at that root. Absolute paths inside it remain supported. Canonical paths,
including symlink targets, must stay inside the root and refer to regular files.
All attachments are read before the first database access, so an invalid later
attachment cannot create a partial sent message. MIME type/filename/content stay
unchanged. Standalone callers with no configured root retain their prior cwd
semantics; the benchmark config always supplies the workspace upload root.
This patch does not change the real SMTP backend or other T2F/native file tools.

```sh
python -m unittest discover -s local_servers/emails-mcp/tests -p test_attachment_upload_root.py -v
```

Install the email server dependencies first. Tests use real temporary files,
`create_services`, EmailService and PG SMTP with a synthetic SQL receipt. No real
email, model or PostgreSQL connection occurs. MCP error signaling is covered by
the separate email error-channel patch; a failed send remains a failure here.
Scores and historical results are not recalculated.
