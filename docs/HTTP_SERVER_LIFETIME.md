# HTTP fixture lifetime

The mock servers in `moex-earnings-calendar-alert` and
`terminal-insales-pw-pricing-excel-word-gcal` must remain available after the
preprocess Python interpreter exits. An asyncio subprocess transport can close
and kill its child while the event loop shuts down. Launch a detached
`subprocess.Popen` with redirected streams and `start_new_session=True`.
The task container remains responsible for process lifetime and cleanup.

Run in the Linux benchmark environment:

```sh
python -m unittest discover -s runners/smoke -p test_preprocess_server_lifetime.py -v
```

The tests copy the actual tracked task inputs to temporary directories, invoke
the real setup functions in a child interpreter, wait for that interpreter to
exit and compare the served bytes. They use ephemeral ports and stop only the
servers they create. They make no database or model calls. macOS child watchers
can mask the original failure; Linux execution is required.
