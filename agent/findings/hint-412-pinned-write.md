# hint-412-pinned-write (#96) — 2026-10-13

## What changed
`errors._HINT_TABLE[(412, "ErrorIrresolvableConflict")]` no longer says "Nothing was modified".
It now says the item changed after it was read, that this is usually someone else's edit, that
after a retried request the change may already have been applied, and to re-read and check
before sending again. Retry behaviour is untouched.

Tests (red → green):
- `tests/test_error_wrapper.py::test_wrap_412_conflict_says_the_change_may_have_landed_and_to_re_read`
- `tests/test_error_text_reaches_client.py::test_pinned_write_conflict_hint_reaches_the_client`
  (real `ODataError` → `_wrap_tool_errors` → MCP client; asserts the text the model reads)

## Facts checked
- VERIFIED (grep `src/`): the only `If-Match` write is `calendar_write.update_event`'s series
  re-send (`patch_config.headers.add("If-Match", ...)`). So every 412 this server can produce
  comes from a pinned write; the hint table entry is the right place, no per-call override needed.
- VERIFIED (read `.venv/.../kiota_http/middleware/retry_handler.py`, kiota-http 1.12.3 per
  `uv.lock`): `RetryHandler` retries 429/503/504 for `PATCH` among its default allowed methods,
  so the scenario in #96 is real in the shipped stack.
- VERIFIED (`src/outlook_mcp/tools/todo_attachments.py:262-271`): the repo already disables
  retry per request with `RetryHandlerOption(should_retry=False)` for a non-idempotent POST.
- ASSUMED (from #96, not reproduced): a 504 can arrive for a PATCH that committed. It cannot be
  forced from a consumer mailbox.

## The other option: disable retry for pinned writes
One line: add `RetryHandlerOption(should_retry=False)` to `patch_config` when `if_match` is set.

| | Keep retry (current) | Disable retry on the pinned PATCH |
|---|---|---|
| 504/503, first attempt did **not** commit | retry succeeds — agent sees success | agent sees a 503/504 error and must retry itself |
| 504/503, first attempt **did** commit | retry gets 412; hint now says "may already have been applied" | agent sees a 503/504 error, equally unsure whether it landed |
| 429 (request was not processed) | retry after Retry-After succeeds | agent sees 429 and must back off itself |
| Concurrent edit between read and write | 412, correct | 412, correct |

The pin already makes the retry safe: a retry can never overwrite a newer change, it can only be
refused. With the hint corrected, the committed-then-retried case ends in the same uncertainty
either way, and keeping retry wins every other row. `should_retry=False` would also turn off the
429 retry, which is pure loss. And re-running the re-anchor after a landed write finds the zones
equal and sends nothing (per #96), so the agent's follow-up is harmless.

**Conclusion:** no concrete reason to change retry behaviour. Not parked as a question.
