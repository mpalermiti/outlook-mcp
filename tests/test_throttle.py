"""Tests for throttle.py — Retry-After honoring on the raw-httpx Graph paths.

The SDK path retries 429/503 via kiota's RetryHandler, but the raw-httpx
$batch / delta paths bypass it. These helpers add that retry behavior:
envelope-level retry for a throttled request, and per-sub-request retry for a
$batch whose envelope is 200 but individual sub-responses are 429/503.
"""


from outlook_mcp.throttle import (
    parse_retry_after,
    retry_throttled_subrequests,
    send_with_retry,
)


class _Resp:
    def __init__(self, status_code, headers=None):
        self.status_code = status_code
        self.headers = headers or {}


class _FakeClient:
    """Returns queued responses; records the requests it saw.

    Exposes ``get``/``post`` (not ``request``) to match how send_with_retry
    dispatches and how the real call sites invoke httpx.
    """

    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = []

    async def get(self, url, headers=None, content=None):
        self.calls.append(("GET", url))
        return self._responses.pop(0)

    async def post(self, url, headers=None, content=None):
        self.calls.append(("POST", url))
        return self._responses.pop(0)


def _recording_sleep():
    slept = []

    async def _sleep(seconds):
        slept.append(seconds)

    return _sleep, slept


# ── parse_retry_after ─────────────────────────────────────────────────


def test_parse_retry_after_numeric():
    assert parse_retry_after({"Retry-After": "5"}, default=1.0) == 5.0


def test_parse_retry_after_case_insensitive():
    assert parse_retry_after({"retry-after": "3"}, default=1.0) == 3.0


def test_parse_retry_after_missing_uses_default():
    assert parse_retry_after({}, default=2.5) == 2.5


def test_parse_retry_after_nonnumeric_uses_default():
    assert parse_retry_after({"Retry-After": "soon"}, default=2.0) == 2.0


# ── send_with_retry ───────────────────────────────────────────────────


async def test_send_returns_immediately_on_200():
    client = _FakeClient([_Resp(200)])
    sleep, slept = _recording_sleep()
    resp = await send_with_retry(client, "GET", "u", headers={}, sleep=sleep)
    assert resp.status_code == 200
    assert len(client.calls) == 1
    assert slept == []


async def test_send_retries_429_then_succeeds_honoring_retry_after():
    client = _FakeClient([_Resp(429, {"Retry-After": "5"}), _Resp(200)])
    sleep, slept = _recording_sleep()
    resp = await send_with_retry(client, "GET", "u", headers={}, sleep=sleep)
    assert resp.status_code == 200
    assert len(client.calls) == 2
    assert slept == [5.0]  # waited the server-specified Retry-After


async def test_send_retries_503():
    client = _FakeClient([_Resp(503, {"Retry-After": "1"}), _Resp(200)])
    sleep, slept = _recording_sleep()
    resp = await send_with_retry(client, "GET", "u", headers={}, sleep=sleep)
    assert resp.status_code == 200
    assert len(client.calls) == 2


async def test_send_gives_up_after_max_retries():
    client = _FakeClient([_Resp(429, {"Retry-After": "1"})] * 5)
    sleep, slept = _recording_sleep()
    resp = await send_with_retry(
        client, "GET", "u", headers={}, max_retries=2, sleep=sleep
    )
    assert resp.status_code == 429  # returns the final throttled response
    assert len(client.calls) == 3  # initial + 2 retries
    assert len(slept) == 2


# ── retry_throttled_subrequests ───────────────────────────────────────


async def test_subrequests_all_ok_single_batch():
    requests = [{"id": "0"}, {"id": "1"}]
    seen = []

    async def post_batch(reqs):
        seen.append([r["id"] for r in reqs])
        return {"responses": [{"id": "0", "status": 200}, {"id": "1", "status": 200}]}

    sleep, slept = _recording_sleep()
    merged = await retry_throttled_subrequests(post_batch, requests, sleep=sleep)
    assert merged["0"]["status"] == 200
    assert merged["1"]["status"] == 200
    assert seen == [["0", "1"]]  # one batch, no retry
    assert slept == []


async def test_subrequest_429_is_retried_then_succeeds():
    requests = [{"id": "0"}, {"id": "1"}]
    seen = []

    async def post_batch(reqs):
        seen.append([r["id"] for r in reqs])
        if len(seen) == 1:
            return {
                "responses": [
                    {"id": "0", "status": 200},
                    {"id": "1", "status": 429, "headers": {"Retry-After": "2"}},
                ]
            }
        return {"responses": [{"id": "1", "status": 200}]}

    sleep, slept = _recording_sleep()
    merged = await retry_throttled_subrequests(post_batch, requests, sleep=sleep)
    assert merged["0"]["status"] == 200
    assert merged["1"]["status"] == 200  # retried, not recorded as permanent failure
    assert seen == [["0", "1"], ["1"]]  # second batch only re-sends the throttled id
    assert slept == [2.0]


async def test_subrequest_429_bounded_records_final():
    requests = [{"id": "0"}]

    async def post_batch(reqs):
        return {"responses": [{"id": "0", "status": 429, "headers": {"Retry-After": "1"}}]}

    sleep, slept = _recording_sleep()
    merged = await retry_throttled_subrequests(
        post_batch, requests, max_retries=2, sleep=sleep
    )
    assert merged["0"]["status"] == 429  # bounded — recorded after retries exhausted
    assert len(slept) == 2
