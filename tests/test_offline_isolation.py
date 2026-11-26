"""The offline suite's isolation and guard (#267) — see "Offline isolation" in conftest.py.

Each breach here is deliberate and harmless if the guard were broken: a credential
that is never used, cache stores in tmp_path, sockets aimed at documentation-only
addresses (RFC 5737 / RFC 3849) or at loopback.
"""

import os
import re
import socket
import subprocess
import sys
from pathlib import Path
from unittest.mock import MagicMock

import msal
import msal_extensions
import pytest
from azure.identity import DeviceCodeCredential, TokenCachePersistenceOptions
from azure.identity._persistent_cache import _load_persistent_cache
from msal_extensions.persistence import BasePersistence

from tests import conftest
from tests.conftest import (
    LIVE_MARKERS,
    PROXY_VARIABLES,
    TOKEN_CACHE_STORES,
    expect_breach,
    hold_breaches_against,
    is_loopback,
    is_offline,
)

CLIENT_ID = "00000000-0000-0000-0000-000000000000"


# ── The guard fires ──────────────────────────────────────────────────────


def test_building_a_real_credential_fails_the_test():
    with expect_breach() as caught:
        DeviceCodeCredential(client_id=CLIENT_ID)
    assert "DeviceCodeCredential" in str(caught.value)


@pytest.mark.real_credential
def test_a_test_marked_real_credential_may_build_one_and_nothing_more():
    credential = DeviceCodeCredential(client_id=CLIENT_ID, authority="127.0.0.1:9")
    with expect_breach():
        credential._get_app()  # the MSAL client is still refused


def test_building_an_msal_client_fails_the_test():
    with expect_breach() as caught:
        msal.PublicClientApplication(
            CLIENT_ID,
            authority="https://127.0.0.1:9/consumers",
            validate_authority=False,
            instance_discovery=False,
        )
    assert "MSAL client" in str(caught.value)


def test_azure_identitys_own_route_to_the_token_cache_fails_the_test():
    """The route the lifespan took: azure-identity opening its persistent cache."""
    options = TokenCachePersistenceOptions(name="outlook-mcp-offline-guard-test")
    with expect_breach() as caught:
        _load_persistent_cache(options, False)
    assert "token cache" in str(caught.value)


@pytest.mark.parametrize("store", ["FilePersistence", "KeychainPersistence"])
def test_opening_a_token_cache_store_fails_the_test(tmp_path, store):
    with expect_breach():
        getattr(msal_extensions, store)(str(tmp_path / "cache"))


def test_every_token_cache_store_msal_extensions_offers_is_guarded():
    stores = {
        name
        for name, value in vars(msal_extensions).items()
        if isinstance(value, type) and issubclass(value, BasePersistence)
    } - {"BasePersistence"}
    assert stores == set(TOKEN_CACHE_STORES)


@pytest.mark.parametrize(
    ("family", "address"),
    [(socket.AF_INET, ("192.0.2.1", 9)), (socket.AF_INET6, ("2001:db8::1", 9, 0, 0))],
)
@pytest.mark.parametrize("method", ["connect", "connect_ex"])
def test_a_connection_off_this_machine_fails_the_test(family, address, method):
    try:
        sock = socket.socket(family, socket.SOCK_STREAM)
    except OSError:
        pytest.skip("no IPv6 on this host")
    with sock:
        sock.settimeout(0.05)
        with expect_breach() as caught:
            getattr(sock, method)(address)
    assert address[0] in str(caught.value)


def test_a_loopback_connection_is_allowed():
    with socket.socket() as server, socket.socket() as client:
        server.bind(("127.0.0.1", 0))
        server.listen(1)
        client.connect(server.getsockname())
        server.accept()[0].close()


@pytest.mark.parametrize(
    ("host", "loopback"),
    [
        ("localhost", True),
        ("127.0.0.1", True),
        ("127.0.0.2", True),
        ("::1", True),
        ("::1%lo0", True),
        ("::ffff:127.0.0.1", True),
        ("0.0.0.0", False),
        ("192.0.2.1", False),
        ("::ffff:192.0.2.1", False),
        ("login.microsoftonline.com", False),
    ],
)
def test_what_counts_as_loopback(host, loopback):
    assert is_loopback(host) is loopback


# ── A breach fails the test even when the exception is swallowed ─────────


def test_a_swallowed_breach_is_still_recorded():
    try:
        DeviceCodeCredential(client_id=CLIENT_ID)
    except BaseException:  # noqa: BLE001 — what uvicorn's lifespan does
        pass
    assert conftest._breaches, "the report hook would have nothing to fail the test with"
    conftest._breaches.clear()


def _report(outcome: str) -> pytest.TestReport:
    return pytest.TestReport("t.py::t", ("t.py", 0, "t"), {}, outcome, None, "call")


def test_a_phase_that_passed_despite_a_breach_is_failed_with_the_reason():
    report = _report("passed")
    hold_breaches_against(report, ["the reason", "the reason"], raised=False)
    assert report.failed
    assert report.longrepr == "the reason"


def test_a_phase_that_failed_for_another_reason_carries_the_breach_too():
    report = _report("failed")
    hold_breaches_against(report, ["the reason"], raised=False)
    assert report.sections == [("offline isolation (#267)", "the reason")]

    raised = _report("failed")
    hold_breaches_against(raised, ["the reason"], raised=True)
    assert raised.sections == []  # the failure already is the breach


# ── Only the offline tests are guarded ───────────────────────────────────


class _Item:
    def __init__(self, *markers: str) -> None:
        self._markers = markers

    def get_closest_marker(self, name: str):
        return name if name in self._markers else None


def test_the_live_tiers_are_left_alone():
    assert is_offline(_Item())
    assert is_offline(_Item("asyncio", "real_credential"))
    for marker in LIVE_MARKERS:
        assert not is_offline(_Item("asyncio", marker))


def test_the_guard_exempts_exactly_the_tiers_the_default_run_deselects(request):
    addopts = " ".join(request.config.getini("addopts"))
    assert set(re.findall(r"not (\w+)", addopts)) == set(LIVE_MARKERS)


# ── An offline test has no settings directory of its own ────────────────


def test_an_offline_test_sees_an_empty_settings_directory():
    from outlook_mcp import auth, config

    settings = Path(os.environ[config.CONFIG_DIR_ENV])
    assert settings.is_dir()
    assert not any(settings.iterdir())
    assert auth._auth_record_path().parent == settings
    assert auth._load_auth_record() is None
    assert not any(os.environ.get(name) for name in PROXY_VARIABLES)


async def test_the_real_lifespan_reaches_no_sign_in():
    """What every `Client(mcp)` test runs: the lifespan, unpatched by the test."""
    from outlook_mcp import server

    async with server.lifespan(MagicMock()) as state:
        assert state["auth"].credential is None


def test_a_child_process_inherits_the_empty_settings_directory():
    probe = (
        "from outlook_mcp import auth, config\n"
        "print(config.DEFAULT_CONFIG_DIR)\n"
        "print(auth._load_auth_record())\n"
    )
    out = subprocess.run(
        [sys.executable, "-c", probe], capture_output=True, text=True, check=True, timeout=60
    ).stdout.splitlines()
    assert out == [os.environ["OUTLOOK_MCP_CONFIG_DIR"], "None"]
