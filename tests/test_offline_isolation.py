"""The offline suite's isolation and guard (#267) — see "Offline isolation" in conftest.py.

Each breach here is deliberate, and harmless if the guard were broken: a credential
that is never used; cache stores whose construction never reads the Keychain item
(azure-identity and msal_extensions read it on first use, which never comes); a
lookup of a `.invalid` name; sockets aimed at documentation-only addresses
(RFC 5737 / RFC 3849) or at loopback.
"""

import json
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
    CONFIG_DIR_ENV,
    LIVE_MARKERS,
    PROXY_VARIABLES,
    TOKEN_CACHE_STORES,
    expect_breach,
    hold_breaches_against,
    is_loopback,
    is_offline,
    needs_dns,
    selects_a_live_tier,
)

CLIENT_ID = "00000000-0000-0000-0000-000000000000"
REAL_SETTINGS = os.path.realpath(os.path.expanduser("~/.outlook-mcp"))
REPO = Path(__file__).resolve().parents[1]


# ── The guard fires ──────────────────────────────────────────────────────


def test_building_a_real_credential_fails_the_test():
    with expect_breach() as caught:
        DeviceCodeCredential(client_id=CLIENT_ID)
    assert "DeviceCodeCredential" in str(caught.value)


@pytest.mark.real_credential
def test_a_test_marked_real_credential_may_build_one_and_nothing_more():
    credential = DeviceCodeCredential(client_id=CLIENT_ID, authority="https://127.0.0.1:9")
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


def test_a_datagram_off_this_machine_fails_the_test():
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock, expect_breach() as caught:
        sock.sendto(b"x", ("192.0.2.1", 9))
    assert "192.0.2.1" in str(caught.value)


def test_a_dns_lookup_fails_the_test():
    with expect_breach() as caught:
        socket.getaddrinfo("outlook-mcp-offline-guard-test.invalid", 443)
    assert ".invalid" in str(caught.value)


def test_a_lookup_that_never_reaches_dns_is_allowed():
    assert socket.getaddrinfo("127.0.0.1", 9, type=socket.SOCK_STREAM)
    assert socket.getaddrinfo("localhost", 9, type=socket.SOCK_STREAM)


@pytest.mark.parametrize(
    ("host", "dns"),
    [
        (None, False),
        ("", False),
        ("localhost", False),
        ("127.0.0.1", False),
        ("192.0.2.1", False),
        ("::1", False),
        ("fe80::1%lo0", False),
        (b"login.microsoftonline.com", True),
        ("login.microsoftonline.com", True),
        ("example.invalid", True),
    ],
)
def test_what_needs_dns(host, dns):
    assert needs_dns(host) is dns


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
    except BaseException:  # what uvicorn's lifespan does
        pass
    assert conftest._breaches, "the report hook would have nothing to fail the test with"
    conftest._breaches.clear()


def test_expect_breach_forgives_only_its_own_breach():
    try:
        DeviceCodeCredential(client_id=CLIENT_ID)
    except BaseException:
        pass
    with expect_breach():
        msal.PublicClientApplication(CLIENT_ID, authority="https://127.0.0.1:9/consumers")
    assert len(conftest._breaches) == 1
    assert "DeviceCodeCredential" in conftest._breaches[0]
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


# ── Only the offline tests are guarded; only offline runs move the directory ──


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


def _default_markexpr(request) -> str:
    addopts = request.config.getini("addopts")
    return addopts[addopts.index("-m") + 1]


def test_the_guard_exempts_exactly_the_tiers_the_default_run_deselects(request):
    assert set(re.findall(r"not (\w+)", _default_markexpr(request))) == set(LIVE_MARKERS)
    assert not selects_a_live_tier(_default_markexpr(request))


@pytest.mark.parametrize(
    ("markexpr", "live"),
    [
        ("", True),
        ("live", True),
        ("integration", True),
        ("live_write", True),
        ("live or integration", True),
        ("not live", True),  # still selects integration
        ("not integration and not live and not live_write", False),
        ("not (integration or live or live_write)", False),
        ("asyncio", False),
        ("real_credential", False),
    ],
)
def test_which_runs_keep_the_real_settings_directory(markexpr, live):
    assert selects_a_live_tier(markexpr) is live


# ── An offline test has no settings directory but its own ────────────────


def test_conftest_names_the_packages_variable():
    from outlook_mcp import config

    assert CONFIG_DIR_ENV == config.CONFIG_DIR_ENV


def test_an_offline_test_sees_an_empty_settings_directory():
    from outlook_mcp import auth

    settings = Path(os.environ[CONFIG_DIR_ENV])
    assert settings.is_dir()
    assert not any(settings.iterdir())
    assert auth._auth_record_path().parent == settings
    assert auth._load_auth_record() is None
    assert not any(os.environ.get(name) for name in PROXY_VARIABLES)


def test_every_default_the_package_holds_points_at_the_tests_own_directory():
    from outlook_mcp import auth, cli, config, errors, server, validation

    here = os.environ[CONFIG_DIR_ENV]
    assert os.path.realpath(here) != REAL_SETTINGS
    for module in (config, auth, cli, server, validation):
        assert module.DEFAULT_CONFIG_DIR == here, module.__name__
    assert config.load_config.__defaults__ == (here,)
    assert config.save_config.__defaults__ == (here,)
    assert cli.load_config is config.load_config
    assert errors._config_dir() == here
    assert os.path.realpath(config.Config().attachments_dir) == os.path.realpath(
        os.path.join(here, "attachments")
    )


def test_an_offline_run_moves_the_settings_directory_before_anything_is_imported(request):
    if selects_a_live_tier(request.config.getoption("markexpr")):
        pytest.skip("this run selects a live tier, which needs the real settings directory")
    assert conftest.run_settings is not None
    assert os.path.realpath(conftest.run_settings) != REAL_SETTINGS


@pytest.mark.parametrize("attempt", [1, 2])
def test_nothing_a_test_writes_to_the_default_directory_outlives_it(attempt):
    from outlook_mcp import cli, config

    here = Path(os.environ[CONFIG_DIR_ENV])
    assert os.path.realpath(config.save_config.__defaults__[0]) == os.path.realpath(here)
    assert not any(here.iterdir()), "an earlier test's settings are still here"

    client_id = f"00000000-0000-0000-0000-00000000000{attempt}"
    config.save_config(config.Config(client_id=client_id))  # no directory given
    assert (here / "config.json").is_file()
    assert cli.load_config().client_id == client_id


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
    assert out == [os.environ[CONFIG_DIR_ENV], "None"]


# ── The hooks, end to end, in a pytest of their own ──────────────────────

_CHILD_TESTS = """
import os

import pytest
from azure.identity import DeviceCodeCredential

# Imported during collection, as the real suite does: server.py reads the config
# at import to choose its tools.
from outlook_mcp import config, server

STAND_IN = os.environ["STAND_IN_REAL_SETTINGS"]
TOOLS = set(server.mcp._tool_manager._tools)


def test_offline_never_sees_the_real_directory_or_its_config():
    assert config.DEFAULT_CONFIG_DIR != STAND_IN
    assert "outlook_create_rule" not in TOOLS


def test_a_swallowed_breach_still_fails():
    try:
        DeviceCodeCredential(client_id="00000000-0000-0000-0000-000000000000")
    except BaseException:
        pass


@pytest.mark.live
def test_live_keeps_the_real_directory_and_runs_unguarded():
    assert config.DEFAULT_CONFIG_DIR == STAND_IN
    assert "outlook_create_rule" in TOOLS  # mailbox_settings_write, read at import
    DeviceCodeCredential(client_id="00000000-0000-0000-0000-000000000000")  # never used
"""


@pytest.fixture
def child_suite(tmp_path):
    """A throwaway pytest project using this conftest, and a stand-in for ~/.outlook-mcp.

    The stand-in's config turns on mailbox_settings_write, which adds the rules
    tools at import: what makes the real config visible in the tool surface.
    """
    stand_in = tmp_path / "stand-in-real-settings"
    stand_in.mkdir(mode=0o700)
    config_file = stand_in / "config.json"
    config_file.write_text(json.dumps({"mailbox_settings_write": True}), encoding="utf-8")
    config_file.chmod(0o600)

    project = tmp_path / "project"
    project.mkdir()
    (project / "conftest.py").write_text(
        "from tests.conftest import *  # noqa: F403 — the hooks under test\n", encoding="utf-8"
    )
    markers = "".join(f"    {marker}: x\n" for marker in (*LIVE_MARKERS, "real_credential"))
    (project / "pytest.ini").write_text(f"[pytest]\nmarkers =\n{markers}", encoding="utf-8")
    (project / "test_hooks.py").write_text(_CHILD_TESTS, encoding="utf-8")

    env = {key: value for key, value in os.environ.items() if key != "OUTLOOK_MCP_TOOLSETS"}
    env.update(
        {
            CONFIG_DIR_ENV: str(stand_in),
            "STAND_IN_REAL_SETTINGS": os.path.abspath(stand_in),
            "PYTHONPATH": os.pathsep.join(filter(None, [str(REPO), env.get("PYTHONPATH")])),
        }
    )

    def run(markexpr: str) -> subprocess.CompletedProcess:
        return subprocess.run(
            [sys.executable, "-m", "pytest", "-p", "no:cacheprovider", "-rA", "-m", markexpr],
            cwd=project,
            env=env,
            capture_output=True,
            text=True,
            timeout=180,
        )

    return run


def test_an_offline_run_isolates_from_import_and_fails_a_swallowed_breach(child_suite):
    result = child_suite("not integration and not live and not live_write")
    out = result.stdout
    assert "PASSED test_hooks.py::test_offline_never_sees_the_real_directory_or_its_config" in out
    assert "FAILED test_hooks.py::test_a_swallowed_breach_still_fails" in out
    assert "built a real DeviceCodeCredential" in out
    assert "1 failed, 1 passed, 1 deselected" in out, out


def test_a_live_run_keeps_the_real_directory_and_installs_no_guard(child_suite):
    result = child_suite("live")
    out = result.stdout
    assert "PASSED test_hooks.py::test_live_keeps_the_real_directory_and_runs_unguarded" in out
    assert "1 passed, 2 deselected" in out, out
    assert result.returncode == 0
