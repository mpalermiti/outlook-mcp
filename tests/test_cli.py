"""Tests for the CLI commands.

`cmd_status` reaches for the auth record, and the auth record lives in the
operator's real settings directory by default. These tests patch
``auth._auth_record_path`` to a tmp_path so the CLI is exercised without ever
touching the host's real ``~/.outlook-mcp/auth_record.json`` — a test that
reads the real record is order-dependent, host-dependent, and one refactor
away from deleting it.
"""

import pytest

from outlook_mcp import cli
from outlook_mcp.config import Config


@pytest.fixture(autouse=True)
def _record_in_tmp(tmp_path, monkeypatch):
    """Point the record path at an empty tmp dir for every test here."""
    from outlook_mcp import auth as auth_module

    monkeypatch.setattr(auth_module, "_auth_record_path", lambda: tmp_path / "auth_record.json")


def test_status_without_config_exits_with_the_fix(capsys, monkeypatch):
    monkeypatch.setattr(cli, "load_config", lambda: Config())
    with pytest.raises(SystemExit) as exc:
        cli.cmd_status()
    assert exc.value.code == 1
    assert "client_id" in capsys.readouterr().out


def test_status_authenticated_flow_reads_only_the_patched_record(capsys, monkeypatch):
    """No record in the tmp dir -> 'not authenticated', and the real settings
    directory was never consulted."""
    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))
    cli.cmd_status()  # must not raise
    out = capsys.readouterr().out
    assert "not authenticated" in out
    assert "outlook-mcp auth" in out


def test_status_surfaces_a_named_remedy_instead_of_the_generic_line(capsys, monkeypatch):
    """The AADSTS70000 dead end must not print "Run: outlook-mcp auth" alone.

    That line reads as "any re-auth will do" when only a fresh login exits,
    so when the refresh set a startup error with its own remedy, that is
    what status prints.
    """
    from unittest.mock import patch

    from outlook_mcp.errors import StaleConsentError

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))

    def _dead_end_refresh(self):
        self.startup_error = StaleConsentError()
        return False

    with patch.object(cli.AuthManager, "try_cached_token", _dead_end_refresh):
        cli.cmd_status()

    out = capsys.readouterr().out
    assert "not authenticated" in out
    assert "AADSTS70000" in out
    assert "Log in again" in out
    assert "Run: outlook-mcp auth" not in out


def test_auth_without_config_exits_with_the_fix(capsys, monkeypatch):
    monkeypatch.setattr(cli, "load_config", lambda: Config())
    with pytest.raises(SystemExit) as exc:
        cli.cmd_auth()
    assert exc.value.code == 1
    assert "client_id" in capsys.readouterr().out


@pytest.mark.parametrize("command", [cli.cmd_auth, cli.cmd_status, cli.cmd_logout])
def test_an_unloadable_config_exits_with_the_repair_not_a_traceback(command, capsys, monkeypatch):
    """Whatever the server's pre-run check catches, the CLI catches too."""
    monkeypatch.setattr(
        cli, "load_config", lambda: (_ for _ in ()).throw(OSError(13, "Permission denied"))
    )
    with pytest.raises(SystemExit) as exc:
        command()

    assert exc.value.code == 1
    err = capsys.readouterr().err
    assert "cannot start" in err
    assert "Traceback" not in err


def test_auth_reports_a_structured_failure_with_its_remedy(capsys, monkeypatch):
    """The unencrypted-cache refusal arrives as an OutlookMCPError whose text
    already says what to change — print that, not a traceback."""
    from unittest.mock import patch

    from outlook_mcp.errors import UnencryptedTokenCacheError

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))
    with (
        patch(
            "outlook_mcp.auth.AuthManager.login_interactive",
            side_effect=UnencryptedTokenCacheError(),
        ),
        pytest.raises(SystemExit) as exc,
    ):
        cli.cmd_auth()

    assert exc.value.code == 1
    err = capsys.readouterr().err
    assert "allow_unencrypted_token_cache" in err
    assert "Traceback" not in err


def test_logout_removes_this_instances_record_and_says_what_stays(capsys, monkeypatch):
    from outlook_mcp import auth as auth_module

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))
    record = auth_module._auth_record_path()  # the autouse fixture's tmp path
    record.write_text("{}")

    cli.cmd_logout()

    assert not record.exists()  # the record this instance serves from is gone
    out = capsys.readouterr().out
    assert "auth_record.json" in out
    assert "outlook-mcp auth" in out  # next start asks for auth again
    # the OS cache entry is shared and stays — and the old advice ("remove
    # 'outlook-mcp' from Keychain Access") named an entry that never existed
    assert "Microsoft.Developer.IdentityService" in out
    assert "left in place" in out
    assert "remove 'outlook-mcp'" not in out


def test_auth_refusal_prints_the_remedy_not_a_traceback(capsys, monkeypatch):
    """A refused first consent (e.g. an app registration missing one of the
    delegated permissions) exits 1 with the registration remedy — the
    concrete scopes make this reachable now, where a .default consent used
    to "succeed" and strand the session instead."""
    from azure.core.exceptions import ClientAuthenticationError

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))

    def _refused(self):
        raise ClientAuthenticationError(
            "Authentication failed: AADSTS65001: the user or administrator "
            "has not consented to use the application"
        )

    monkeypatch.setattr(cli.AuthManager, "login_interactive", _refused)
    with pytest.raises(SystemExit) as exc:
        cli.cmd_auth()
    assert exc.value.code == 1
    err = capsys.readouterr().err
    assert "Sign-in was refused" in err
    assert "app registration" in err


@pytest.mark.parametrize(
    "refusal",
    [
        # azure-identity's own device-code timeout text (device_code.py):
        # no AADSTS code at all.
        "Timed out waiting for user to authenticate",
        # A declined sign-in surfaces as access_denied, not a consent-code.
        "Authentication failed: access_denied: The user has denied access to the app",
        # A generic invalid grant (the AADSTS70000 family) — the code a
        # stale consent or a revoked refresh token reports.
        "Authentication failed: AADSTS70000: The requested user must first "
        "sign-in and grant the client application access",
    ],
)
def test_auth_non_consent_refusals_get_no_registration_remedy(refusal, capsys, monkeypatch):
    """Only the AADSTS65xxx refusals say anything about the app registration.

    A timed-out device code, a declined sign-in and a generic invalid grant
    are not the registration's fault, so they get the error text alone —
    blaming the registration there would send the operator re-reading the
    app setup for a problem their app does not have.
    """
    from azure.core.exceptions import ClientAuthenticationError

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))

    def _refused(self):
        raise ClientAuthenticationError(refusal)

    monkeypatch.setattr(cli.AuthManager, "login_interactive", _refused)
    with pytest.raises(SystemExit) as exc:
        cli.cmd_auth()
    assert exc.value.code == 1
    err = capsys.readouterr().err
    assert "Sign-in was refused" in err
    assert "app registration" not in err


@pytest.mark.parametrize(
    ("config", "expected", "unexpected"),
    [
        pytest.param(Config(client_id="test-id"), "read-write scopes", "read-only", id="default"),
        pytest.param(
            Config(client_id="test-id", read_only=True),
            "read-write scopes",
            "read-only scopes",
            id="read_only-alone-still-asks-wide",
        ),
        pytest.param(
            Config(client_id="test-id", read_only=True, read_only_consent=True),
            "read-only scopes",
            "read-write",
            id="read_only_consent",
        ),
    ],
)
def test_auth_says_which_scopes_it_is_about_to_ask_for(
    capsys, monkeypatch, config, expected, unexpected
):
    """What the consent screen will list, said before the browser opens."""
    monkeypatch.setattr(cli, "load_config", lambda: config)
    monkeypatch.setattr(cli.AuthManager, "login_interactive", lambda self: None)

    cli.cmd_auth()

    out = capsys.readouterr().out
    assert expected in out
    assert unexpected not in out


def test_status_names_a_sign_in_saved_for_another_app(capsys, monkeypatch, tmp_path):
    """`client_id` changed, nobody signed in again: status must not say "authenticated"."""
    from azure.identity import AuthenticationRecord

    record = AuthenticationRecord(
        tenant_id="consumers",
        client_id="99999999-aaaa-4bbb-8ccc-000000000009",
        authority="https://login.microsoftonline.com/consumers",
        home_account_id="home-1",
        username="user@example.com",
    )
    (tmp_path / "auth_record.json").write_text(record.serialize(), encoding="utf-8")
    monkeypatch.setattr(
        cli, "load_config", lambda: Config(client_id="11111111-aaaa-4bbb-8ccc-000000000001")
    )

    cli.cmd_status()

    out = capsys.readouterr().out
    assert "not authenticated" in out
    assert "different app" in out
    assert "Status: authenticated" not in out


def test_auth_warns_a_read_only_config_before_asking_for_write_access(capsys, monkeypatch):
    """`read_only: true` alone still consents the read-write set (#101).

    That is right when the flag will be flipped later, and a trap for the one
    setup the README offers for a credential that cannot write: a second,
    read-only app registration. Its owner upgrading from 1.23.0 — which signed
    in with `.default` — would see a consent screen asking that app for write
    access. Say so before the browser opens.
    """
    signed_in = []
    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id", read_only=True))
    monkeypatch.setattr(cli.AuthManager, "login_interactive", lambda self: signed_in.append(True))

    cli.cmd_auth()

    out = capsys.readouterr().out
    assert "read_only_consent" in out
    assert "read-only app registration" in out
    assert out.index("read_only_consent") < out.index("Done.")
    assert signed_in == [True]


@pytest.mark.parametrize(
    "config",
    [
        pytest.param(Config(client_id="test-id"), id="read-write"),
        pytest.param(
            Config(client_id="test-id", read_only=True, read_only_consent=True),
            id="read_only_consent",
        ),
    ],
)
def test_auth_says_nothing_extra_when_the_consent_matches_the_config(capsys, monkeypatch, config):
    monkeypatch.setattr(cli, "load_config", lambda: config)
    monkeypatch.setattr(cli.AuthManager, "login_interactive", lambda self: None)

    cli.cmd_auth()

    assert "read-only app registration" not in capsys.readouterr().out


def test_status_prints_a_refusal_with_its_remedy_not_a_traceback(capsys, monkeypatch):
    """`try_cached_token` re-raises the plaintext-cache refusal; status must not crash on it."""
    from unittest.mock import patch

    from outlook_mcp.errors import UnencryptedTokenCacheError

    monkeypatch.setattr(cli, "load_config", lambda: Config(client_id="test-id"))

    def _refuses(self):
        raise UnencryptedTokenCacheError()

    with patch.object(cli.AuthManager, "try_cached_token", _refuses):
        cli.cmd_status()

    out = capsys.readouterr().out
    assert "not authenticated" in out
    assert "allow_unencrypted_token_cache" in out
