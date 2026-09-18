"""The macOS shared-Keychain collision, reproduced offline (review of #61).

On macOS every cache *name* maps to the same Keychain item — service
``Microsoft.Developer.IdentityService``, account ``MSALCache``; the name only
picks msal_extensions' signal file, which is what a credential consults to
decide whether the shared item belongs to its cache. A first cut of
multi-account support gave each account its own name (``outlook-mcp-<name>``),
which works on Windows (a file per name) and Linux (libsecret keyed by name),
but on macOS means ``auth <b>`` after ``auth <a>`` serializes b's still-empty
cache over the shared item and a's next silent refresh fails — a bug no offline
suite could see, because none of them model the single-slot store.

These tests model exactly that store, so reintroducing per-account cache
names (or losing the per-account AuthenticationRecord pin) fails loudly here
instead of only in a live run against a real Keychain.

The design under test, from ``outlook_mcp.auth``: one CACHE_NAME for every
account, with each account's AuthenticationRecord pinning which identity in
the merged cache its credential serves.
"""

from azure.core.credentials import AccessToken
from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AuthenticationRecord

from outlook_mcp.auth import AuthManager
from outlook_mcp.config import AccountConfig, Config

GRAPH_SCOPE = "https://graph.microsoft.com/.default"

# client_id -> (home_account_id, username): the identity the human signs in
# as during each app registration's (simulated) device-code flow.
IDENTITIES = {
    "id1-abcd": ("net-home-id", "net@example.com"),
    "id2-efgh": ("neko-home-id", "neko@example.com"),
}


def _two_account_config() -> Config:
    return Config(
        accounts=[
            AccountConfig(name="net", client_id="id1-abcd"),
            AccountConfig(name="neko", client_id="id2-efgh"),
        ]
    )


class _MacKeychainStore:
    """Persistence with macOS Keychain semantics: ONE shared slot.

    ``load(name)`` is what a credential named ``name`` sees when it opens the
    cache: the shared slot when ``name``'s signal file exists (that name has
    written before), an empty cache otherwise. ``save(name, cache)`` replaces
    the shared slot — there is nothing else to write to — and creates
    ``name``'s signal file. The name never keys the storage; that is the
    whole collision.
    """

    def __init__(self) -> None:
        self._slot: dict[str, str] = {}  # home_account_id -> access token
        self._signal_files: set[str] = set()
        self.written_names: list[str] = []  # audit: the name given to each save

    def load(self, name: str) -> dict[str, str]:
        if name in self._signal_files:
            return dict(self._slot)
        return {}

    def save(self, name: str, cache: dict[str, str]) -> None:
        self.written_names.append(name)
        self._slot = dict(cache)
        self._signal_files.add(name)


class _FakeDeviceCodeCredential:
    """DeviceCodeCredential driven against the macOS-semantics store.

    Faithful to the surface ``auth.py`` drives: constructed with ``client_id``
    / ``tenant_id`` / ``cache_persistence_options`` /
    ``authentication_record`` / ``disable_automatic_authentication`` /
    ``prompt_callback``, then asked for tokens via ``get_token(scope)``.

    Token semantics mirror azure-identity + MSAL over a shared cache:

    - a credential with an ``authentication_record`` serves exactly that
      record's identity — the record is the pin;
    - without a record, the interactive device flow signs in as the identity
      the app registration maps to, and the account is MERGED into the
      credential's view of the cache before the whole view is serialized;
    - a cache miss in silent mode (``disable_automatic_authentication``)
      raises ``ClientAuthenticationError`` — azure-identity's own "cannot
      refresh this identity, re-authenticate" signal.
    """

    def __init__(
        self,
        store: _MacKeychainStore,
        *,
        client_id: str,
        tenant_id: str,
        cache_persistence_options,
        timeout: int = 0,
        disable_automatic_authentication: bool = False,
        prompt_callback=None,
        authentication_record: AuthenticationRecord | None = None,
    ) -> None:
        self._store = store
        self._client_id = client_id
        self._tenant_id = tenant_id
        self._cache_name = cache_persistence_options.name
        self._silent_only = disable_automatic_authentication
        self._record = authentication_record

    def get_token(self, *scopes) -> AccessToken:
        if self._record is not None:
            home = self._record.home_account_id
            username = self._record.username
        else:
            home, username = IDENTITIES[self._client_id]

        cache = self._store.load(self._cache_name)
        if home in cache:
            return AccessToken(cache[home], 0)

        if self._silent_only or self._record is not None:
            # The identity this credential is pinned to is not in its view of
            # the cache: silent refresh cannot proceed.
            raise ClientAuthenticationError(
                f"no token for {home} in the cache named {self._cache_name!r}"
            )

        # Interactive device-code flow: MSAL merges the new account into the
        # credential's view, then serializes the whole view to the store.
        cache[home] = f"token-for-{home}"
        self._store.save(self._cache_name, cache)
        self._auth_record = AuthenticationRecord(
            tenant_id=self._tenant_id,
            client_id=self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            home_account_id=home,
            username=username,
        )
        return AccessToken(cache[home], 0)


def _install_mac_keychain(monkeypatch, tmp_path) -> _MacKeychainStore:
    """Point AuthManager's credential + record storage at the simulated store."""
    from outlook_mcp import auth as auth_module

    store = _MacKeychainStore()
    record_dir = tmp_path / "records"
    record_dir.mkdir()

    monkeypatch.setattr(
        auth_module,
        "DeviceCodeCredential",
        lambda **kwargs: _FakeDeviceCodeCredential(store, **kwargs),
    )
    monkeypatch.setattr(
        auth_module,
        "_auth_record_path",
        lambda account=None: record_dir / f"auth_record-{account or 'default'}.json",
    )
    # The unencrypted-cache check is a Linux concern; this suite models macOS,
    # where storage is always encrypted.
    monkeypatch.setattr(auth_module, "_unencrypted_fallback_will_be_used", lambda: False)
    return store


# ── The collision: auth <b> after <a> must not evict <a> ────────────────


def test_second_login_keeps_the_first_accounts_tokens_offline(monkeypatch, tmp_path):
    """`auth neko` after `auth net`, then a server restart: BOTH accounts still
    authenticate.

    With per-account cache names this is exactly the live-only macOS failure:
    neko's first save serialized a cache that had never seen net over the
    shared Keychain item, and net's silent refresh came back empty — here that
    runs offline, so the regression cannot ship quietly again.
    """
    from outlook_mcp.auth import CACHE_NAME

    store = _install_mac_keychain(monkeypatch, tmp_path)
    config = _two_account_config()

    # Day 1: the CLI authenticates each account in turn.
    cli = AuthManager(config)
    cli.login_interactive("net")
    cli.login_interactive("neko")

    # Every save went through the ONE shared name — the invariant that keeps
    # the single Keychain item multi-account. Per-account names fail here (and
    # below) because neko's save replaced, not merged, the shared slot.
    assert set(store.written_names) == {CACHE_NAME}

    # Day 2: the server starts and refreshes silently, per account.
    server = AuthManager(config)
    assert server.try_cached_token() is True
    assert server.authenticated_accounts == ["net", "neko"]

    # The FIRST account's token specifically survived the second save…
    net_cred = server.get_account_credential("net")
    token = net_cred.get_token(GRAPH_SCOPE)
    assert token.token == "token-for-net-home-id"  # …and is still net's


# ── The pin: the record, not the client, selects the identity ───────────


def test_each_accounts_record_pins_its_own_identity(monkeypatch, tmp_path):
    """One shared cache name only works because the per-account
    AuthenticationRecord selects which identity in the merged cache a
    credential serves: both records are on disk, both identities are in the
    cache, and each account's silent credential comes back with its own
    token — not the other account's."""
    from outlook_mcp.auth import _load_auth_record

    _install_mac_keychain(monkeypatch, tmp_path)
    config = _two_account_config()

    cli = AuthManager(config)
    cli.login_interactive("net")
    cli.login_interactive("neko")

    records = {name: _load_auth_record(name) for name in ("net", "neko")}
    assert records["net"].home_account_id == "net-home-id"
    assert records["neko"].home_account_id == "neko-home-id"

    for name, record in records.items():
        server = AuthManager(config)
        cred = server._make_credential(
            account=name, auth_record=record, silent=True
        )
        token = cred.get_token(GRAPH_SCOPE)
        assert token.token == f"token-for-{records[name].home_account_id}"
