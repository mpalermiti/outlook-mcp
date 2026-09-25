"""The two-process invariants for a shared token cache, pinned offline.

How the cache really works, from azure-identity's own persistence layer:

- On macOS, every persisted azure-identity cache on the host — this server's,
  Visual Studio's, the Azure CLI's — serializes into ONE Keychain item:
  service ``Microsoft.Developer.IdentityService``, account ``MSALCache``.
  ``TokenCachePersistenceOptions(name=...)`` does NOT buy a second Keychain
  item. It buys a distinct *signal file* — on disk
  ``~/.IdentityService/outlook-mcp.nocae``: azure-identity appends a
  suffix to every cache name (``.nocae`` non-CAE, ``.cae`` CAE), so the
  name keys the signal file, never the storage —
  which is the lock-and-merge coordinator: a credential whose signal file
  says "changed" reloads the shared item, merges its accounts, and writes the
  whole view back under a lock. Windows (a DPAPI file per name) and Linux
  (libsecret keyed by name) really do key storage by name; macOS does not.
  The comment must not claim more for a cache name than that.

That fact makes the one-process-per-account design (each instance sets
``OUTLOOK_MCP_CONFIG_DIR``) correct only while two invariants hold, and this
suite holds them offline:

1. every credential writes through the ONE cache name — one signal file, so
   every writer goes through the same lock-and-merge into the one item;
2. authenticating instance B must not evict instance A: with one shared name,
   B's save serializes a view that *includes* A's account; giving each
   instance its own name (the tempting "isolation") makes B load an empty
   view instead, and B's save clobbers the shared item — a bug no mock of
   Graph can see, because it happens entirely inside the cache;
3. each instance's AuthenticationRecord is what pins which identity in the
   merged item its credential serves — the record, not the client, selects
   the account.

Reintroducing per-instance cache names, or losing the per-instance record,
fails loudly here instead of only in a live run against a real Keychain.
"""

import os
import subprocess
import sys

from azure.core.credentials import AccessToken
from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AuthenticationRecord

from outlook_mcp import auth as auth_module
from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config

GRAPH_SCOPE = "https://graph.microsoft.com/.default"

# client_id -> (home_account_id, username): the identity the human signs in
# as during that instance's (simulated) device-code flow.
IDENTITIES = {
    "id1-abcd": ("net-home-id", "net@example.com"),
    "id2-efgh": ("neko-home-id", "neko@example.com"),
}


class _MacKeychainStore:
    """Persistence with macOS Keychain semantics: ONE shared slot.

    ``load(name)`` is what a credential named ``name`` sees when it opens the
    cache: the shared slot when ``name``'s signal file exists (that name has
    written before), an empty cache otherwise. ``save(name, cache)`` replaces
    the shared slot with the caller's serialized view — there is nothing else
    to write to — and creates ``name``'s signal file. The name never keys the
    storage; it only keys the signal file. That is the whole collision.
    """

    def __init__(self) -> None:
        self._slot: dict[str, str] = {}  # home_account_id -> access token
        self._signal_files: set[str] = set()
        self.written_names: list[str] = []  # audit: the name given to each save
        self.constructed_cache_names: list[str] = []  # audit: per construction

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

    Token semantics mirror azure-identity + MSAL over the shared item:

    - a credential with an ``authentication_record`` serves exactly that
      record's identity, whatever ``client_id`` it was constructed with —
      the record is the pin;
    - without a record, the device flow signs in as the identity the app
      registration maps to, and the account is MERGED into the credential's
      view of the cache before the whole view is serialized;
    - a cache miss in silent mode (``disable_automatic_authentication``)
      raises ``ClientAuthenticationError`` — azure-identity's "cannot serve
      this identity, re-authenticate" signal.
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
        store.constructed_cache_names.append(self._cache_name)

    def get_token(self, *scopes) -> AccessToken:
        if self._record is not None:
            home = self._record.home_account_id
        else:
            home = IDENTITIES[self._client_id][0]

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
            username=IDENTITIES[self._client_id][1],
        )
        return AccessToken(cache[home], 0)


class _TwoInstances:
    """Two server instances against one shared cache item.

    Each instance is one process with its own settings directory (in
    production, its own ``OUTLOOK_MCP_CONFIG_DIR``) — its own config, its own
    auth record file. The directories are simulated by routing the record
    load/save functions per instance; only the token cache is shared, exactly
    as on a real host.
    """

    def __init__(self, monkeypatch) -> None:
        self.store = _MacKeychainStore()
        self._records: dict[str, AuthenticationRecord] = {}
        self._active: str | None = None
        self.managers = {
            name: AuthManager(Config(client_id=client_id))
            for name, client_id in (("net", "id1-abcd"), ("neko", "id2-efgh"))
        }

        def save_record(record: AuthenticationRecord) -> None:
            assert self._active is not None
            self._records[self._active] = record

        def load_record() -> AuthenticationRecord | None:
            assert self._active is not None
            return self._records.get(self._active)

        monkeypatch.setattr(
            auth_module,
            "DeviceCodeCredential",
            lambda **kwargs: _FakeDeviceCodeCredential(self.store, **kwargs),
        )
        monkeypatch.setattr(auth_module, "_save_auth_record", save_record)
        monkeypatch.setattr(auth_module, "_load_auth_record", load_record)
        # The unencrypted-cache check is a Linux concern; this suite models
        # macOS, where storage is always encrypted.
        monkeypatch.setattr(
            auth_module, "_unencrypted_fallback_will_be_used", lambda: False
        )

    def record_for(self, name: str) -> AuthenticationRecord:
        """The record an instance's settings directory holds."""
        return self._records[name]

    def as_instance(self, name: str) -> AuthManager:
        """Make `name` the process whose settings directory is active."""
        self._active = name
        return self.managers[name]

    def restart(self, name: str) -> AuthManager:
        """A fresh process for `name`: new AuthManager, same settings dir."""
        self._active = name
        client_id = {"net": "id1-abcd", "neko": "id2-efgh"}[name]
        manager = AuthManager(Config(client_id=client_id))
        self.managers[name] = manager
        return manager


# Instance B runs as a real second process: its own settings directory via
# OUTLOOK_MCP_CONFIG_DIR, exactly how a second server is deployed. It reports
# the cache name its credential was constructed with and stops there — the
# construction is the fact under test, not the flow.
_INSTANCE_B = """
from unittest.mock import patch

from outlook_mcp import auth as auth_module
from outlook_mcp.auth import AuthManager
from outlook_mcp.config import Config


def probe(**kwargs):
    print(kwargs["cache_persistence_options"].name)
    raise SystemExit(0)


with (
    patch.object(auth_module, "DeviceCodeCredential", probe),
    # the encrypted-store probe is a host property, not a property of the
    # name under test
    patch.object(auth_module, "_unencrypted_fallback_will_be_used", lambda: False),
):
    AuthManager(Config(client_id="id2-efgh")).login_interactive()
"""


def _run_instance_b(settings_dir) -> str:
    env = dict(os.environ)
    env["OUTLOOK_MCP_CONFIG_DIR"] = str(settings_dir)
    out = subprocess.run(
        [sys.executable, "-c", _INSTANCE_B],
        env=env,
        capture_output=True,
        text=True,
        check=True,
        timeout=60,
    )
    return out.stdout.strip()


def test_every_write_goes_through_the_one_cache_name(monkeypatch, tmp_path):
    """Invariant 1: one cache name, however many instances.

    Asserted as the literal every host uses — not the module constant, which
    a per-instance renaming would move along with the test — for both the
    in-process writer (the store audits every save) and a second process
    with its own settings directory constructing its credential the way a
    second deployed server would.

    A distinct name per instance would not isolate anything on macOS — the
    Keychain item is the same — it would only split the signal file, and each
    writer would then reload-through-its-own-lock believing it owns the item.
    """
    two = _TwoInstances(monkeypatch)

    two.as_instance("net").login_interactive()
    assert set(two.store.written_names) == {"outlook-mcp"}
    assert set(two.store.constructed_cache_names) == {"outlook-mcp"}

    settings_dir = tmp_path / "instance-neko"
    settings_dir.mkdir()
    assert _run_instance_b(settings_dir) == "outlook-mcp"


def test_second_instance_does_not_evict_the_first(monkeypatch):
    """Invariant 2: auth B after A, then a restart of A: BOTH still work.

    With per-instance cache names this is the live-only macOS failure: B's
    first save serialized a cache that had never seen A over the shared
    item, and A's silent refresh came back empty. Here that runs offline, so
    the regression cannot ship quietly again.
    """
    two = _TwoInstances(monkeypatch)

    # Day 1: each instance's operator runs `outlook-mcp auth`.
    two.as_instance("net").login_interactive()
    two.as_instance("neko").login_interactive()

    # Day 2: both servers start fresh and refresh silently — A first, i.e.
    # after B has already written through the shared item.
    net = two.restart("net")
    assert net.try_cached_token() is True
    token = net.get_credential().get_token(GRAPH_SCOPE)
    assert token.token == "token-for-net-home-id"  # still net's, not neko's

    neko = two.restart("neko")
    assert neko.try_cached_token() is True
    token = neko.get_credential().get_token(GRAPH_SCOPE)
    assert token.token == "token-for-neko-home-id"


def test_each_instances_record_pins_its_own_identity(monkeypatch):
    """Invariant 3: the record, not the client, selects the identity.

    Both identities live in the one merged item; what keeps instance A from
    serving instance B's mailbox is that A's credential is constructed with
    A's AuthenticationRecord. A record pinned to the wrong identity — or no
    record at all, where the client_id would pick a default — fails here.
    """
    two = _TwoInstances(monkeypatch)

    two.as_instance("net").login_interactive()
    two.as_instance("neko").login_interactive()

    # Cross-check: neko's record serves neko's identity even when handed to
    # a credential constructed with net's client_id — azure-identity serves
    # the record's account, not the client's.
    neko_record = two.record_for("neko")
    assert neko_record.home_account_id == "neko-home-id"
    cred = two.managers["net"]._make_credential(auth_record=neko_record, silent=True)
    assert cred.get_token(GRAPH_SCOPE).token == "token-for-neko-home-id"

    net_record = two.record_for("net")
    assert net_record.home_account_id == "net-home-id"
    cred = two.managers["neko"]._make_credential(auth_record=net_record, silent=True)
    assert cred.get_token(GRAPH_SCOPE).token == "token-for-net-home-id"
