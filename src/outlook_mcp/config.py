"""Config file management for outlook-mcp."""

import difflib
import locale
import logging
import os
import stat
import sys
import tempfile
from pathlib import Path

from pydantic import BaseModel, Field, ValidationError, field_validator, model_validator

from outlook_mcp.permissions import VALID_CATEGORIES

logger = logging.getLogger(__name__)

DEFAULT_TENANT_ID = "consumers"

# The one directory every setting lives in: config.json here, and the auth
# record next to it (auth._auth_record_path derives from DEFAULT_CONFIG_DIR).
# Overriding it (and ONLY it) via OUTLOOK_MCP_CONFIG_DIR is how you run a
# second instance against a different mailbox: the MSAL signal file stays at
# ~/.IdentityService/, so both processes still share the lock-and-merge path
# into the one Keychain item. Moving HOME instead would give each process its
# own signal file, neither would see the other's write, and they would clobber
# each other's token. Empty or unset leaves the default in place.
CONFIG_DIR_ENV = "OUTLOOK_MCP_CONFIG_DIR"


def _resolve_config_dir() -> str:
    """Resolve the settings directory once, absolutely.

    An override may be relative — a launch script's shorthand — and the
    terminal that runs ``outlook-mcp auth`` and the client that starts the
    server rarely share a working directory. A relative value left relative
    names a different settings directory in each, and the server asks for
    re-authentication forever. Anchoring at first use pins both to one
    directory. Empty or unset leaves the default in place.
    """
    override = os.environ.get(CONFIG_DIR_ENV)
    return os.path.abspath(os.path.expanduser(override or "~/.outlook-mcp"))


DEFAULT_CONFIG_DIR = _resolve_config_dir()


def _default_attachments_dir() -> str:
    """Attachments live under the settings directory, wherever it was moved to.

    Derived from ``DEFAULT_CONFIG_DIR`` and nothing else — the same single
    constant every other path derives from, so an override can never split
    the settings directory from the attachments directory.
    """
    return os.path.join(DEFAULT_CONFIG_DIR, "attachments")


# Top-level keys an older release accepted. One process serves one account
# now; a config carrying these still loads, the operator just hears about it.
# Both keys described the same removed feature, so they share one sentence.
_LEGACY_MULTI_ACCOUNT_MESSAGE = (
    "configuring multiple accounts in one process is no longer supported; "
    "run one server per account and give each its own settings directory "
    f"via the {CONFIG_DIR_ENV} environment variable"
)
_LEGACY_KEYS = {
    "accounts": _LEGACY_MULTI_ACCOUNT_MESSAGE,
    "default_account": _LEGACY_MULTI_ACCOUNT_MESSAGE,
}


# Keys whose default is the open side: absent, the server is writable, every
# write category is allowed, and sign-in asks for write access. An unknown key
# is normally accepted with a warning (a newer release's config must still boot
# an older one), but a misspelling of one of these is refused instead — ignoring
# `"readOnly": true` ran a writable server its operator believed was read-only.
_SAFETY_KEYS = ("read_only", "allow_categories", "read_only_consent")


def _squash(key: str) -> str:
    """`readOnly`, `read-only`, `READ_ONLY` and `readonly` all squash to `readonly`."""
    return "".join(ch for ch in key.lower() if ch.isalnum())


def _misspelt_safety_key(key: str) -> str | None:
    """The safety key an unknown `key` plainly means, or None.

    A different spelling of the same words, or a one-or-two-letter slip
    (`difflib` ratio 0.8 or better). Deliberately narrow: only these three keys,
    so every other unknown key keeps loading with a warning.
    """
    for safety_key in _SAFETY_KEYS:
        if _squash(key) == _squash(safety_key):
            return safety_key
    close = difflib.get_close_matches(key.lower(), _SAFETY_KEYS, n=1, cutoff=0.8)
    return close[0] if close else None


class Config(BaseModel):
    """Outlook MCP server configuration."""

    client_id: str | None = Field(default=None, description="Azure AD app client ID (BYOID)")
    tenant_id: str = Field(default=DEFAULT_TENANT_ID)
    read_only: bool = Field(default=False)
    read_only_consent: bool = Field(
        default=False,
        description=(
            "Ask Microsoft for the read scopes only at sign-in, instead of the "
            "read-write set. For a credential that cannot write: point client_id "
            "at an app registration that has never been granted write access, set "
            "this and read_only to true, then run `outlook-mcp auth`. Requires "
            "read_only: true."
        ),
    )
    allow_categories: list[str] = Field(
        default_factory=list,
        description=(
            "Optional whitelist of write-tool categories. Empty list = fully open "
            "(all writes allowed when read_only=False). Non-empty = only the listed "
            "categories are permitted."
        ),
    )
    timezone: str = Field(
        default="UTC",
        description=(
            "IANA timezone. Interprets zone-less dates, and anchors every event "
            "created — a recurring event is expanded in this zone, so the UTC "
            "default makes one shift an hour across a daylight-saving change."
        ),
    )
    attachments_dir: str = Field(
        default_factory=_default_attachments_dir,
        description=(
            "The only directory the attachment tools may read from or write to. "
            "Point it somewhere else to widen the surface; every path an agent "
            "supplies is resolved and must land inside it. Defaults to an "
            "`attachments` folder inside the settings directory."
        ),
    )
    allow_unencrypted_token_cache: bool = Field(
        default=False,
        description=(
            "Permit the OAuth token cache to be written in cleartext when no "
            "encrypted store is available (Linux without libsecret). Off by "
            "default: without it, authentication stops rather than silently "
            "persisting a reusable Graph token in plaintext."
        ),
    )

    @model_validator(mode="before")
    @classmethod
    def _warn_on_unknown_keys(cls, data: object) -> object:
        """Accept-and-warn: unknown top-level keys are ignored, not silent.

        A typo'd key used to vanish without a trace — the setting silently
        kept its default while the operator believed they had changed it.
        Unknown keys still don't fail the load (a config written for a newer
        release should still boot an older one), but each one is named on
        the way out. Known-legacy keys get the same treatment with a pointer
        to what replaced them.

        The exception is a misspelt safety key (see ``_SAFETY_KEYS``), whose
        default is the open side: that one refuses the load, so the server does
        not start with a safety setting silently missing.
        """
        if not isinstance(data, dict):
            return data
        for key in data:
            if key in _LEGACY_KEYS:
                logger.warning("Config key %r ignored: %s.", key, _LEGACY_KEYS[key])
            elif key not in cls.model_fields and (meant := _misspelt_safety_key(key)):
                raise ValueError(
                    f"config key {key!r} looks like {meant!r} misspelt. {meant!r} is a "
                    "safety setting whose default is the permissive one, so the server "
                    f"will not start with it ignored. Rename the key to {meant!r}."
                )
            elif key not in cls.model_fields:
                logger.warning(
                    "Unknown config key %r ignored — supported keys: %s.",
                    key,
                    ", ".join(sorted(cls.model_fields)),
                )
        return data

    @model_validator(mode="after")
    def _read_only_consent_needs_read_only(self) -> "Config":
        """Refuse a read-only sign-in on a server that expects to write.

        A session that consented only the read scopes cannot write, and no
        refresh widens it — every later request asks for `.default`, which
        means exactly what was consented. Left to load, each write tool would
        fail with a 403 whose hint says nothing about the consent. This is the
        trap a consent keyed on `read_only` alone used to set whenever that
        flag was flipped; with its own key, the mismatch is a config error
        with both names in it.
        """
        if self.read_only_consent and not self.read_only:
            raise ValueError(
                "read_only_consent: true needs read_only: true. A sign-in that asked "
                "only for the read scopes cannot write, so every write tool would "
                "fail. Set read_only to true, or remove read_only_consent and run "
                "`outlook-mcp auth` again to consent the write scopes."
            )
        return self

    @field_validator("allow_categories")
    @classmethod
    def _validate_allow_categories(cls, value: list[str]) -> list[str]:
        """Reject unknown category names at config load time."""
        unknown = [c for c in value if c not in VALID_CATEGORIES]
        if unknown:
            valid_list = ", ".join(sorted(VALID_CATEGORIES))
            raise ValueError(
                f"Unknown permission categories: {unknown}. Valid categories: {valid_list}"
            )
        return value


# The 0o700 directory and 0o600 file modes below are POSIX-only. On Windows os.chmod can set
# nothing but the read-only attribute, so these calls cannot enforce the owner-only access they
# ask for. What governs the path there is its Windows ACL, including whatever it inherits from
# the directory it was created under — which this code neither applies nor verifies, and which
# is not necessarily the user profile's: both the config directory and attachments_dir are
# configurable and may sit on another drive or a network share. auth.py writes the auth record
# through these same two helpers, so this covers that file too. Recorded rather than pretended
# (#85); applying a real DACL is a separate change.
def _ensure_dir(dir_path: str) -> Path:
    """Create the config directory, restricted to 0700 where the platform can enforce it."""
    path = Path(dir_path)
    path.mkdir(parents=True, exist_ok=True)
    path.chmod(0o700)  # POSIX-only; see the note above.
    return path


def atomic_write(file_path: Path, data: str) -> None:
    """Write file atomically with fsync, restricted to 0600 where the platform can enforce it."""
    dir_path = file_path.parent
    fd, tmp_path = tempfile.mkstemp(dir=str(dir_path), suffix=".tmp")
    try:
        # UTF-8, never the locale default (cp1252 on Windows): this writes both config.json and
        # the auth record, and neither file's bytes should depend on the machine that wrote it.
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.chmod(tmp_path, stat.S_IRUSR | stat.S_IWUSR)  # 0600, POSIX-only; see the note above.
        os.replace(tmp_path, str(file_path))
    except Exception:
        os.unlink(tmp_path)
        raise


def save_config(config: Config, config_dir: str = DEFAULT_CONFIG_DIR) -> None:
    """Save config to disk."""
    dir_path = _ensure_dir(config_dir)
    file_path = dir_path / "config.json"
    atomic_write(file_path, config.model_dump_json(indent=2))


def config_repair_lines(exc: Exception) -> list[str]:
    """Operator-facing repair guidance for a config the server cannot load.

    Written for stderr, one line per entry, by every entry point that can
    hit an unloadable config: the server's ``main`` (before the transport
    starts), its lifespan backstop, and the CLI. No traceback — every line
    is something the operator can act on.
    """
    lines: list[str]
    if isinstance(exc, ValidationError):
        lines = ["The config file is invalid — the server cannot start:"]
        for e in exc.errors():
            field = ".".join(str(p) for p in e["loc"]) or "config"
            lines.append(f"  {field}: {e['msg']}")
        lines.append(f"Fix {DEFAULT_CONFIG_DIR}/config.json and restart the server.")
    elif isinstance(exc, PermissionError):
        # load_config refuses a symlinked config with PermissionError.
        lines = [
            f"Cannot load the config file — the server cannot start: {exc}",
            "A symlinked config.json is refused on purpose: replace it with "
            "a real file, then restart the server.",
        ]
    elif isinstance(exc, UnicodeDecodeError):
        # A file that is neither UTF-8 nor a valid config in this machine's code
        # page: UTF-16 (what Windows PowerShell 5.1's `>` and Out-File write),
        # bytes that are not text at all, or a code-page file that also fails
        # validation, whose value error shows once it is re-saved. Checked
        # before the generic branch, which it would otherwise fall into as a
        # ValueError.
        lines = [
            f"Cannot load the config file — the server cannot start: {exc}",
            "config.json must be saved as UTF-8 (a byte-order mark is accepted). "
            f"Re-save it as UTF-8 in {DEFAULT_CONFIG_DIR}, then restart the server.",
        ]
    else:
        # OSError: unreadable file or directory, chmod-protected path.
        # ValueError: a settings path that exists but is not a directory.
        lines = [
            f"Cannot load the config file — the server cannot start: {exc}",
            "Check the file and its directory are readable and owned by you, "
            f"in {DEFAULT_CONFIG_DIR}, then restart the server.",
        ]
    return lines


def _refuse_non_directory(config_dir: str) -> None:
    """Refuse a settings path that exists but is a file.

    Surfaced here — at load, as a ``ValueError`` the server and CLI already
    translate into a clean exit with the repair — instead of later, when
    creating the settings directory would fail as a confusing
    ``FileExistsError`` after part of a flow already ran.
    """
    if os.path.exists(config_dir) and not os.path.isdir(config_dir):
        raise ValueError(
            f"The settings path exists but is not a directory: {config_dir}. "
            f"Move or remove it, or point {CONFIG_DIR_ENV} at a directory "
            "(one is created if missing), then restart."
        )


def load_config(config_dir: str = DEFAULT_CONFIG_DIR) -> Config:
    """Load config from disk. Returns defaults if no config file exists."""
    _refuse_non_directory(config_dir)
    file_path = Path(config_dir) / "config.json"

    if not file_path.exists():
        return Config()

    if file_path.is_symlink():
        raise PermissionError(f"Refusing to load symlinked config: {file_path}")

    # Windows cannot represent 0o600, so the mode never reads back as 0o600 there and this
    # check re-chmodded the file on every single load without ever converging. Skipped where
    # the bits do not apply. That carries one real behaviour change: a config.json the user
    # has marked read-only used to have that attribute cleared here — chmod on Windows does
    # honour read-only, so 0o444 became 0o666 — and it now stays read-only. See the note
    # above _ensure_dir.
    if sys.platform != "win32":
        mode = file_path.stat().st_mode & 0o777
        if mode != 0o600:
            file_path.chmod(0o600)

    raw = file_path.read_bytes()
    try:
        # config.json is written by hand and copied between machines, so it is UTF-8 whatever
        # this machine's locale is (#99). utf-8-sig also accepts the byte-order mark that
        # Windows PowerShell 5.1's `Set-Content -Encoding utf8` writes.
        data = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        legacy = _load_legacy_locale_config(raw, file_path)
        if legacy is None:
            raise
        return legacy
    return Config.model_validate_json(data)


def _load_legacy_locale_config(raw: bytes, file_path: Path) -> Config | None:
    """Read a config.json saved in this machine's code page, as releases before #99 did.

    Those releases decoded the file in the locale encoding, so a config written in the ANSI
    code page — Windows PowerShell 5.1's `Set-Content` default — with an accented value
    loaded correctly. Refusing it now would stop a working install from starting, so it keeps
    loading, with a warning asking for a re-save as UTF-8.

    A locale fallback for an otherwise valid configuration, nothing more. It runs only on
    bytes that are not UTF-8, and it accepts only a result that validates, which is what
    turns everything else away without a case of its own. Where the locale is UTF-8, the
    decode fails again. UTF-16, which is what Windows PowerShell 5.1's `>` and `Out-File`
    write, decodes as cp1252 byte for byte, but its NULs can never parse as JSON. Validating
    proves the text is a usable config, not which encoding wrote it. A legacy file whose
    bytes also happen to be valid UTF-8 never reaches this; it is read as UTF-8, which is the
    inherent limit of a UTF-8 file format. None means "no rescue": the caller re-raises the
    UTF-8 error, whose remedy names the fix.
    """
    # Exactly what Path.read_text() used before #99: locale.getencoding() outside UTF-8 mode,
    # UTF-8 inside it.
    encoding = locale.getpreferredencoding(False)
    try:
        config = Config.model_validate_json(raw.decode(encoding))
    except (LookupError, UnicodeDecodeError, ValidationError):  # LookupError: unknown codec
        return None
    logger.warning(
        "%s is not UTF-8; read it in this machine's code page (%s), as earlier releases did. "
        "Re-save it as UTF-8 so it reads the same on every machine.",
        file_path,
        encoding,
    )
    return config
