"""Config file management for outlook-mcp."""

import logging
import os
import stat
import tempfile
from pathlib import Path

from pydantic import BaseModel, Field, field_validator, model_validator

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


class Config(BaseModel):
    """Outlook MCP server configuration."""

    client_id: str | None = Field(default=None, description="Azure AD app client ID (BYOID)")
    tenant_id: str = Field(default=DEFAULT_TENANT_ID)
    read_only: bool = Field(default=False)
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
        """
        if not isinstance(data, dict):
            return data
        for key in data:
            if key in _LEGACY_KEYS:
                logger.warning(
                    "Config key %r ignored: %s.", key, _LEGACY_KEYS[key]
                )
            elif key not in cls.model_fields:
                logger.warning(
                    "Unknown config key %r ignored — supported keys: %s.",
                    key,
                    ", ".join(sorted(cls.model_fields)),
                )
        return data

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


def _ensure_dir(dir_path: str) -> Path:
    """Create config directory with 0700 permissions."""
    path = Path(dir_path)
    path.mkdir(parents=True, exist_ok=True)
    path.chmod(0o700)
    return path


def atomic_write(file_path: Path, data: str) -> None:
    """Write file atomically with fsync, set 0600 permissions."""
    dir_path = file_path.parent
    fd, tmp_path = tempfile.mkstemp(dir=str(dir_path), suffix=".tmp")
    try:
        with os.fdopen(fd, "w") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.chmod(tmp_path, stat.S_IRUSR | stat.S_IWUSR)  # 0600
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
    from pydantic import ValidationError

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
    else:
        # OSError: unreadable file or directory, chmod-protected path.
        # ValueError: non-UTF-8 bytes in the file, or a settings path that
        # exists but is not a directory.
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

    mode = file_path.stat().st_mode & 0o777
    if mode != 0o600:
        file_path.chmod(0o600)

    data = file_path.read_text()
    return Config.model_validate_json(data)
