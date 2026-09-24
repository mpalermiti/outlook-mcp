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
DEFAULT_CONFIG_DIR = os.path.expanduser(
    os.environ.get(CONFIG_DIR_ENV) or "~/.outlook-mcp"
)


def _default_attachments_dir() -> str:
    """Attachments live under the settings directory, wherever it was moved to."""
    override = os.environ.get(CONFIG_DIR_ENV)
    if override:
        return os.path.join(override, "attachments")
    return "~/.outlook-mcp/attachments"


# Top-level keys an older release accepted. One process serves one account
# now; a config carrying these still loads, the operator just hears about it.
_LEGACY_KEYS = {
    "accounts": (
        "configuring multiple accounts in one process is no longer supported; "
        "run one server per account and give each its own settings directory "
        f"via the {CONFIG_DIR_ENV} environment variable"
    ),
    "default_account": (
        "configuring multiple accounts in one process is no longer supported; "
        "run one server per account and give each its own settings directory "
        f"via the {CONFIG_DIR_ENV} environment variable"
    ),
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


def _atomic_write(file_path: Path, data: str) -> None:
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
    _atomic_write(file_path, config.model_dump_json(indent=2))


def load_config(config_dir: str = DEFAULT_CONFIG_DIR) -> Config:
    """Load config from disk. Returns defaults if no config file exists."""
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
