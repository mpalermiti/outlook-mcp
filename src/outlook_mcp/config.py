"""Config file management for outlook-mcp."""

from __future__ import annotations

import logging
import os
import re
import stat
import tempfile
from pathlib import Path

from pydantic import BaseModel, Field, field_validator, model_validator

from outlook_mcp.permissions import VALID_CATEGORIES

DEFAULT_TENANT_ID = "consumers"
DEFAULT_CONFIG_DIR = os.path.expanduser("~/.outlook-mcp")

logger = logging.getLogger(__name__)

# Capabilities an account routing decision can be made for. Derived from the
# toolset groups: mail-centric groups (drafts/attachments/folders/admin/digest)
# fold into "mail"; "account" tools follow the active account, not a capability.
ROUTING_CAPABILITIES = {"mail", "calendar", "contacts", "todo"}

# The canonical account-name shape: short, filesystem- and CLI-safe.
# fullmatch, not match-with-$ — `re.match("...$", "net\n")` accepts the
# trailing newline, and a name like that must not reach a filename.
_ACCOUNT_NAME_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]{0,31}")

# Names that cannot serve as `auth_record-<name>.json` on some filesystem.
# Anything here is a hard error, not a compat warning: the account could never
# authenticate, whatever version wrote the config.
_UNSAFE_NAME_PATTERNS = re.compile(r"[/\\\x00-\x1f]")
# Windows reserved names also match the canonical regex (CON, NUL, COM1…),
# so they need their own check ahead of it.
_WINDOWS_RESERVED = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def _windows_reserved_member(value: str) -> bool:
    stem = value.split(".")[0].upper()
    return stem in _WINDOWS_RESERVED


def _check_account_name(value: str) -> str:
    """Validate an account name without breaking configs older versions took.

    1.21 and earlier accepted any string as an account name, so an upgrade
    must not kill a working install over an unusual-but-harmless name
    (review of #61). Only names that cannot work at all — path separators,
    control characters, reserved filenames, absurd length — are rejected;
    the merely unusual pass with a warning.
    """
    if _windows_reserved_member(value):
        raise ValueError(
            f"Account name {value[:30]!r} is a reserved filename on Windows — "
            "auth_record-{name}.json could not be written there. Rename the "
            "account in ~/.outlook-mcp/config.json."
        )
    if _ACCOUNT_NAME_RE.fullmatch(value):
        return value
    if not value or len(value) > 64 or _UNSAFE_NAME_PATTERNS.search(value) or value in {".", ".."}:
        raise ValueError(
            f"Account name {value[:30]!r} is not usable as a filename "
            "(path separators, control characters, or over 64 chars). "
            "Rename the account in ~/.outlook-mcp/config.json."
        )
    logger.warning(
        "Account name %r does not follow the recommended shape (1-32 chars "
        "of letters, digits, '-' or '_'); accepted for compatibility with "
        "configs written by older versions.",
        value,
    )
    return value


class AccountConfig(BaseModel):
    """Configuration for a single account."""

    name: str
    client_id: str
    tenant_id: str = DEFAULT_TENANT_ID

    @field_validator("name")
    @classmethod
    def _validate_name(cls, value: str) -> str:
        """Account names become auth-record filenames and CLI arguments.

        See _check_account_name: unusual names warn, unusable ones fail —
        an upgrade must not kill a working install.
        """
        return _check_account_name(value)


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
    timezone: str = Field(default="UTC", description="IANA timezone for relative date computations")
    attachments_dir: str = Field(
        default="~/.outlook-mcp/attachments",
        description=(
            "The only directory the attachment tools may read from or write to. "
            "Point it somewhere else to widen the surface; every path an agent "
            "supplies is resolved and must land inside it."
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
    accounts: list[AccountConfig] = Field(default_factory=list)
    default_account: str | None = Field(
        default=None,
        description="Account serving capabilities without an explicit routing.",
    )
    capability_accounts: dict[str, str] = Field(
        default_factory=dict,
        description=(
            'Per-capability account routing, e.g. {"mail": "net", "todo": "neko"}. '
            "A tool routes to the account named for its capability; capabilities "
            "not listed fall back to default_account."
        ),
    )
    allow_cross_account: bool = Field(
        default=False,
        description=(
            "Master switch for cross-account access. False (default): the agent "
            "sees one merged account and cannot address any other account's "
            "content — outlook_switch_account refuses. True: the agent may "
            "switch routing to query non-default content."
        ),
    )

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

    @model_validator(mode="after")
    def _validate_accounts(self) -> Config:
        """Reconcile account routing — accept-and-warn for legacy shapes.

        1.21 and earlier accepted `accounts` (and `default_account`) without
        any cross-field validation, and `load_config()` runs uncaught in the
        server lifespan and every CLI command — a validator that hard-rejects
        what main accepted makes a working install die at startup, unable even
        to `logout` (review of #61). So every shape an older version could
        have written is accepted with a warning and normalized. Hard errors
        are reserved for fields no older version knew: a `capability_accounts`
        entry naming an unknown account is a typo in *new* config, and
        silently re-routing that capability to the default account is the
        identity substitution this PR exists to prevent.
        """
        if not self.accounts:
            if self.capability_accounts:
                logger.warning(
                    "capability_accounts is set but 'accounts' is empty — "
                    "ignoring the routing. Define the accounts first."
                )
                self.capability_accounts = {}
            if self.default_account is not None:
                logger.warning(
                    "default_account '%s' is set but 'accounts' is empty — ignoring it.",
                    self.default_account,
                )
                self.default_account = None
            return self

        names = [acc.name for acc in self.accounts]
        if len(names) != len(set(names)):
            logger.warning(
                "Duplicate account names in 'accounts': %s — keeping the first of each.",
                names,
            )
            seen: list[str] = []
            deduped: list[AccountConfig] = []
            for acc in self.accounts:
                if acc.name not in seen:
                    seen.append(acc.name)
                    deduped.append(acc)
            self.accounts = deduped
            names = seen

        unknown_caps = [c for c in self.capability_accounts if c not in ROUTING_CAPABILITIES]
        if unknown_caps:
            logger.warning(
                "Unknown capabilities in capability_accounts: %s — dropping "
                "them. Valid capabilities: %s",
                unknown_caps,
                sorted(ROUTING_CAPABILITIES),
            )
            self.capability_accounts = {
                c: a for c, a in self.capability_accounts.items() if c in ROUTING_CAPABILITIES
            }

        # A routing that names a missing account is a typo in new config:
        # honoring it would need a fallback, and falling back to default
        # silently serves the wrong mailbox. Fail loudly instead.
        unknown_accounts = [a for a in self.capability_accounts.values() if a not in names]
        if unknown_accounts:
            raise ValueError(
                f"capability_accounts references unknown accounts: {unknown_accounts}. "
                f"Configured accounts: {names}"
            )

        if self.default_account is None:
            self.default_account = names[0]
        elif self.default_account not in names:
            logger.warning(
                "default_account '%s' is not a configured account (%s) — using '%s' instead.",
                self.default_account,
                names,
                names[0],
            )
            self.default_account = names[0]
        return self


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
