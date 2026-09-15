"""Config file management for outlook-mcp."""

import os
import re
import stat
import tempfile
from pathlib import Path

from pydantic import BaseModel, Field, field_validator, model_validator

from outlook_mcp.permissions import VALID_CATEGORIES

DEFAULT_TENANT_ID = "consumers"
DEFAULT_CONFIG_DIR = os.path.expanduser("~/.outlook-mcp")

# Capabilities an account routing decision can be made for. Derived from the
# toolset groups: mail-centric groups (drafts/attachments/folders/admin/digest)
# fold into "mail"; "account" tools follow the active account, not a capability.
ROUTING_CAPABILITIES = {"mail", "calendar", "contacts", "todo"}

_ACCOUNT_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,31}$")


class AccountConfig(BaseModel):
    """Configuration for a single account."""

    name: str
    client_id: str
    tenant_id: str = DEFAULT_TENANT_ID

    @field_validator("name")
    @classmethod
    def _validate_name(cls, value: str) -> str:
        """Account names become keyring/cache filenames and CLI arguments.

        Keep them short and filesystem-safe rather than accepting anything and
        discovering the exotic cases as OSError deep inside msal_extensions.
        """
        if not _ACCOUNT_NAME_RE.match(value):
            raise ValueError(
                f"Account name '{value[:30]}' must be 1-32 chars of letters, "
                "digits, '-' or '_', starting with a letter or digit."
            )
        return value


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
        description="Fallback account for capabilities without an explicit routing.",
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
    allow_aggregate: bool = Field(
        default=False,
        description=(
            "Master switch for the aggregated read tools (outlook_list_inbox_all, "
            "outlook_list_events_all, outlook_list_tasks_all). False (default): "
            "they refuse. True: they fan out concurrently to every authenticated "
            "account and return one merged, per-item account-tagged listing. "
            "Orthogonal to allow_cross_account: cross gates deliberately "
            "switching routing; aggregate gates bulk cross-account reads."
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
    def _validate_accounts(self) -> "Config":
        """Cross-field account validation, at load time where the fix is cheap.

        - capability_accounts keys must be real capabilities and values must
          name configured accounts — a typo here would otherwise surface as a
          confusing AuthRequiredError on the first tool call.
        - default_account / capability values must exist.
        - default_account defaults to the first configured account.
        """
        if not self.accounts:
            if self.capability_accounts:
                raise ValueError(
                    "capability_accounts is set but 'accounts' is empty — define "
                    "the accounts first."
                )
            if self.default_account is not None:
                raise ValueError(
                    f"default_account '{self.default_account}' is set but 'accounts' is empty."
                )
            return self

        names = [acc.name for acc in self.accounts]
        if len(names) != len(set(names)):
            raise ValueError(f"Duplicate account names in 'accounts': {names}")

        unknown_caps = [c for c in self.capability_accounts if c not in ROUTING_CAPABILITIES]
        if unknown_caps:
            raise ValueError(
                f"Unknown capabilities in capability_accounts: {unknown_caps}. "
                f"Valid capabilities: {sorted(ROUTING_CAPABILITIES)}"
            )
        unknown_accounts = [a for a in self.capability_accounts.values() if a not in names]
        if unknown_accounts:
            raise ValueError(
                f"capability_accounts references unknown accounts: {unknown_accounts}. "
                f"Configured accounts: {names}"
            )
        if self.default_account is None:
            self.default_account = names[0]
        elif self.default_account not in names:
            raise ValueError(
                f"default_account '{self.default_account}' not in configured accounts: {names}"
            )
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
