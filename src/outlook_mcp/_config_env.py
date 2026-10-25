"""Optional environment input constrained by the valid file's security policy."""

from __future__ import annotations

import json
import ntpath
import os
from collections.abc import Mapping
from pathlib import Path

from pydantic import ValidationError

from outlook_mcp.config import Config, EnvironmentConfigError, _misspelt_safety_key, _squash
from outlook_mcp.permissions import VALID_CATEGORIES

PREFIX = "OUTLOOK_MCP_"
PLAIN_FIELDS = ("client_id", "tenant_id", "timezone")
SECURITY_FIELDS = (
    "read_only",
    "read_only_consent",
    "allow_categories",
    "attachments_dir",
    "allow_unencrypted_token_cache",
)
ENV_FIELDS = {PREFIX + field.upper(): field for field in (*PLAIN_FIELDS, *SECURITY_FIELDS)}
BOOL_FIELDS = {"read_only", "read_only_consent", "allow_unencrypted_token_cache"}
_WINDOWS = os.name == "nt"


def _refuse(variable: str, reason: str) -> EnvironmentConfigError:
    return EnvironmentConfigError(f"{variable!r}: {reason}")


def _parse(variable: str, field: str, value: str) -> object:
    if field in BOOL_FIELDS:
        value = value.strip().lower()
        if value not in ("true", "false"):
            raise _refuse(variable, "use true or false explicitly.")
        return value == "true"
    if field == "allow_categories":
        try:
            categories = json.loads(value)
        except (ValueError, RecursionError):
            raise _refuse(variable, "use a JSON array of permission category names.") from None
        if not isinstance(categories, list) or not all(isinstance(c, str) for c in categories):
            raise _refuse(variable, "use a JSON array of permission category names.")
        return categories
    value = value.strip()
    if not value or any(ord(c) < 32 or ord(c) == 127 for c in value):
        raise _refuse(variable, "use a nonempty value on one line.")
    return value


def _narrow_attachment_root(value: str, baseline: str) -> tuple[str, str]:
    """Resolve without creating directories; refuse wider roots and symlink escapes.

    A new Windows UNC/device location is rejected lexically before resolution can
    contact a server. An operator-approved UNC root can only name its descendants.
    """
    variable = PREFIX + "ATTACHMENTS_DIR"
    value, baseline = os.path.expanduser(value), os.path.expanduser(baseline)
    if _WINDOWS:
        root_drive = ntpath.splitdrive(ntpath.abspath(baseline))[0].casefold()
        target_drive = ntpath.splitdrive(ntpath.abspath(value))[0].casefold()
        if target_drive != root_drive:
            raise _refuse(variable, "an override on a different drive or share is refused.")
    if _WINDOWS and len(value) >= 2 and all(c in "\\/" for c in value[:2]):
        root = ntpath.normcase(ntpath.abspath(baseline))
        target = ntpath.normcase(ntpath.abspath(value))
        try:
            within = (
                len(root) >= 2
                and all(c in "\\/" for c in root[:2])
                and ntpath.commonpath((root, target)) == root
            )
        except ValueError:
            within = False
        if not within:
            raise _refuse(
                variable, "a network or device path outside the configured root is refused."
            )
    try:
        root = Path(baseline).resolve()
        candidate = Path(value).resolve()
        if not candidate.is_relative_to(root):
            raise _refuse(variable, "the resolved directory must stay inside the configured root.")
    except (OSError, RuntimeError, ValueError) as exc:
        if isinstance(exc, EnvironmentConfigError):
            raise
        raise _refuse(variable, "the directory could not be resolved safely.") from None
    # Keep the checked absolute root, independent of later working-directory changes.
    return str(candidate), str(root)


def apply_environment(baseline: Config, environ: Mapping[str, str]) -> Config:
    """Validate the file first, then apply only explicit, non-persistent inputs."""
    prefix = _squash(PREFIX)
    for variable in environ:
        normalized = _squash(variable.casefold())
        if normalized.startswith(prefix) and variable not in ENV_FIELDS:
            meant = _misspelt_safety_key(normalized[len(prefix) :], SECURITY_FIELDS)
            if meant:
                raise _refuse(
                    variable, f"a security setting appears misspelt; use {PREFIX + meant.upper()}."
                )
    updates = {
        field: _parse(variable, field, environ[variable])
        for variable, field in ENV_FIELDS.items()
        if variable in environ
    }
    if not updates:
        return baseline
    if (
        "client_id" in updates
        and baseline.client_id
        and updates["client_id"].casefold() != baseline.client_id.casefold()
    ):
        raise _refuse(
            PREFIX + "CLIENT_ID", "cannot change the app registration set by config.json."
        )
    # Config remains the only settings model and validator for field values and
    # relationships. Environment values cannot rescue an invalid config file.
    try:
        effective = Config.model_validate({**baseline.model_dump(), **updates})
    except ValidationError as exc:
        fields = {e["loc"][0] for e in exc.errors() if e["loc"]}
        variables = [PREFIX + f.upper() for f in updates if not fields or f in fields]
        reason = "invalid configuration value."
        if updates.get("read_only_consent", baseline.read_only_consent) and not updates.get(
            "read_only", baseline.read_only
        ):
            reason = "read_only_consent requires read_only to be true."
        elif "allow_categories" in fields:
            reason = "use valid permission categories: " + ", ".join(sorted(VALID_CATEGORIES)) + "."
        raise _refuse(", ".join(variables), reason) from None
    effective._attachments_file_root = baseline._attachments_file_root
    for field in ("read_only", "read_only_consent"):
        if field in updates and getattr(baseline, field) and not getattr(effective, field):
            raise _refuse(PREFIX + field.upper(), "cannot disable a restriction from config.json.")
    if (
        "allow_unencrypted_token_cache" in updates
        and not baseline.allow_unencrypted_token_cache
        and effective.allow_unencrypted_token_cache
    ):
        raise _refuse(
            PREFIX + "ALLOW_UNENCRYPTED_TOKEN_CACHE",
            "cannot enable unencrypted storage unless config.json already permits it.",
        )
    if "allow_categories" in updates and baseline.allow_categories:
        if not effective.allow_categories or not set(effective.allow_categories) <= set(
            baseline.allow_categories
        ):
            raise _refuse(
                PREFIX + "ALLOW_CATEGORIES",
                "cannot widen the file's allowed categories; an empty list means all categories.",
            )
    if "attachments_dir" in updates:
        effective.attachments_dir, root = _narrow_attachment_root(
            effective.attachments_dir, baseline.attachments_dir
        )
        effective._attachments_file_root = baseline._attachments_file_root or root
    if "timezone" in updates:
        # Reuse the existing timezone resolver, without changing file-only loading.
        from outlook_mcp.validation import resolve_timezone

        try:
            resolve_timezone(effective.timezone)
        except ValueError:
            raise _refuse(
                PREFIX + "TIMEZONE", "use a valid IANA timezone such as UTC or Europe/London."
            ) from None
    return effective
