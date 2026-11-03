"""Config file management for outlook-mcp."""

import difflib
import json
import locale
import logging
import ntpath
import os
import re
import stat
import sys
import tempfile
from pathlib import Path

from pydantic import (
    BaseModel,
    Field,
    PrivateAttr,
    ValidationError,
    field_validator,
    model_validator,
)
from pydantic_core import InitErrorDetails, PydanticCustomError

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


_WINDOWS = os.name == "nt"


def _is_network_path(path: str) -> bool:
    r"""On Windows, a path that begins with two separators: ``\\host\share``, ``\\?\``, ``\\.\``.

    Resolving or opening one connects to its host, so such a path is only ever
    compared by its text here.
    """
    return _WINDOWS and len(path) >= 2 and path[0] in "\\/" and path[1] in "\\/"


# `\\?\` or `\\.\` (either separator) in front of a drive letter and its colon.
_DRIVE_DEVICE_PREFIX = re.compile(r"[\\/]{2}[?.][\\/](?=[A-Za-z]:)")


def _without_device_prefix(path: str) -> str:
    r"""On Windows, ``C:\x`` for ``\\?\C:\x`` or ``\\.\C:\x``: all three name one folder.

    Any other path that begins with two separators, ``\\host\share`` or
    ``\\?\UNC\host\share``, keeps its prefix and stays a network path.

    So a volume spelling, ``\\?\Volume{GUID}\...``, and an admin share on this
    machine, ``\\localhost\C$\...``, are judged by their text only, by design: they
    come from the operator's own config, and resolving them to the folder they name
    is the lookup this rule avoids.
    """
    if _WINDOWS and _DRIVE_DEVICE_PREFIX.match(path):
        return path[4:]
    return path


def _comparable(path: str) -> str:
    """``path`` resolved (symlinks included) and case-normalised, without creating anything.

    A device spelling of a drive path is resolved as that drive path.
    """
    path = _without_device_prefix(path)
    if _is_network_path(path):
        return ntpath.normcase(ntpath.normpath(path))
    return os.path.normcase(os.path.realpath(path))


# macOS shows the whole Data volume, home included, a second time here.
_MACOS_DATA_VOLUME = "/System/Volumes/Data"


def _is_same_or_inside(path: str, folder: str) -> bool:
    """Whether ``path`` is ``folder`` or lies inside it; both come from ``_comparable``.

    Compared by text first, then by identity: on a case-insensitive filesystem
    (the macOS default) two spellings can name one directory, and ``normcase``
    folds case only on Windows. So each existing directory on the way up from
    ``path`` is also compared with ``folder`` itself, and on macOS so is each
    one on the way up from ``path`` spelt through the Data volume
    (``/System/Volumes/Data`` and the folders above it hold home too).

    A folder can also hold ``path`` without being above either spelling: a
    second view of a volume, or a link inside it. So the rest of ``path`` below
    each of its parents is looked up under ``folder`` too, and counts if it is
    ``path`` itself. Only on ``path``'s own filesystem: on an automounter's
    root (``/net``, or ``/home`` on macOS) a lookup under the folder is a
    directory-service or network request, and loading the config makes none.
    """
    try:
        if os.path.commonpath((path, folder)) == folder:
            return True
    except ValueError:  # different drives, or a share beside a local path
        return False
    if _is_network_path(path) or _is_network_path(folder):
        return False
    try:
        folder_stat = os.stat(folder)
    except OSError:
        return False  # a folder that does not exist contains nothing that does
    if not folder_stat.st_ino:
        return False  # this filesystem gives no identity to compare (some Windows ones)
    if _walks_up_to(path, folder_stat):
        return True
    try:
        path_stat = os.stat(path)
    except OSError:
        return False  # not created yet: there is nothing more to find
    if not path_stat.st_ino:
        return False
    again = _spelt_through_the_data_volume(path, path_stat)
    if again is not None and _walks_up_to(again, folder_stat):
        return True
    if folder_stat.st_dev != path_stat.st_dev:
        return False
    return _found_under(folder, path, path_stat)


def _walks_up_to(path: str, folder_stat: os.stat_result) -> bool:
    """Whether ``path``, or a directory on the way up from it, is the folder ``folder_stat``."""
    current = path
    while True:
        try:
            if os.path.samestat(os.stat(current), folder_stat):
                return True
        except OSError:
            pass  # not created yet; its parent may still be the folder
        parent = os.path.dirname(current)
        if parent == current:
            return False
        current = parent


def _spelt_through_the_data_volume(path: str, path_stat: os.stat_result) -> str | None:
    """On macOS, ``path`` spelt through the Data volume, when that names ``path`` itself.

    Anywhere else that spelling does not exist, and ``None`` is returned.
    """
    if not path.startswith("/") or (path + "/").startswith(_MACOS_DATA_VOLUME + "/"):
        return None
    again = _MACOS_DATA_VOLUME + path
    try:
        return again if os.path.samestat(os.stat(again), path_stat) else None
    except OSError:
        return None


def _found_under(folder: str, path: str, path_stat: os.stat_result) -> bool:
    """Whether the rest of ``path`` below one of its parents, under ``folder``, is ``path``."""
    current = path
    while True:
        parent = os.path.dirname(current)
        if parent == current:
            return False
        current = parent
        again = os.path.join(folder, os.path.relpath(path, current))
        try:
            if os.path.samestat(os.stat(again), path_stat):
                return True
        except OSError:
            pass  # not under the folder by this route


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
_SAFETY_KEYS = ("read_only", "allow_categories", "read_only_consent", "expected_username")


def _squash(key: str) -> str:
    """`readOnly`, `read-only`, `READ_ONLY` and `readonly` all squash to `readonly`."""
    return "".join(ch for ch in key.lower() if ch.isalnum())


def _misspelt_safety_key(key: str, keys: tuple[str, ...] = _SAFETY_KEYS) -> str | None:
    """The safety key an unknown `key` plainly means, or None.

    A different spelling of the same words, or a one-or-two-letter slip
    (`difflib` ratio 0.8 or better). Deliberately narrow: only these keys,
    so every other unknown key keeps loading with a warning in file configuration.
    """
    for safety_key in keys:
        if _squash(key) == _squash(safety_key):
            return safety_key
    close = difflib.get_close_matches(key.lower(), keys, n=1, cutoff=0.8)
    return close[0] if close else None


class EnvironmentConfigError(ValueError):
    """An explicitly supplied environment value is invalid or would loosen policy."""


class Config(BaseModel):
    """Outlook MCP server configuration."""

    # A narrowed environment directory must stay inside the file's resolved root
    # at use time too: a child can become a symlink after configuration loads.
    # This is runtime policy metadata, never another persisted configuration key.
    _attachments_file_root: str | None = PrivateAttr(default=None)

    client_id: str | None = Field(default=None, description="Azure AD app client ID (BYOID)")
    tenant_id: str = Field(default=DEFAULT_TENANT_ID)
    expected_username: str | None = Field(
        default=None,
        description=(
            "Optional sign-in username for this instance. Refuse a saved or new "
            "sign-in for another username; case-insensitive, without alias resolution."
        ),
    )
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
    mailbox_settings_write: bool = Field(
        default=False,
        description=(
            "Ask Microsoft for `MailboxSettings.ReadWrite` at sign-in instead of "
            "`MailboxSettings.Read`, which is what the inbox-rule write tools need. "
            "Off by default because it widens the token to every mailbox setting "
            "(automatic replies, working hours, time zone), not only rules. Adding "
            "the permission in the Azure portal is not enough on its own: consent is "
            "granted per *requested* scope, so the sign-in has to name it and the "
            "account has to consent again."
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
        if not _warn_about_unknown_keys:
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

    @field_validator("expected_username")
    @classmethod
    def _validate_expected_username(cls, value: str | None) -> str | None:
        if value is None:
            return None
        value = value.strip()
        if not value or any(ord(char) < 32 or ord(char) == 127 for char in value):
            raise ValueError("expected_username must be a non-empty sign-in username on one line.")
        return value

    @model_validator(mode="after")
    def _mailbox_settings_write_needs_a_write_consent(self) -> "Config":
        """Refuse a scope request that the sign-in would never make.

        `read_only_consent` asks for the read scopes only, so `get_scopes` returns
        the read list whatever `mailbox_settings_write` says. Loaded together, the
        key would be silently ignored and every rule write would fail with a 403
        that names none of it — the same dead end as documenting a permission the
        code never requests.
        """
        if self.mailbox_settings_write and self.read_only_consent:
            raise ValueError(
                "mailbox_settings_write: true cannot be combined with read_only_consent: "
                "true. That consent asks for the read scopes only, so MailboxSettings."
                "ReadWrite would never be requested and every inbox-rule write would "
                "fail. Drop one: read_only_consent for a registration that must never "
                "write, mailbox_settings_write for one that should manage rules."
            )
        return self

    @field_validator("attachments_dir")
    @classmethod
    def _validate_attachments_dir(cls, value: str) -> str:
        """Refuse a folder that would hand the attachment tools more than attachments.

        Those tools read and write every file under ``attachments_dir``. So it
        must not be the home directory or a folder that contains it, nor the
        settings directory (config.json and the sign-in record) or a folder that
        contains it; a subfolder of either is fine. It must be absolute after
        ``~`` is expanded: a relative path names a different folder in every
        working directory the server starts in. Checked on the resolved path, so
        a symlink to a refused folder is refused too. The stored value is kept as
        written. ``load_config`` checks the default folder the same way.
        """
        # A device spelling of a drive path is judged as that drive path.
        expanded = _without_device_prefix(os.path.expanduser(value))
        if _WINDOWS:
            # A drive or a share, and a root: `\x` (which Python before 3.13 calls
            # absolute) and `C:x` both still depend on the current drive or directory.
            absolute = ntpath.isabs(expanded) and bool(ntpath.splitdrive(expanded)[0])
        else:
            absolute = os.path.isabs(expanded)
        if not absolute or "\x00" in expanded:
            raise ValueError(
                "attachments_dir must be an absolute folder path (a leading ~ is "
                "expanded). A relative path names a different folder depending on "
                "where the server starts."
            )
        folder = _comparable(expanded)
        if _is_same_or_inside(_comparable(os.path.expanduser("~")), folder):
            raise ValueError(
                "attachments_dir must not be the home directory or a folder that "
                "contains it: the attachment tools can read and write every file "
                "under it. Use a folder of its own; one inside the home directory "
                "is fine."
            )
        if _is_same_or_inside(_comparable(DEFAULT_CONFIG_DIR), folder):
            raise ValueError(
                "attachments_dir must not be the settings directory or a folder that "
                "contains it: the attachment tools could then read and replace the "
                "settings and the sign-in record. Use a folder of its own; the default "
                "attachments folder inside the settings directory is fine."
            )
        return value

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
    if isinstance(exc, EnvironmentConfigError):
        lines = [
            f"Cannot load environment configuration — the server cannot start: {exc}",
            "Fix the named variable in the launch environment and restart the server.",
        ]
    elif isinstance(exc, ValidationError):
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


class _RepeatedKeyError(Exception):
    """One JSON object in config.json names the same key twice."""

    def __init__(self, key: str):
        super().__init__(key)
        self.key = key


def _refuse_repeated_key(pairs: list[tuple[str, object]]) -> dict[str, object]:
    seen: set[str] = set()
    for key, _ in pairs:
        if key in seen:
            raise _RepeatedKeyError(key)
        seen.add(key)
    return dict(pairs)


def _validate_config_json(text: str) -> Config:
    """Validate config.json's text, after refusing any object that repeats a key.

    A JSON parser keeps one of the two values without a word, so such a file does
    not say which one was meant, and the one that silently loses may be the
    restriction. Repeats are found by parsing, at any depth, so a key spelt with
    an escape counts as the same key. Text this parser cannot read is left to
    ``model_validate_json``, which reports it exactly as before. The refusal is a
    ``ValidationError``, like any other invalid config, and names the key but
    neither value.
    """
    try:
        json.loads(text, object_pairs_hook=_refuse_repeated_key)
    except _RepeatedKeyError as repeated:
        raise ValidationError.from_exception_data(
            Config.__name__,
            [
                InitErrorDetails(
                    type=PydanticCustomError(
                        "repeated_key",
                        "the key {key} appears more than once in one JSON object, so the "
                        "file does not say which value is meant. Keep it once.",
                        {"key": repr(repeated.key)},
                    ),
                    loc=(),
                    input=repeated.key,
                )
            ],
        ) from None
    except (ValueError, RecursionError):
        pass  # not JSON this parser reads; model_validate_json reports it below
    return Config.model_validate_json(text)


# Silences the accept-and-warn pass while a caller reads one field for an
# internal decision (toolset selection at import). Those warnings belong to the
# start-up load; repeating them for an import-time peek printed every one of
# them twice on every start.
_warn_about_unknown_keys = True
# Same idea for warnings that are not about unknown keys — the legacy-encoding
# note, for one. A `quiet` load silences every notice it would print, because a
# peek repeated it once more on every start.
_quiet_config_load = False


def load_config(
    config_dir: str = DEFAULT_CONFIG_DIR,
    *,
    warn_unknown: bool = True,
    quiet: bool = False,
) -> Config:
    """Load valid file/default settings, then non-persistent environment inputs.

    ``warn_unknown=False`` suppresses the unknown/legacy-key warnings for this
    call, so an internal read does not duplicate the ones the start-up load
    prints.
    """
    global _warn_about_unknown_keys, _quiet_config_load
    previous = _warn_about_unknown_keys
    previous_quiet = _quiet_config_load
    _warn_about_unknown_keys = warn_unknown and not quiet
    _quiet_config_load = quiet
    try:
        baseline = _load_file_config(config_dir)
        if "attachments_dir" not in baseline.model_fields_set:
            # The default folder is checked as if the file had set it: it can be a
            # symlink to somewhere refused. Here rather than on every Config(), so
            # the server's fail-safe substitute after a refused load still builds.
            Config.model_validate({"attachments_dir": baseline.attachments_dir})
        from outlook_mcp._config_env import apply_environment

        return apply_environment(baseline, os.environ)
    finally:
        _warn_about_unknown_keys = previous
        _quiet_config_load = previous_quiet


def _load_file_config(config_dir: str) -> Config:
    """The existing file loader; environment input never bypasses its refusals."""
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
    return _validate_config_json(data)


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
        # A repeated key is refused here too, as a ValidationError: no rescue.
        config = _validate_config_json(raw.decode(encoding))
    except (LookupError, UnicodeDecodeError, ValidationError):  # LookupError: unknown codec
        return None
    if not _quiet_config_load:
        logger.warning(
            "%s is not UTF-8; read it in this machine's code page (%s), as earlier releases did. "
            "Re-save it as UTF-8 so it reads the same on every machine.",
            file_path,
            encoding,
        )
    return config
