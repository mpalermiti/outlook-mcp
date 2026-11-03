"""A download lands in a new file: never over an existing one, never under a hidden name.

A download counts as a read in the permission model, so ``read_only`` does not stop it,
and its ``save_path`` comes from the model. So the landing rules hold for every call:
a name already taken in ``attachments_dir`` (a file the user staged to send, or a
symlink, dangling or not) is refused, and so is a name or folder starting with a dot.
Both are refused before Graph is asked, and the write itself never replaces a file
that appears while Graph answers.

The mail and To Do downloads share the landing code, so every check runs through both.
"""

import errno
import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.config import Config
from outlook_mcp.errors import LEAVE_SETTINGS_TO_THE_USER
from outlook_mcp.tools import mail_attachments
from outlook_mcp.tools.mail_attachments import (
    _validated_download_target,
    download_attachment,
    write_download,
)
from outlook_mcp.tools.todo_attachments import download_task_attachment

CONTENT = b"%PDF-1.7 downloaded bytes"
STAGED = b"the file the user staged to send"

no_symlinks_on_windows = pytest.mark.skipif(
    sys.platform == "win32", reason="symlinks need privileges on Windows"
)
posix_modes_only = pytest.mark.skipif(
    sys.platform == "win32",
    reason="os.chmod on Windows honours only the read-only attribute (#85).",
)


def _attachment():
    att = MagicMock()
    att.id = "att1"
    att.name = "report.pdf"
    att.size = len(CONTENT)
    att.content_type = "application/pdf"
    att.content_bytes = CONTENT
    return att


def _graph_get(on_get):
    async def _get():
        if on_get is not None:
            on_get()
        return _attachment()

    return AsyncMock(side_effect=_get)


def _mail(attachments_dir, on_get=None):
    client = MagicMock()
    get = _graph_get(on_get)
    client.me.messages.by_message_id.return_value.attachments.by_attachment_id.return_value.get = (
        get
    )

    async def download(save_path):
        return await download_attachment(
            client,
            message_id="AAMkAG123=",
            attachment_id="att1",
            save_path=save_path,
            config=Config(attachments_dir=str(attachments_dir)),
        )

    return download, get


def _todo(attachments_dir, on_get=None):
    client = MagicMock()
    get = _graph_get(on_get)
    task = client.me.todo.lists.by_todo_task_list_id.return_value.tasks.by_todo_task_id
    task.return_value.attachments.by_attachment_base_id.return_value.get = get

    async def download(save_path):
        return await download_task_attachment(
            client,
            task_id="task1",
            attachment_id="att1",
            save_path=save_path,
            list_id="list1",
            config=Config(attachments_dir=str(attachments_dir)),
        )

    return download, get


@pytest.fixture(params=[_mail, _todo], ids=["mail", "todo"])
def tool(request):
    """Builds (download(save_path), the Graph GET mock) for one of the two download tools."""
    return request.param


@pytest.fixture
def base(tmp_path):
    attachments = tmp_path / "attachments"
    attachments.mkdir()
    return attachments


def _already_exists(name: str) -> str:
    return (
        f"A file named '{name}' already exists in attachments_dir, and downloads never "
        "replace a file. Pass a different save_path."
    )


def _hidden(component: str) -> str:
    return (
        "Downloads cannot create hidden files or write inside hidden folders "
        f"('{component}'). Pass a save_path without a leading dot."
    )


def _leftover_temp_files(directory: Path) -> list[str]:
    """Anything in ``directory`` besides the download's own name: a temp file left behind."""
    return [p.name for p in directory.iterdir() if p.name != "report.pdf"]


# ── An existing name is refused before Graph is asked ─────────────────────────


@pytest.mark.parametrize("where", ["bare", "absolute", "subfolder"])
async def test_an_existing_file_is_refused_before_graph_is_asked(tool, base, where):
    (base / "sub").mkdir()
    relative = Path("sub", "report.pdf") if where == "subfolder" else Path("report.pdf")
    staged = base / relative
    staged.write_bytes(STAGED)
    save_path = str(staged) if where == "absolute" else str(relative)
    download, get = tool(base)

    with pytest.raises(ValueError) as error:
        await download(save_path)

    assert str(error.value) == _already_exists(str(relative))
    assert staged.read_bytes() == STAGED
    get.assert_not_awaited()


@no_symlinks_on_windows
async def test_a_dangling_symlink_at_the_target_is_refused(tool, base):
    link = base / "report.pdf"
    link.symlink_to(base / "not-there-yet.pdf")
    download, get = tool(base)

    with pytest.raises(ValueError) as error:
        await download("report.pdf")

    assert str(error.value) == _already_exists("report.pdf")
    assert link.is_symlink()
    assert not (base / "not-there-yet.pdf").exists()
    get.assert_not_awaited()


@no_symlinks_on_windows
async def test_a_symlink_to_an_existing_file_is_refused(tool, base):
    real = base / "real.pdf"
    real.write_bytes(STAGED)
    (base / "report.pdf").symlink_to(real)
    download, get = tool(base)

    with pytest.raises(ValueError) as error:
        await download("report.pdf")

    assert str(error.value) == _already_exists("report.pdf")
    assert real.read_bytes() == STAGED
    get.assert_not_awaited()


# ── Hidden names are refused before Graph is asked ────────────────────────────


@pytest.mark.parametrize(
    ("save_path", "component"),
    [(".zshrc", ".zshrc"), (".dir/x.txt", ".dir"), ("sub/.hidden", ".hidden")],
)
async def test_a_hidden_name_is_refused_before_graph_is_asked(tool, base, save_path, component):
    (base / "sub").mkdir()
    (base / ".dir").mkdir()
    download, get = tool(base)

    with pytest.raises(ValueError) as error:
        await download(save_path)

    assert str(error.value) == _hidden(component)
    assert not (base / save_path).exists()
    get.assert_not_awaited()


@no_symlinks_on_windows
async def test_a_visible_name_that_resolves_to_a_hidden_one_is_refused(tool, base):
    """What gets written is the resolved path, so that is the one that must not be hidden."""
    (base / "profile").symlink_to(base / ".zshrc")
    download, get = tool(base)

    with pytest.raises(ValueError) as error:
        await download("profile")

    assert str(error.value) == _hidden(".zshrc")
    assert not (base / ".zshrc").exists()
    get.assert_not_awaited()


async def test_a_root_inside_a_dot_folder_still_works(tool, tmp_path):
    """The default root is ~/.outlook-mcp/attachments: only names below the root count."""
    base = tmp_path / ".outlook-mcp" / "attachments"
    download, _ = tool(base)

    result = await download("report.pdf")

    assert (base / "report.pdf").read_bytes() == CONTENT
    assert result["saved_to"] == str((base / "report.pdf").resolve())


def test_the_landing_check_refuses_a_target_outside_the_root_on_its_own(tmp_path):
    """It judges names below the root, so a target not under it (the root moved after the
    path was confined) is refused rather than judged against the wrong root."""
    base = tmp_path / "attachments"
    base.mkdir()
    elsewhere = tmp_path / "elsewhere"
    elsewhere.mkdir()

    with pytest.raises(ValueError, match="outside the permitted directory"):
        _validated_download_target(str(elsewhere / "report.pdf"), str(base), requested="report.pdf")


def test_a_moved_root_is_refused_in_the_resolvers_words_and_leaves_settings_alone(tmp_path):
    base = tmp_path / "attachments"
    base.mkdir()
    elsewhere = tmp_path / "elsewhere"
    elsewhere.mkdir()

    with pytest.raises(ValueError) as error:
        _validated_download_target(str(elsewhere / "report.pdf"), str(base), requested="report.pdf")

    message = str(error.value)
    assert message.startswith("Attachment path is outside the permitted directory: report.pdf. ")
    assert f"Attachments may only be read from or written to {base}" in message
    assert message.endswith(LEAVE_SETTINGS_TO_THE_USER)


# ── The write itself never replaces a file ────────────────────────────────────


async def test_a_file_created_while_graph_answers_is_not_replaced(tool, base):
    target = base / "report.pdf"
    download, _ = tool(base, on_get=lambda: target.write_bytes(STAGED))

    with pytest.raises(ValueError) as error:
        await download("report.pdf")

    assert str(error.value) == _already_exists("report.pdf")
    assert target.read_bytes() == STAGED
    assert _leftover_temp_files(base) == []


@posix_modes_only
async def test_a_download_lands_readable_by_its_owner_only(tool, base):
    download, _ = tool(base)

    result = await download("report.pdf")

    saved = Path(result["saved_to"])
    assert saved.read_bytes() == CONTENT
    assert oct(os.stat(saved).st_mode)[-3:] == "600"
    assert _leftover_temp_files(base) == []


async def test_the_temp_file_is_not_a_hidden_one(tool, base, monkeypatch):
    """Downloads never create a hidden file below attachments_dir, their temp file included."""
    real_mkstemp = tempfile.mkstemp
    made = []

    def mkstemp(*args, **kwargs):
        fd, path = real_mkstemp(*args, **kwargs)
        made.append(Path(path))
        return fd, path

    monkeypatch.setattr(tempfile, "mkstemp", mkstemp)
    download, _ = tool(base)

    await download("report.pdf")

    assert len(made) == 1
    assert made[0].parent.resolve() == base.resolve()
    assert not made[0].name.startswith("."), made[0].name
    assert _leftover_temp_files(base) == []


# ── Where hard links are unsupported, an exclusive create does the same job ───


def _no_hard_links(monkeypatch, how):
    if how == "refused":

        def link(*args, **kwargs):
            raise OSError(errno.EPERM, "Operation not permitted")

        monkeypatch.setattr(os, "link", link)
    else:
        monkeypatch.delattr(os, "link")


@posix_modes_only
@pytest.mark.parametrize("how", ["refused", "missing"])
async def test_without_hard_links_a_download_still_lands_owner_only(tool, base, monkeypatch, how):
    _no_hard_links(monkeypatch, how)
    download, _ = tool(base)

    result = await download("report.pdf")

    saved = Path(result["saved_to"])
    assert saved.read_bytes() == CONTENT
    assert oct(os.stat(saved).st_mode)[-3:] == "600"
    assert _leftover_temp_files(base) == []


async def test_without_hard_links_a_file_created_meanwhile_is_not_replaced(tool, base, monkeypatch):
    _no_hard_links(monkeypatch, "refused")
    target = base / "report.pdf"
    download, _ = tool(base, on_get=lambda: target.write_bytes(STAGED))

    with pytest.raises(ValueError) as error:
        await download("report.pdf")

    assert str(error.value) == _already_exists("report.pdf")
    assert target.read_bytes() == STAGED
    assert _leftover_temp_files(base) == []


@no_symlinks_on_windows
async def test_without_hard_links_a_symlink_appearing_meanwhile_is_not_written_through(
    tool, base, tmp_path, monkeypatch
):
    """A symlink (not a regular file) planted at the target while Graph answers is refused
    by the exclusive create: nothing reaches what it points at, and no temp file stays."""
    _no_hard_links(monkeypatch, "refused")
    outside = tmp_path / "outside.txt"
    outside.write_bytes(STAGED)
    target = base / "report.pdf"
    download, _ = tool(base, on_get=lambda: target.symlink_to(outside))

    with pytest.raises(ValueError) as error:
        await download("report.pdf")

    assert str(error.value) == _already_exists("report.pdf")
    assert target.is_symlink() and os.readlink(target) == str(outside)
    assert outside.read_bytes() == STAGED
    assert _leftover_temp_files(base) == []


@pytest.mark.parametrize("how", ["refused", "missing"])
def test_without_hard_links_the_fallbacks_error_is_not_chained_to_the_links(base, monkeypatch, how):
    """The exclusive create runs once the link's failure is dealt with, not inside it.

    Run from the handler, its own error was reported as raised "during handling" of
    the link's OSError (or AttributeError), which reads as two failures.
    """
    _no_hard_links(monkeypatch, how)

    def disk_full(save_path, content):
        raise OSError(errno.ENOSPC, "No space left on device")

    monkeypatch.setattr(mail_attachments, "_create_exclusively", disk_full)

    with pytest.raises(OSError, match="No space left") as error:
        write_download(str(base / "report.pdf"), CONTENT, attachments_dir=str(base))

    assert error.value.__context__ is None
    assert _leftover_temp_files(base) == []


def test_without_hard_links_a_failed_write_leaves_no_partial_file(base, monkeypatch):
    """The exclusive create made the file, so a write that fails part-way removes it."""
    _no_hard_links(monkeypatch, "refused")
    real_fdopen = os.fdopen
    opened = []

    class _DiskFull:
        def __init__(self, fd):
            self._file = real_fdopen(fd, "wb")

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            self._file.close()

        def write(self, data):
            self._file.write(data[:4])
            self._file.flush()
            raise OSError(errno.ENOSPC, "No space left on device")

    def fdopen(fd, mode="r", *args, **kwargs):
        opened.append(fd)
        # The first open is the temp file; the second is the exclusive create at the target.
        return _DiskFull(fd) if len(opened) == 2 else real_fdopen(fd, mode, *args, **kwargs)

    monkeypatch.setattr(os, "fdopen", fdopen)
    target = base / "report.pdf"

    with pytest.raises(OSError, match="No space left"):
        write_download(str(target), CONTENT, attachments_dir=str(base))

    assert len(opened) == 2
    assert not os.path.lexists(target)
    assert _leftover_temp_files(base) == []
