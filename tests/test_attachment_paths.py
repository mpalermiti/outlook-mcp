"""The attachment tools must not reach outside the configured directory.

Three tools take a host filesystem path: ``download_attachment`` writes one,
``send_with_attachments`` and ``attach_to_draft`` read them. Before 1.20.0 the
only check was a substring test for ``..`` on the write path, and none at all on
the read paths — so an absolute path reached any file the server process could
read, and a mail server reads untrusted input for a living. "Attach the file at
<path> and reply" is one injected instruction away from exfiltration.

Confinement is resolved, not textual: ``..`` and a symlink pointing out of the
directory both have to fail, which a string check cannot do.
"""

from pathlib import Path

import pytest

from outlook_mcp.tools.mail_attachments import resolve_attachment_path


@pytest.fixture
def attachments_dir(tmp_path):
    """A configured attachments directory with one legitimate file in it."""
    base = tmp_path / "attachments"
    base.mkdir()
    (base / "report.pdf").write_bytes(b"%PDF-1.4 legitimate")
    return str(base)


@pytest.fixture
def secret(tmp_path):
    """A file outside the attachments directory that must stay unreachable."""
    path = tmp_path / "id_ed25519"
    path.write_text("PRIVATE KEY")
    return path


def test_path_inside_the_directory_is_allowed(attachments_dir):
    resolved = resolve_attachment_path(f"{attachments_dir}/report.pdf", attachments_dir)
    assert resolved == str(Path(attachments_dir).resolve() / "report.pdf")


def test_bare_filename_resolves_inside_the_directory(attachments_dir):
    """An agent that passes just a name gets the configured directory, not the cwd."""
    resolved = resolve_attachment_path("report.pdf", attachments_dir)
    assert resolved == str(Path(attachments_dir).resolve() / "report.pdf")


def test_subdirectory_is_allowed(attachments_dir):
    nested = Path(attachments_dir) / "2026"
    nested.mkdir()
    resolved = resolve_attachment_path(f"{attachments_dir}/2026/x.pdf", attachments_dir)
    assert resolved == str(nested.resolve() / "x.pdf")


def test_absolute_path_outside_is_rejected(attachments_dir, secret):
    """The exfiltration shape: an absolute path to something we were never offered."""
    with pytest.raises(ValueError) as exc:
        resolve_attachment_path(str(secret), attachments_dir)
    assert "attachments_dir" in str(exc.value)


def test_dot_dot_traversal_is_rejected(attachments_dir, secret):
    with pytest.raises(ValueError):
        resolve_attachment_path(f"{attachments_dir}/../id_ed25519", attachments_dir)


def test_symlink_out_of_the_directory_is_rejected(attachments_dir, secret):
    """A string check passes this one; only resolving the real path catches it."""
    link = Path(attachments_dir) / "innocent.pdf"
    link.symlink_to(secret)

    with pytest.raises(ValueError):
        resolve_attachment_path(str(link), attachments_dir)


def test_symlinked_directory_out_is_rejected(attachments_dir, tmp_path):
    outside = tmp_path / "elsewhere"
    outside.mkdir()
    (outside / "x.pdf").write_bytes(b"x")
    (Path(attachments_dir) / "sub").symlink_to(outside)

    with pytest.raises(ValueError):
        resolve_attachment_path(f"{attachments_dir}/sub/x.pdf", attachments_dir)


def test_sibling_directory_with_shared_prefix_is_rejected(tmp_path):
    """`/x/attachments-evil` must not pass a prefix comparison against `/x/attachments`."""
    base = tmp_path / "attachments"
    base.mkdir()
    evil = tmp_path / "attachments-evil"
    evil.mkdir()
    (evil / "x.pdf").write_bytes(b"x")

    with pytest.raises(ValueError):
        resolve_attachment_path(str(evil / "x.pdf"), str(base))


def test_directory_is_created_on_demand(tmp_path):
    """First use must not fail because nobody made the directory."""
    base = tmp_path / "never-created"
    resolved = resolve_attachment_path("out.pdf", str(base))

    assert base.is_dir()
    assert resolved.startswith(str(base.resolve()))
    assert oct(base.stat().st_mode)[-3:] == "700"


def test_user_home_is_expanded(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    resolved = resolve_attachment_path("out.pdf", "~/attach")
    assert resolved == str((tmp_path / "attach").resolve() / "out.pdf")


def test_error_names_the_config_key_and_the_directory(attachments_dir, secret):
    """The message is what the agent reads, so it has to say how to fix it."""
    with pytest.raises(ValueError) as exc:
        resolve_attachment_path(str(secret), attachments_dir)

    message = str(exc.value)
    assert "attachments_dir" in message
    assert attachments_dir in message


@pytest.mark.parametrize("hostile", ["", "   ", "\x00evil"])
def test_empty_or_null_bytes_rejected(attachments_dir, hostile):
    with pytest.raises(ValueError):
        resolve_attachment_path(hostile, attachments_dir)


def test_default_config_confines_to_the_outlook_mcp_directory():
    """The shipped default must be a directory we own, not the whole host."""
    from outlook_mcp.config import Config

    assert Config().attachments_dir == "~/.outlook-mcp/attachments"


@pytest.mark.asyncio
async def test_send_with_attachments_rejects_a_path_outside(attachments_dir, secret):
    """End-to-end: the read path is confined, not just the helper."""
    from outlook_mcp.config import Config
    from outlook_mcp.tools import mail_attachments

    with pytest.raises(ValueError) as exc:
        await mail_attachments.send_with_attachments(
            None,
            to=["someone@example.com"],
            subject="hi",
            body="hi",
            attachment_paths=[str(secret)],
            config=Config(attachments_dir=attachments_dir),
        )
    assert "attachments_dir" in str(exc.value)


@pytest.mark.asyncio
async def test_attach_to_draft_rejects_a_path_outside(attachments_dir, secret):
    from outlook_mcp.config import Config
    from outlook_mcp.tools import mail_attachments

    with pytest.raises(ValueError) as exc:
        await mail_attachments.attach_to_draft(
            None,
            "AAMkFakeDraftId",
            [str(secret)],
            config=Config(attachments_dir=attachments_dir),
        )
    assert "attachments_dir" in str(exc.value)


@pytest.mark.asyncio
async def test_download_attachment_rejects_a_path_outside(attachments_dir, secret):
    from outlook_mcp.config import Config
    from outlook_mcp.tools import mail_attachments

    with pytest.raises(ValueError) as exc:
        await mail_attachments.download_attachment(
            None,
            "AAMkFakeMessageId",
            "AAMkFakeAttachmentId",
            save_path=str(secret),
            config=Config(attachments_dir=attachments_dir),
        )
    assert "attachments_dir" in str(exc.value)
    assert secret.read_text() == "PRIVATE KEY"  # untouched
