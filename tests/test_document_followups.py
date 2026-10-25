"""Optional reader diagnostics stay useful within the parent's bounded error fields."""

import json
import os
import subprocess
import sys

import pytest
from kiota_abstractions.base_request_configuration import RequestConfiguration

from outlook_mcp.errors import OutlookMCPError
from outlook_mcp.tools import mail_attachment_read, mail_attachments
from tests.test_readable_attachments import ARGS, graph, pdf


async def test_many_blank_pages_report_total_and_explicit_sample_omission():
    sdk, _ = graph(pdf("Synthetic", *([None] * 151)), "blank-pages.pdf", "application/pdf")
    result = await mail_attachment_read.read_attachment(sdk, **ARGS)
    warning = result["warnings"][0]
    assert len(warning) < 300
    assert "151 pages" in warning and "143 more omitted" in warning
    assert "[2, 3, 4, 5, 6, 7, 8, 9]" in warning


async def test_parser_refusal_prints_the_download_remedy_once():
    sdk, _ = graph(pdf("Synthetic", password="synthetic"), "locked.pdf", "application/pdf")
    with pytest.raises(OutlookMCPError) as error:
        await mail_attachment_read.read_attachment(sdk, **ARGS)
    assert str(error.value).lower().count("download") == 1


async def test_generic_parser_failure_prints_the_download_remedy_once():
    sdk, _ = graph(b"not a PDF", "broken.pdf", "application/pdf")
    with pytest.raises(OutlookMCPError) as error:
        await mail_attachment_read.read_attachment(sdk, **ARGS)
    assert "Could not extract this document (" in str(error.value)
    assert str(error.value).lower().count("download") == 1


async def test_attachment_metadata_select_uses_the_current_request_configuration():
    sdk, get = graph(b"synthetic text")
    await mail_attachments.fetch_attachment(sdk, **ARGS, max_bytes=1000)
    first, second = get.await_args_list
    configuration = first.kwargs["request_configuration"]
    assert type(configuration) is RequestConfiguration
    assert configuration.query_parameters.select == ["id", "name", "contentType", "size"]
    assert not second.args and not second.kwargs


def test_documents_only_guidance_does_not_describe_an_unavailable_download(tmp_path):
    code = """
import json
from outlook_mcp import server
print(json.dumps({
    "instructions": server.INSTRUCTIONS,
    "tools": sorted(server.mcp._tool_manager._tools),
}))
"""
    environment = {
        key: value for key, value in os.environ.items() if not key.startswith("OUTLOOK_MCP_")
    }
    environment.update(
        {
            "HOME": str(tmp_path),
            "OUTLOOK_MCP_CONFIG_DIR": str(tmp_path / "config"),
            "OUTLOOK_MCP_TOOLSETS": "documents",
        }
    )
    result = subprocess.run(
        [sys.executable, "-c", code],
        env=environment,
        capture_output=True,
        text=True,
        check=True,
        timeout=20,
    )
    registered = json.loads(result.stdout)
    assert len(registered["tools"]) == 3
    assert "outlook_read_attachment" in registered["tools"]
    assert "outlook_download_attachment" not in registered["tools"]
    reader_rule = next(
        line
        for line in registered["instructions"].splitlines()
        if "outlook_read_attachment" in line
    )
    assert "download" not in reader_rule.lower()
    assert (
        "Only PDF" in reader_rule
        and "XLSX/text chunks may start mid-cell or mid-line" in reader_rule
    )
