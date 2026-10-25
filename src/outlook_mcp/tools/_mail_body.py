"""Mail bodies: plain text as HTML that keeps its layout, and a reply's text above the quote.

Graph's reply and forward actions, and the drafts their ``create*`` forms make, take the
new text in one of two places, and neither keeps a formatted reply intact (checked
against an Exchange Online mailbox):

- ``comment`` is plain text, and Graph joins its lines with spaces: two paragraphs, a
  line break and a list arrive as one run-on paragraph. On a plain-text original it
  also turns the draft into HTML.
- ``message.body`` keeps its markup but replaces the whole body, so the quoted original
  is gone.

So the draft is created with neither. Graph quotes the original into it, as HTML for
an HTML original and as plain text for a plain-text one, and the new text goes in by
one PATCH, above the quote, in the draft's own format.

A body can also be written in Markdown (``is_markdown``): :func:`markdown_to_html`
converts a small, safe subset and escapes the rest.
"""

from __future__ import annotations

import re
from html import escape
from typing import Any

# The opening tag of a draft's HTML body, whatever its case and attributes.
_BODY_OPEN_TAG = re.compile(r"<body\b[^>]*>", re.IGNORECASE)
# A list item: "- " or "* " at the start of a line, after any indentation.
_BULLET = re.compile(r"^\s*[-*]\s+(\S.*)$")
# A blank line, which may hold spaces or tabs.
_BLANK_LINE = re.compile(r"\n[ \t]*\n")

# Markdown lists: "- ", "* " or "+ ", and "1. " or "1) ".
_MD_BULLET = re.compile(r"^\s*[-*+]\s+(\S.*)$")
_MD_NUMBERED = re.compile(r"^\s*(\d{1,9})[.)]\s+(\S.*)$")
# Markdown's inline marks, tried in this order at each position: a backslash
# escape, a code span, a link (an image's leading "!" taken with it, so the image
# becomes a link rather than "!" and a link), bold, italic. No mark's text can hold
# its own closing mark, so a line is scanned in linear time whatever it holds.
# Underscores are not emphasis here: snake_case names are common in mail an agent
# writes.
_MD_INLINE = re.compile(
    r"\\(?P<esc>[\\`*_\[\]()#+.!>-])"
    r"|`(?P<code>[^`\n]+)`"
    r"|!?\[(?P<label>[^\[\]\n]+)\]\((?P<url>[^()\s]+)\)"
    r"|\*\*(?P<bold>[^\s*](?:(?:[^*\n]|\*(?!\*))*[^\s*])?)\*\*"
    r"|(?<![\w*])\*(?P<em>[^\s*](?:[^\n*]*?[^\s*])?)\*(?![\w*])"
)
# The only link targets that become links; anything else stays text.
_MD_SAFE_URL = re.compile(r"(?:https://|mailto:)\S+\Z", re.IGNORECASE)


def text_to_html(text: str) -> str:
    """Plain text as HTML that reads the way the text does.

    A blank line separates paragraphs (``<p>``), a single line break stays one
    (``<br>``), and lines starting ``- `` or ``* `` become a list (``<ul><li>``).
    Everything is HTML-escaped, so ``<b>`` in plain text is shown, not applied.
    Nothing else is interpreted.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    parts: list[str] = []
    for block in _BLANK_LINE.split(text):
        run: list[str] = []
        in_list = False
        for line in block.split("\n"):
            if not line.strip():
                continue
            bullet = _BULLET.match(line)
            if run and (bullet is not None) != in_list:
                parts.append(_render(run, in_list))
                run = []
            in_list = bullet is not None
            run.append(escape(bullet.group(1).rstrip() if bullet else line.rstrip()))
        if run:
            parts.append(_render(run, in_list))
    return "".join(parts)


def _render(lines: list[str], as_list: bool) -> str:
    if as_list:
        return "<ul>" + "".join(f"<li>{line}</li>" for line in lines) + "</ul>"
    return "<p>" + "<br>".join(lines) + "</p>"


def markdown_to_html(text: str) -> str:
    """A safe subset of Markdown as HTML; everything else is escaped text.

    Paragraphs and line breaks as in :func:`text_to_html` (a single line break
    stays one, as in the text), ``-``/``*``/``+`` and numbered lists, ``**bold**``,
    ``*italic*``, ``[text](url)`` links to ``https:`` and ``mailto:`` addresses
    only (an image, ``![alt](url)``, becomes a link to it with ``alt`` as its
    text), ``code`` spans and backslash escapes. No raw HTML passes through:
    ``<b>`` is shown, not applied, and so is a link to any other scheme.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    parts: list[str] = []
    for block in _BLANK_LINE.split(text):
        run: list[str] = []
        kind = start = None
        for line in block.split("\n"):
            if not line.strip():
                continue
            bullet = _MD_BULLET.match(line)
            number = None if bullet else _MD_NUMBERED.match(line)
            this = "ul" if bullet else "ol" if number else "p"
            if run and this != kind:
                parts.append(_render_markdown(run, kind, start))
                run = []
            if not run:
                start = int(number.group(1)) if number else None
            kind = this
            content = bullet.group(1) if bullet else number.group(2) if number else line
            run.append(_markdown_inline(content.strip()))
        if run:
            parts.append(_render_markdown(run, kind, start))
    return "".join(parts)


def _render_markdown(lines: list[str], kind: str | None, start: int | None) -> str:
    if kind == "p":
        return "<p>" + "<br>".join(lines) + "</p>"
    items = "".join(f"<li>{line}</li>" for line in lines)
    if kind == "ol":
        return (f'<ol start="{start}">' if start not in (None, 1) else "<ol>") + items + "</ol>"
    return f"<ul>{items}</ul>"


def _markdown_inline(text: str) -> str:
    out: list[str] = []
    pos = 0
    for mark in _MD_INLINE.finditer(text):
        out.append(escape(text[pos : mark.start()]))
        if mark["esc"] is not None:
            out.append(escape(mark["esc"]))
        elif mark["code"] is not None:
            out.append(f"<code>{escape(mark['code'])}</code>")
        elif mark["url"] is not None:
            if _MD_SAFE_URL.match(mark["url"]):
                label = _markdown_inline(mark["label"])
                out.append(f'<a href="{escape(mark["url"])}">{label}</a>')
            else:
                out.append(escape(mark.group(0)))
        elif mark["bold"] is not None:
            out.append(f"<strong>{_markdown_inline(mark['bold'])}</strong>")
        else:
            out.append(f"<em>{_markdown_inline(mark['em'])}</em>")
        pos = mark.end()
    out.append(escape(text[pos:]))
    return "".join(out)


def check_body_format(is_html: bool, is_markdown: bool) -> None:
    """Refuse a body that claims to be both HTML and Markdown."""
    if is_html and is_markdown:
        raise ValueError("Pass is_html or is_markdown, not both: a body is one or the other.")


def message_body(body: str, is_html: bool, is_markdown: bool) -> Any:
    """The ``ItemBody`` of a new message or a draft update: as written, or Markdown as HTML."""
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.item_body import ItemBody

    check_body_format(is_html, is_markdown)
    result = ItemBody()
    result.content = markdown_to_html(body) if is_markdown else body
    result.content_type = BodyType.Html if is_html or is_markdown else BodyType.Text
    return result


def above_quote(draft_body: Any, body: str, is_html: bool, *, is_markdown: bool = False) -> Any:
    """The body for a reply or forward draft: ``body`` above the original Graph quoted.

    - An HTML draft: ``body`` goes right after its ``<body ...>`` tag, plain text
      converted by :func:`text_to_html` first. That is where Outlook puts a reply, above
      the ``<hr>`` and header block that open the quote, and the original's markup stays
      as Graph wrote it.
    - A plain-text draft, which Graph makes for a plain-text original: plain text stays
      plain text, above the quote, with the draft's line endings. HTML turns the draft
      into HTML, its quote escaped, line breaks kept.

    ``is_markdown`` converts ``body`` with :func:`markdown_to_html` for an HTML draft;
    a plain-text draft takes the Markdown as written, which reads as plain text.
    """
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.item_body import ItemBody

    content = getattr(draft_body, "content", None) or ""
    result = ItemBody()
    result.content_type = BodyType.Html
    if getattr(draft_body, "content_type", None) == BodyType.Text:
        if is_html:
            quoted = escape(content).replace("\r\n", "\n").replace("\n", "<br>")
            result.content = f"{body}<br>{quoted}" if quoted else body
            return result
        newline = "\r\n" if "\r\n" in content else "\n"
        text = body.replace("\r\n", "\n").rstrip("\n").replace("\n", newline)
        # Graph's plain-text quote opens with a line break; one blank line either way.
        gap = newline if content.startswith(("\r\n", "\n")) else newline * 2
        result.content = text + gap + content if content else text
        result.content_type = BodyType.Text
        return result
    if is_html:
        html = body
    elif is_markdown:
        html = markdown_to_html(body)
    else:
        html = text_to_html(body)
    match = _BODY_OPEN_TAG.search(content)
    if match is None:
        result.content = html + content
    else:
        result.content = content[: match.end()] + html + content[match.end() :]
    return result
