"""Plain text as HTML, and a reply's text above the quote (``tools/_mail_body.py``).

Graph's ``comment`` joins a multi-line text's lines with spaces on an HTML message, and
``message.body`` drops the quoted original (checked against an Exchange Online mailbox),
so reply and forward text is converted here and patched into the draft Graph quoted.
"""

from __future__ import annotations

from types import SimpleNamespace

from msgraph.generated.models.body_type import BodyType

from outlook_mcp.tools._mail_body import (
    above_quote,
    check_body_format,
    markdown_to_html,
    message_body,
    text_to_html,
)

# The quote a createReply draft carries: Outlook's header block below an <hr>.
_HTML_DRAFT = (
    '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>'
    '<body><hr tabindex="-1" style="display:inline-block; width:98%">'
    '<div id="divRplyFwdMsg" dir="ltr"><b>From:</b> Sender &lt;sender@example.com&gt;</div>'
    "<div><p>The <b>original</b> message.</p></div></body></html>"
)
# The same for a plain-text original: a line break, a rule, the header block.
_TEXT_DRAFT = (
    "\r\n________________________________________\r\n"
    "From: Sender <sender@example.com>\r\nSubject: Plans\r\n\r\nThe original message.\r\n"
)


def _draft(content, content_type=BodyType.Html):
    return SimpleNamespace(content=content, content_type=content_type)


class TestTextToHtml:
    def test_paragraphs_line_breaks_and_lists_survive(self):
        text = (
            "First paragraph line one.\nFirst paragraph line two.\n\nSecond paragraph.\n\n"
            "- bullet one\n- bullet two\n\n**bold** text"
        )

        assert text_to_html(text) == (
            "<p>First paragraph line one.<br>First paragraph line two.</p>"
            "<p>Second paragraph.</p>"
            "<ul><li>bullet one</li><li>bullet two</li></ul>"
            "<p>**bold** text</p>"
        )

    def test_everything_is_escaped(self):
        assert text_to_html('Use <b>x</b> & "y"\n- a <i>') == (
            "<p>Use &lt;b&gt;x&lt;/b&gt; &amp; &quot;y&quot;</p><ul><li>a &lt;i&gt;</li></ul>"
        )

    def test_crlf_and_cr_line_endings(self):
        assert text_to_html("a\r\nb\r\n\r\nc\rd") == "<p>a<br>b</p><p>c<br>d</p>"

    def test_a_line_of_spaces_still_separates_paragraphs(self):
        assert text_to_html("a\n \t\nb\n\n\n\nc") == "<p>a</p><p>b</p><p>c</p>"

    def test_a_list_can_follow_and_precede_text_without_a_blank_line(self):
        assert text_to_html("Items:\n  * one\n  - two\nThat's all.") == (
            "<p>Items:</p><ul><li>one</li><li>two</li></ul><p>That&#x27;s all.</p>"
        )

    def test_only_a_marker_and_a_space_make_an_item(self):
        text = "-- \nA signature\n-1 degrees\n*not a list*\n- \n*"

        assert text_to_html(text) == (
            "<p>--<br>A signature<br>-1 degrees<br>*not a list*<br>-<br>*</p>"
        )

    def test_trailing_spaces_go_and_nothing_is_nothing(self):
        assert text_to_html("a  \n- b  ") == "<p>a</p><ul><li>b</li></ul>"
        assert text_to_html("") == ""
        assert text_to_html("\n \n") == ""


class TestAboveQuote:
    def test_plain_text_goes_into_an_html_draft_as_html_right_after_body(self):
        body = above_quote(_draft(_HTML_DRAFT), "Thanks.\nSee you\n\n- at 10", is_html=False)

        assert body.content_type == BodyType.Html
        assert body.content == _HTML_DRAFT.replace(
            "<body>", "<body><p>Thanks.<br>See you</p><ul><li>at 10</li></ul>"
        )

    def test_html_goes_in_as_written(self):
        body = above_quote(_draft(_HTML_DRAFT), "<p>Thanks, <b>see you</b>.</p>", is_html=True)

        assert body.content_type == BodyType.Html
        assert body.content == _HTML_DRAFT.replace("<body>", "<body><p>Thanks, <b>see you</b>.</p>")

    def test_the_body_tag_is_found_whatever_its_case_and_attributes(self):
        body = above_quote(_draft('<HTML><BODY class="x" lang=en>quoted</BODY></HTML>'), "n", False)

        assert body.content == '<HTML><BODY class="x" lang=en><p>n</p>quoted</BODY></HTML>'

    def test_a_draft_with_no_body_tag_gets_the_text_in_front(self):
        assert above_quote(_draft("<div>quoted</div>"), "n", False).content == (
            "<p>n</p><div>quoted</div>"
        )

    def test_plain_text_stays_plain_text_on_a_plain_text_draft(self):
        body = above_quote(_draft(_TEXT_DRAFT, BodyType.Text), "Thanks.\nSee you\n\n- at 10", False)

        assert body.content_type == BodyType.Text
        # The draft's own CRLF, then Graph's quote unchanged: its opening line
        # break makes the one blank line.
        assert body.content == "Thanks.\r\nSee you\r\n\r\n- at 10\r\n" + _TEXT_DRAFT

    def test_a_plain_text_quote_without_a_leading_break_still_gets_a_blank_line(self):
        body = above_quote(_draft("> quoted\n", BodyType.Text), "Thanks.\r\n\r\n", False)

        assert body.content == "Thanks.\n\n> quoted\n"

    def test_html_on_a_plain_text_draft_makes_it_html_with_the_quote_escaped(self):
        body = above_quote(
            _draft("From: a <a@example.com>\r\nquoted line", BodyType.Text), "<p>n</p>", True
        )

        assert body.content_type == BodyType.Html
        assert body.content == "<p>n</p><br>From: a &lt;a@example.com&gt;<br>quoted line"

    def test_an_empty_draft_body_is_just_the_new_text(self):
        assert above_quote(None, "<p>n</p>", True).content == "<p>n</p>"
        assert above_quote(None, "n", False).content == "<p>n</p>"
        assert above_quote(_draft("", BodyType.Text), "<p>n</p>", True).content == "<p>n</p>"
        text = above_quote(_draft("", BodyType.Text), "a\nb\n", False)
        assert (text.content, text.content_type) == ("a\nb", BodyType.Text)


class TestMarkdownToHtml:
    """`is_markdown`: a safe subset of Markdown, everything else escaped."""

    def test_the_subset(self):
        text = (
            "Hi **Sam**, *quick* note.\nSee `make test` and [the docs](https://example.com/a?b=1&c=2).\n\n"
            "- one\n* two\n+ three\n\n"
            "1. first\n2) second"
        )

        assert markdown_to_html(text) == (
            "<p>Hi <strong>Sam</strong>, <em>quick</em> note.<br>"
            "See <code>make test</code> and "
            '<a href="https://example.com/a?b=1&amp;c=2">the docs</a>.</p>'
            "<ul><li>one</li><li>two</li><li>three</li></ul>"
            "<ol><li>first</li><li>second</li></ol>"
        )

    def test_raw_html_is_escaped_not_passed_through(self):
        assert markdown_to_html('<img src=x onerror="alert(1)"> **<b>x</b>**') == (
            "<p>&lt;img src=x onerror=&quot;alert(1)&quot;&gt; "
            "<strong>&lt;b&gt;x&lt;/b&gt;</strong></p>"
        )

    def test_only_https_and_mailto_links_become_links(self):
        assert markdown_to_html("[a](mailto:sam@example.com) [b](HTTPS://example.com)") == (
            '<p><a href="mailto:sam@example.com">a</a> <a href="HTTPS://example.com">b</a></p>'
        )
        for url in (
            "http://example.com",
            "javascript:alert(1)",
            "data:text/html,x",
            "//example.com",
            "ftp://x",
        ):
            out = markdown_to_html(f"[x]({url})")
            assert "<a" not in out, out
            assert out.startswith("<p>[x](") and out.endswith(")</p>"), out

    def test_an_image_becomes_a_link_without_its_bang(self):
        """`![alt](url)` would otherwise read as "!" and a link."""
        assert markdown_to_html("See ![the chart](https://example.com/c.png).") == (
            '<p>See <a href="https://example.com/c.png">the chart</a>.</p>'
        )
        # Anything that is not a link stays as written, its "!" too.
        assert markdown_to_html("![x](http://example.com/c.png) and ![](https://e.com)") == (
            "<p>![x](http://example.com/c.png) and ![](https://e.com)</p>"
        )
        # An escaped "!" is text, and the link after it still a link.
        assert markdown_to_html("Hi\\![x](https://example.com)") == (
            '<p>Hi!<a href="https://example.com">x</a></p>'
        )

    def test_a_link_address_cannot_break_out_of_its_attribute(self):
        assert markdown_to_html('[x](https://example.com/"onmouseover="a)') == (
            '<p><a href="https://example.com/&quot;onmouseover=&quot;a">x</a></p>'
        )

    def test_code_is_literal(self):
        assert markdown_to_html("`**not bold** <i>`") == (
            "<p><code>**not bold** &lt;i&gt;</code></p>"
        )

    def test_marks_nest_inside_bold_and_links(self):
        assert markdown_to_html("**bold *it* `c`** [**x**](https://example.com)") == (
            "<p><strong>bold <em>it</em> <code>c</code></strong> "
            '<a href="https://example.com"><strong>x</strong></a></p>'
        )

    def test_stars_that_are_not_emphasis_stay_stars(self):
        assert markdown_to_html("2*3*4 and a * b * c and **  ** and a*b") == (
            "<p>2*3*4 and a * b * c and **  ** and a*b</p>"
        )

    def test_underscores_are_left_alone(self):
        """snake_case names are common in an agent's mail; `_x_` is not emphasis here."""
        assert markdown_to_html("set max_retries and _this_") == (
            "<p>set max_retries and _this_</p>"
        )

    def test_backslash_escapes(self):
        assert markdown_to_html(r"\*not italic\* and \`tick\` and \\") == (
            "<p>*not italic* and `tick` and \\</p>"
        )

    def test_everything_else_is_text(self):
        assert markdown_to_html("# Title\n> quoted\n---\n```\ncode\n```") == (
            "<p># Title<br>&gt; quoted<br>---<br>```<br>code<br>```</p>"
        )

    def test_a_numbered_list_keeps_its_first_number(self):
        assert markdown_to_html("3. c\n4. d") == '<ol start="3"><li>c</li><li>d</li></ol>'
        assert markdown_to_html("01. a") == "<ol><li>a</li></ol>"

    def test_lists_and_text_switch_without_a_blank_line(self):
        assert markdown_to_html("Steps:\n1. a\n- b\nDone") == (
            "<p>Steps:</p><ol><li>a</li></ol><ul><li>b</li></ul><p>Done</p>"
        )

    def test_nothing_is_nothing(self):
        assert markdown_to_html("") == ""
        assert markdown_to_html(" \n\n ") == ""


class TestMarkdownInADraft:
    def test_markdown_goes_into_an_html_draft_converted(self):
        body = above_quote(_draft(_HTML_DRAFT), "**Thanks**\n- a", False, is_markdown=True)

        assert body.content_type == BodyType.Html
        assert body.content == _HTML_DRAFT.replace(
            "<body>", "<body><p><strong>Thanks</strong></p><ul><li>a</li></ul>"
        )

    def test_markdown_on_a_plain_text_draft_stays_as_written(self):
        """Markdown reads as plain text, so a plain-text thread stays plain text."""
        body = above_quote(
            _draft(_TEXT_DRAFT, BodyType.Text), "**Thanks**\n- a", False, is_markdown=True
        )

        assert body.content_type == BodyType.Text
        assert body.content == "**Thanks**\r\n- a\r\n" + _TEXT_DRAFT


class TestMessageBody:
    def test_each_format(self):
        text = message_body("a\n**b**", is_html=False, is_markdown=False)
        html = message_body("<p>a</p>", is_html=True, is_markdown=False)
        markdown = message_body("a\n**b**", is_html=False, is_markdown=True)

        assert (text.content, text.content_type) == ("a\n**b**", BodyType.Text)
        assert (html.content, html.content_type) == ("<p>a</p>", BodyType.Html)
        assert (markdown.content, markdown.content_type) == (
            "<p>a<br><strong>b</strong></p>",
            BodyType.Html,
        )

    def test_html_and_markdown_together_are_refused(self):
        import pytest

        with pytest.raises(ValueError, match="is_html or is_markdown, not both"):
            message_body("x", is_html=True, is_markdown=True)
        with pytest.raises(ValueError, match="is_html or is_markdown, not both"):
            check_body_format(True, True)
        check_body_format(True, False)
        check_body_format(False, True)


def test_a_hostile_line_converts_in_linear_time():
    """Mail an agent quotes can hold anything; no line may stall the server."""
    import time

    for line in ("**a " * 20_000, "*a " * 30_000, "[" * 50_000 + "](", "`" + "a" * 100_000):
        started = time.perf_counter()
        markdown_to_html(line)
        assert time.perf_counter() - started < 1.0, line[:20]
