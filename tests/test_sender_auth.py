"""Only Microsoft's own, topmost Authentication-Results may authenticate a sender (#130).

The From line is whatever the sender wrote. Microsoft checks SPF, DKIM and DMARC
when a message arrives and stamps the verdict above every header the sender
wrote, so the topmost ``Authentication-Results`` is Microsoft's and anything
below it may be the sender's own text. Every header here is synthetic.
"""

import pytest

from outlook_mcp.sender_auth import sender_authentication

FROM = "alice@contoso.com"

# The shape Microsoft stamps on inbound mail: no authserv-id, `action=` without a
# ptype, `;` with and without a following space.
MS_PASS = (
    "spf=pass (sender IP is 192.0.2.10) smtp.mailfrom=contoso.com; "
    "dkim=pass (signature was verified) header.d=contoso.com;"
    "dmarc=pass action=none header.from=contoso.com;compauth=pass reason=100"
)
MS_FAIL = (
    "spf=fail (sender IP is 203.0.113.9) smtp.mailfrom=contoso.com; "
    "dkim=none (message not signed) header.d=none;"
    "dmarc=fail action=none header.from=contoso.com;compauth=fail reason=001"
)
# What a sender can write into the message itself, below Microsoft's stamp.
FORGED = "dkim=pass header.d=contoso.com; dmarc=pass header.from=contoso.com"

NO_HEADER = "no Microsoft Authentication-Results header"
UNREADABLE = "unreadable Authentication-Results header"


def auth(*values, sender=FROM, name="Authentication-Results"):
    return sender_authentication([(name, v) for v in values], sender)


# ── Microsoft's verdict, read whole ─────────────────────────────────────


def test_microsoft_pass_authenticates():
    assert auth(MS_PASS) == {
        "authenticated": True,
        "from_domain": "contoso.com",
        "dkim": "pass",
        "dmarc": "pass",
        "compauth": "pass",
        "reason": "dmarc=pass",
    }


def test_microsoft_fail_is_reported_with_its_verdicts():
    assert auth(MS_FAIL) == {
        "authenticated": False,
        "from_domain": "contoso.com",
        "dkim": None,  # no signature by the From domain
        "dmarc": "fail",
        "compauth": "fail",
        "reason": "dmarc=fail",
    }


@pytest.mark.parametrize(
    "value",
    [
        "mx.microsoft.com; dkim=pass header.d=contoso.com",
        "mx.microsoft.com 1; dkim=pass header.d=contoso.com",
        "AB1CD234EF5.mail.protection.outlook.com; dmarc=pass header.from=contoso.com",
    ],
)
def test_microsoft_authserv_ids_are_accepted(value):
    assert auth(value)["authenticated"] is True


@pytest.mark.parametrize(
    "authserv",
    ["mx.example.net", "mx.google.com", "outlook.com.example.net", "protection.outlook.com"],
)
def test_another_services_header_is_not_microsofts(authserv):
    result = auth(f"{authserv}; dkim=pass header.d=contoso.com")
    assert result["authenticated"] is False
    assert result["reason"] == NO_HEADER


@pytest.mark.parametrize(
    "headers",
    [[], [("Received", "from mail.contoso.com")]],
    ids=["no headers", "no Authentication-Results"],
)
def test_no_header_is_not_authenticated(headers):
    result = sender_authentication(headers, FROM)
    assert result == {
        "authenticated": False,
        "from_domain": "contoso.com",
        "dkim": None,
        "dmarc": None,
        "compauth": None,
        "reason": NO_HEADER,
    }


def test_only_the_header_named_authentication_results_counts():
    # ARC carries a copy of a verdict, and anyone can write an X- header.
    headers = [
        ("ARC-Authentication-Results", "i=1; mx.microsoft.com 1; " + FORGED),
        ("X-Authentication-Results", FORGED),
        ("Authentication-Results", MS_FAIL),
    ]
    assert sender_authentication(headers, FROM)["reason"] == "dmarc=fail"


def test_the_header_name_is_matched_without_regard_to_case():
    assert auth(MS_PASS, name="authentication-results")["authenticated"] is True


# ── Only the topmost header ─────────────────────────────────────────────


def test_a_forged_pass_below_microsofts_fail_is_ignored():
    assert auth(MS_FAIL, FORGED)["authenticated"] is False


def test_a_forged_fail_below_microsofts_pass_is_ignored():
    assert auth(MS_PASS, MS_FAIL)["authenticated"] is True


# ── A pass must be aligned with the From domain ─────────────────────────


def test_a_pass_and_the_domain_must_come_from_one_clause():
    # The pass belongs to example.net's signature; contoso.com's failed.
    result = auth("dkim=pass header.d=example.net; dkim=fail header.d=contoso.com")
    assert result["authenticated"] is False
    assert result["dkim"] == "fail"
    assert result["reason"] == "no pass aligned with the From domain"


def test_one_aligned_pass_among_failures_authenticates():
    result = auth("dkim=fail header.d=example.net; dkim=pass header.d=contoso.com")
    assert result["authenticated"] is True
    assert result["dkim"] == "pass"
    assert result["reason"] == "aligned dkim=pass"


def test_a_pass_of_the_from_domain_wins_over_its_own_failure():
    # Two signatures by the From domain, the first broken: one verified is enough.
    result = auth("dkim=fail header.d=contoso.com; dkim=pass header.d=mail.contoso.com")
    assert result["dkim"] == "pass"
    assert result["authenticated"] is True


def test_an_unaligned_signature_is_not_reported_as_the_senders():
    result = auth("dkim=pass header.d=example.net")
    assert result["dkim"] is None
    assert result["authenticated"] is False


@pytest.mark.parametrize(
    "value",
    [
        "dkim=pass header.d=mail.contoso.com",
        "dkim=pass header.i=@contoso.com",
        "dkim=pass header.i=alice@mail.contoso.com",
    ],
)
def test_the_from_domain_or_a_subdomain_may_sign(value):
    assert auth(value)["authenticated"] is True


@pytest.mark.parametrize(
    "signer",
    ["xcontoso.com", "contoso.com.example.net", "com", "ontoso.com"],
)
def test_a_lookalike_signer_is_not_aligned(signer):
    # Each of these is registrable by someone other than contoso.com's owner.
    assert auth(f"dkim=pass header.d={signer}")["authenticated"] is False


@pytest.mark.parametrize(
    "value",
    [
        "dmarc=pass action=none header.from=example.net",
        "dkim=fail header.d=contoso.com; dmarc=pass action=none header.from=example.net;"
        " compauth=pass reason=100",
        # Microsoft states header.from in the dmarc clause only; compauth has none.
        "spf=pass smtp.mailfrom=contoso.com; compauth=pass reason=100",
        # Every header.from must be the From domain, not just one of them.
        "dmarc=pass header.from=contoso.com header.from=example.net",
        "dmarc=none header.from=contoso.com; spf=pass header.from=example.net; compauth=pass",
    ],
)
def test_dmarc_and_compauth_bind_to_the_from_domain(value):
    assert auth(value)["authenticated"] is False


def test_compauth_pass_authenticates_with_microsofts_header_from():
    # Same-tenant and consumer-to-consumer mail often has no DKIM hop.
    result = auth(
        "spf=pass smtp.mailfrom=contoso.com; dkim=none header.d=none;"
        "dmarc=none action=none header.from=contoso.com;compauth=pass reason=100"
    )
    assert result["authenticated"] is True
    assert result["reason"] == "compauth=pass"


@pytest.mark.parametrize(
    "value",
    [
        "dmarc=bestguesspass action=none header.from=contoso.com;compauth=none reason=905",
        "dmarc=none action=none header.from=contoso.com;compauth=softpass reason=201",
    ],
)
def test_weaker_verdicts_are_not_a_pass(value):
    assert auth(value)["authenticated"] is False


# ── Microsoft's DMARC and composite verdicts are final ──────────────────


@pytest.mark.parametrize(
    ("value", "reason"),
    [
        (
            "dkim=pass header.d=contoso.com; dmarc=fail header.from=contoso.com",
            "dmarc=fail",
        ),
        ("dkim=pass header.d=contoso.com; compauth=fail reason=000", "compauth=fail"),
        ("dkim=pass header.d=contoso.com; compauth=softfail reason=601", "compauth=softfail"),
        (
            "dmarc=pass header.from=contoso.com; compauth=fail reason=001",
            "compauth=fail",
        ),
    ],
)
def test_a_failing_dmarc_or_compauth_refuses_any_pass(value, reason):
    result = auth(value)
    assert result["authenticated"] is False
    assert result["reason"] == reason


@pytest.mark.parametrize(
    ("value", "reason"),
    [
        (
            "spf=none smtp.helo=x;dmarc=pass header.from=contoso.com;"
            "dmarc=none header.from=contoso.com",
            "dmarc stated twice",
        ),
        (
            "spf=none smtp.helo=x;compauth=pass reason=100;dkim=pass header.d=contoso.com;"
            "compauth=pass reason=100",
            "compauth stated twice",
        ),
    ],
)
def test_a_method_stated_twice_refuses_the_header(value, reason):
    # Microsoft states each once; a second is text a sender got into the header.
    assert auth(value) == {
        "authenticated": False,
        "from_domain": "contoso.com",
        "dkim": None,
        "dmarc": None,
        "compauth": None,
        "reason": reason,
    }


# A sender controls the HELO name Microsoft copies into `smtp.helo=`; a `;` in it
# would make the rest read like results Microsoft wrote.
HELO_INJECTED = (
    "spf=none (sender IP is 203.0.113.9) smtp.helo=x;dkim=pass "
    "header.d=contoso.com;dmarc=fail action=none "
    "header.from=contoso.com;compauth=fail reason=001"
)


@pytest.mark.parametrize(
    "value",
    [
        HELO_INJECTED,
        # dkim=pass must be the method result, not text inside another clause.
        "spf=none smtp.helo=dkim=pass header.d=contoso.com",
        "spf=none (dkim=pass header.d=contoso.com; x) smtp.mailfrom=example.net",
        'spf=none smtp.helo="x;dkim=pass header.d=contoso.com"',
    ],
)
def test_results_smuggled_into_other_text_do_not_authenticate(value):
    assert auth(value)["authenticated"] is False


# ── RFC 8601 syntax ─────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "value",
    [
        # Comments with `;` and nesting, CFWS around `=`, a quoted value.
        "spf=pass (sender IP is 192.0.2.10; ok (really)) smtp.mailfrom=contoso.com; "
        'dkim = pass (signature was verified) header.d = "contoso.com"',
        # An escaped quote inside a quoted value, an escaped paren inside a comment.
        'dkim=pass (a \\) b) header.d=contoso.com header.s="sel\\"1"',
        "dkim/1=pass header.d=contoso.com",
        # A property with an empty value, at the end of a clause and before another.
        "spf=none smtp.mailfrom=; dkim=pass header.d=contoso.com",
        "spf=none smtp.mailfrom= smtp.helo=mail.contoso.com; dkim=pass header.d=contoso.com",
        # Folded across lines, as a header may be.
        "spf=pass smtp.mailfrom=contoso.com;\r\n dkim=pass header.d=contoso.com",
    ],
)
def test_rfc8601_forms_are_read(value):
    assert auth(value)["authenticated"] is True


def test_case_is_not_significant():
    result = auth("DKIM=Pass header.d=CONTOSO.com", sender="Alice@Contoso.COM")
    assert result["authenticated"] is True
    assert result["from_domain"] == "contoso.com"


@pytest.mark.parametrize(
    "value",
    [
        "dkim=pass header.d=contoso.com (unterminated",
        'dkim=pass header.d="contoso.com',
        # A clause that is neither the authserv-id nor a method result.
        "dkim=pass header.d=contoso.com; this is not a result",
        "spf=pass smtp.mailfrom=contoso.com; dkim x pass header.d=contoso.com",
        "dkim=pass header.d=contoso.com; policy.dkim=pass",
        "dkim=pass header.d=contoso.com stray",
        # An authserv-id takes at most a version after it.
        "mx.microsoft.com 1 extra; dkim=pass header.d=contoso.com",
        "",
    ],
)
def test_an_unreadable_header_is_not_authenticated(value):
    result = auth(value)
    assert result["authenticated"] is False
    assert result["reason"] == UNREADABLE


# ── The From address ────────────────────────────────────────────────────


@pytest.mark.parametrize("sender", ["", "   ", "alice", "alice@"])
def test_no_from_domain_is_not_authenticated(sender):
    assert auth(MS_PASS, sender=sender) == {
        "authenticated": False,
        "from_domain": None,
        "dkim": None,
        "dmarc": None,
        "compauth": None,
        "reason": "no From address",
    }


def test_the_from_domain_is_the_part_after_the_last_at():
    assert auth(MS_PASS, sender='"a@b"@contoso.com')["authenticated"] is True


def test_a_large_hostile_clause_never_slices_the_remaining_tokens():
    from outlook_mcp.sender_auth import _parse_result

    class NoSlices(list):
        def __getitem__(self, index):
            assert not isinstance(index, slice), (
                "copying the remaining tokens makes parsing quadratic"
            )
            return super().__getitem__(index)

    tokens = NoSlices(["dkim", "=", "pass"] + ["header.d", "=", "contoso.com"] * 16_384)
    result = _parse_result(tokens)
    assert result.result == "pass"
    assert result.props["header.d"] == ("contoso.com",) * 16_384


def test_a_lookalike_domain_can_authenticate_only_itself():
    result = auth("dkim=pass header.d=rnicrosoft.com", sender="alice@rnicrosoft.com")
    assert result["authenticated"] is True
    assert result["from_domain"] == "rnicrosoft.com"


@pytest.mark.parametrize(
    "clause",
    [
        'dmarc="fail\n" header.from=contoso.com',
        '"dmarc"=fail header.from=contoso.com',
        'dmarc="fail" header.from=contoso.com',
        'dm"arc"=fail header.from=contoso.com',
        'dmarc=fa"il" header.from=contoso.com',
        'dmarc=fail header."from\n"=contoso.com',
    ],
)
def test_quoted_keywords_cannot_hide_a_failure_beside_aligned_dkim(clause):
    verdict = auth(clause + "; dkim=pass header.d=contoso.com")
    assert verdict["authenticated"] is False
    assert verdict["reason"] == UNREADABLE


@pytest.mark.parametrize("method", ["dkim", "dmarc", "compauth"])
def test_an_unrecognized_result_is_reported_as_unknown(method):
    verdict = auth(f"{method}=future-result header.d=contoso.com header.from=contoso.com")
    assert verdict[method] == "unknown"
    assert verdict["authenticated"] is False
