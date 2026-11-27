"""People for the intent tools: names, lookalikes, the sent table and the person score.

Pure functions, no Graph and no I/O (contracts §2.3, §2.5, §2.6). The rules here
decide whether a name is safe to act on without asking, so each is stated where
it is enforced:

- **Known means written to.** Only letters in Sent Items make someone known
  (decision D4). Meeting requests and responses are dropped, so accepting a
  stranger's invitation does not make them known, and the owner's own addresses
  never count as recipients. Saved contacts raise the rank but never make anyone
  known: an injected email can get an agent to save its sender as a contact.
- **Received mail never adds to a score.** ``person_score`` has no input for it.
- **Lookalikes.** A display name, domain or address built to look like someone the
  owner writes to is flagged; two shared mail providers never flag each other, and
  a name written in two scripts that belong together (a surname in Latin and Han)
  is fine.
"""

from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime, timedelta

from outlook_mcp.intents.envelope import warn
from outlook_mcp.validation import validate_email

# ── Normalization ─────────────────────────────────────────────────────────

# Letters from other scripts that render like Latin ones, written as escapes: in the
# source they would be indistinguishable from the Latin letters they imitate.
_CONFUSABLES = str.maketrans(
    {
        # Cyrillic
        "\u0430": "a",  # cyrillic small letter a
        "\u0435": "e",  # cyrillic small letter ie
        "\u043e": "o",  # cyrillic small letter o
        "\u0440": "p",  # cyrillic small letter er
        "\u0441": "c",  # cyrillic small letter es
        "\u0443": "y",  # cyrillic small letter u
        "\u0445": "x",  # cyrillic small letter ha
        "\u0456": "i",  # cyrillic small letter byelorussian-ukrainian i
        "\u0458": "j",  # cyrillic small letter je
        "\u0455": "s",  # cyrillic small letter dze
        "\u0501": "d",  # cyrillic small letter komi de
        "\u04cf": "l",  # cyrillic small letter palochka
        # Greek
        "\u03b1": "a",  # greek small letter alpha
        "\u03bf": "o",  # greek small letter omicron
        "\u03c1": "p",  # greek small letter rho
        "\u03bd": "v",  # greek small letter nu
        "\u03c4": "t",  # greek small letter tau
        "\u03ba": "k",  # greek small letter kappa
        "\u03b9": "i",  # greek small letter iota
        # Armenian, beyond the contracts' table (review N1)
        "\u0585": "o",  # armenian small letter oh
        "\u0578": "n",  # armenian small letter vo
        "\u057d": "u",  # armenian small letter seh
        "\u0570": "h",  # armenian small letter ho
    }
)
_NOT_NAME_CHAR = re.compile(r"[^a-z0-9]+")


def _visible(text: str) -> str:
    """``text`` without format characters (zero-width joiners and the like), which
    render as nothing but would split a word or an address in two."""
    return "".join(c for c in text if unicodedata.category(c) != "Cf")


def _fold(text: str) -> str:
    """NFKC, casefold, then NFD with combining marks dropped; format characters go first."""
    folded = unicodedata.normalize("NFKC", _visible(text)).casefold()
    decomposed = unicodedata.normalize("NFD", folded)
    return "".join(c for c in decomposed if not unicodedata.combining(c))


def norm_name(text: str) -> str:
    """A name reduced to comparable ASCII words (§2.5).

    Folded, confusable letters mapped to Latin, and every run of anything outside
    ``[a-z0-9]`` made one space. A name in a script with no Latin lookalikes (Han,
    Hangul) normalizes to nothing: callers never treat two empty names as equal.
    """
    mapped = _fold(text).translate(_CONFUSABLES)
    return _NOT_NAME_CHAR.sub(" ", mapped).strip()


def _loose(text: str) -> str:
    """Folded and whitespace-collapsed, every script kept: matching for names
    ``norm_name`` would empty."""
    return " ".join(_fold(text).split())


_SKELETON_EXTRA = (("rn", "m"), ("vv", "w"), ("0", "o"), ("1", "l"))


def domain_skeleton(domain: str) -> str:
    """A domain with lookalike letters folded: ``rnodern.com`` reads as ``modern.com``."""
    skeleton = _fold(domain).translate(_CONFUSABLES)
    for lookalike, letter in _SKELETON_EXTRA:
        skeleton = skeleton.replace(lookalike, letter)
    return skeleton


# A country's shared second level, under which registrations sit one label lower.
_SHARED_SECOND_LEVEL = frozenset({"co", "com", "org", "net", "ac", "gov", "edu"})


def registrable(domain: str) -> str:
    """The registered part of a domain: its last two labels, or three under a country's
    shared second level (``example.co.uk``). No public-suffix list, on purpose."""
    labels = domain.lower().strip(".").split(".")
    if len(labels) >= 3 and len(labels[-1]) == 2 and labels[-2] in _SHARED_SECOND_LEVEL:
        return ".".join(labels[-3:])
    return ".".join(labels[-2:])


# Shared mail providers' registrable domains: two of these never flag each other.
PROVIDERS = frozenset(
    {
        "gmail.com",
        "googlemail.com",
        "outlook.com",
        "hotmail.com",
        "live.com",
        "msn.com",
        "yahoo.com",
        "ymail.com",
        "rocketmail.com",
        "icloud.com",
        "me.com",
        "mac.com",
        "aol.com",
        "proton.me",
        "protonmail.com",
        "pm.me",
        "gmx.com",
        "gmx.net",
        "gmx.de",
        "mail.com",
        "email.com",
        "zoho.com",
        "yandex.com",
        "yandex.ru",
        "fastmail.com",
        "hey.com",
        "tutanota.com",
    }
)
# Providers that ignore dots in the local part.
_DOTLESS = frozenset({"gmail.com", "googlemail.com"})


def damerau_levenshtein(a: str, b: str) -> int:
    """Edit distance counting a swap of two neighbours as one edit (optimal string alignment)."""
    rows = [list(range(len(b) + 1))]
    for i in range(1, len(a) + 1):
        row = [i] + [0] * len(b)
        for j in range(1, len(b) + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            row[j] = min(rows[i - 1][j] + 1, row[j - 1] + 1, rows[i - 1][j - 1] + cost)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                row[j] = min(row[j], rows[i - 2][j - 2] + 1)
        rows.append(row)
    return rows[len(a)][len(b)]


_CHECKED_SCRIPTS = ("LATIN", "CYRILLIC", "GREEK")


def mixed_script(name: str) -> bool:
    """Whether a name mixes letters from two or more of Latin, Cyrillic and Greek.

    Han, Hiragana, Katakana and Hangul are never counted, so neither a surname written
    in Latin and in Han nor kanji written with kana is flagged.
    """
    seen = set()
    for char in name:
        if not unicodedata.category(char).startswith("L"):
            continue
        script = unicodedata.name(char, "").split(" ", 1)[0]
        if script in _CHECKED_SCRIPTS:
            seen.add(script)
    return len(seen) >= 2


_EMAIL_LIKE = re.compile(r"[^\s<>()\[\]\"',;:]+@[^\s<>()\[\]\"',;:]+\.[^\s<>()\[\]\"',;:]{2,}")


def contains_other_address(name: str, address: str) -> bool:
    """Whether a display name holds an email-like token other than the real address.

    Read after NFKC, so a fullwidth ``＠`` is an ``@``, and without format characters.
    """
    real = address.casefold()
    shown = unicodedata.normalize("NFKC", _visible(name or ""))
    return any(token.casefold() != real for token in _EMAIL_LIKE.findall(shown))


def is_valid_address(address: str) -> bool:
    try:
        validate_email(address)
    except (ValueError, TypeError):
        return False
    return True


# ── The sent table (Q1) ───────────────────────────────────────────────────


# The subjects Outlook gives a meeting response or cancellation it sends. English
# only: a backstop for P-SI1, not the main signal.
MEETING_SUBJECT_PREFIXES = (
    "accepted:",
    "declined:",
    "tentative:",
    "tentatively accepted:",
    "new time proposed:",
    "canceled:",
    "cancelled:",
)


def is_meeting_type(row: dict) -> bool:
    """Whether a message is a meeting request or response by its type alone: its
    ``@odata.type`` names ``eventMessage``, or ``meetingMessageType`` came back.

    The test for any mail that is counted or listed (review B2). The subject is the
    sender's choice, so a subject prefix never hides received mail; only the owner's
    sent table also drops on it (``is_meeting_message``), where dropping is safe.
    """
    if "eventmessage" in str(row.get("@odata.type") or "").lower():
        return True
    return row.get("meetingMessageType") is not None


def is_meeting_message(row: dict) -> bool:
    """Whether a Sent Items row is a meeting request or response rather than a letter.

    Probe P-SI1 settles how Graph marks these in a personal mailbox's Sent Items.
    Until it runs, any one of three signals drops the row, so the owner's accept
    or decline can never make a stranger known:

    - ``@odata.type`` names ``eventMessage`` or a subtype, ignoring case;
    - ``meetingMessageType`` came back, which only meeting messages carry (the
      sent table asks for it through the eventMessage cast);
    - the subject starts with a standard response prefix ("Accepted:" and so on).
      A letter whose subject happens to start that way is dropped too, which can
      only make someone less known, never more.

    **P-SI1 release gate.** A response with no type annotation, no
    ``meetingMessageType`` and a localized subject ("Zugesagt: …") is still read as a
    letter, so accepting a stranger's invitation would make them known. The probe must
    show that Graph marks responses in a personal mailbox's Sent Items by type or by
    the cast property before a release
    (``test_p_si1_gate_a_localized_untyped_accept_does_not_make_a_stranger_known``).
    """
    if is_meeting_type(row):
        return True
    subject = row.get("subject")
    return isinstance(subject, str) and subject.strip().casefold().startswith(
        MEETING_SUBJECT_PREFIXES
    )


def parse_time(value: object) -> datetime | None:
    """A Graph UTC timestamp as an aware datetime, or ``None`` when unreadable."""
    if not isinstance(value, str) or not value:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return None
    return parsed


def recipients(row: dict, keys: Iterable[str] = ("toRecipients", "ccRecipients", "bccRecipients")):
    """``(address, name)`` for each recipient of a Graph message row, addresses lowercased."""
    for key in keys:
        for entry in row.get(key) or []:
            email = (entry or {}).get("emailAddress") if isinstance(entry, dict) else None
            if not isinstance(email, dict):
                continue
            address = email.get("address")
            if isinstance(address, str) and address:
                name = email.get("name")
                yield address.strip().lower(), name if isinstance(name, str) else None


def sender_of(row: dict) -> tuple[str | None, str | None]:
    """``(address, name)`` of a row's ``from``, the address lowercased."""
    email = (
        (row.get("from") or {}).get("emailAddress") if isinstance(row.get("from"), dict) else None
    )
    if not isinstance(email, dict):
        return None, None
    address = email.get("address")
    name = email.get("name")
    return (
        address.strip().lower() if isinstance(address, str) and address else None,
        name if isinstance(name, str) else None,
    )


@dataclass
class SentEntry:
    """Letters to one address in the scanned Sent Items."""

    address: str
    count: int = 0
    last: datetime | None = None
    first: datetime | None = None
    name: str | None = None  # the recipient name on the latest letter


@dataclass
class SentTable:
    """Who the owner writes to, from the newest letters in Sent Items (Q1)."""

    entries: dict[str, SentEntry] = field(default_factory=dict)
    me: frozenset[str] = frozenset()
    scanned: int = 0
    meeting_messages_dropped: int = 0
    # The Sent Items folder's id, read off the scanned rows: a message there is one the
    # owner sent. None when Sent Items is empty, and then no message counts as theirs.
    sent_folder: str | None = None
    # Conversation id → when the owner's letters in it were sent (catch_up's C_me).
    letters: dict[str, list[datetime]] = field(default_factory=dict)

    def get(self, address: str) -> SentEntry | None:
        return self.entries.get(address.lower())

    def known_names(self) -> list[tuple[str, str]]:
        """``(norm_name, address)`` for every recipient whose name normalizes to something."""
        return [
            (norm_name(e.name), e.address)
            for e in self.entries.values()
            if e.name and norm_name(e.name)
        ]


def build_sent_table(rows: Iterable[dict], *, signed_in: str | None) -> SentTable:
    """The sent table from Q1's rows: letters only, the owner's addresses dropped.

    ``me`` is every ``from`` address in Sent Items plus the signed-in username,
    casefolded: what any rule means by "from me". A row that can't be read is
    skipped rather than failing the table.
    """
    rows = list(rows)
    me = {signed_in.strip().lower()} if signed_in and signed_in.strip() else set()
    for row in rows:
        address, _ = sender_of(row)
        if address:
            me.add(address)

    folders = Counter(
        row.get("parentFolderId") for row in rows if isinstance(row.get("parentFolderId"), str)
    )
    table = SentTable(
        me=frozenset(me), sent_folder=folders.most_common(1)[0][0] if folders else None
    )
    for row in rows:
        table.scanned += 1
        # The owner's letters by conversation, for catch_up: a meeting message told by
        # its type only, since these sort mail rather than rank people (review B2).
        sent_at = parse_time(row.get("sentDateTime"))
        conversation = row.get("conversationId")
        if isinstance(conversation, str) and sent_at is not None and not is_meeting_type(row):
            table.letters.setdefault(conversation, []).append(sent_at)
        if is_meeting_message(row):
            table.meeting_messages_dropped += 1
            continue
        sent = parse_time(row.get("sentDateTime"))
        for address, name in recipients(row):
            if address in me:
                continue
            entry = table.entries.setdefault(address, SentEntry(address))
            entry.count += 1
            if sent is not None:
                if entry.last is None or sent > entry.last:
                    entry.last = sent
                    entry.name = name or entry.name
                if entry.first is None or sent < entry.first:
                    entry.first = sent
            elif entry.name is None:
                entry.name = name
    return table


# ── Scores (§2.3) ─────────────────────────────────────────────────────────


def written_score(count: int) -> int:
    """``min(25, 5 + 5 * floor(log2 c))``, 0 for no letters: 1 gives 5, 2 gives 10, 4 gives 15,
    8 gives 20, 16 or more 25."""
    if count <= 0:
        return 0
    return min(25, 5 + 5 * int(math.floor(math.log2(count))))


def recency_score(last: datetime | None, now: datetime) -> int:
    """10 for a letter in the last 90 days, 5 in the last year, else 0."""
    if last is None:
        return 0
    age = now - last
    if age <= timedelta(days=90):
        return 10
    if age <= timedelta(days=365):
        return 5
    return 0


def person_score(
    *, name: int, count: int, last: datetime | None, saved_contact: bool, now: datetime
) -> int:
    """name + written + recency + contact. Received mail is not an input."""
    return name + written_score(count) + recency_score(last, now) + (10 if saved_contact else 0)


def name_match(words: str, display_name: str | None, address: str) -> tuple[int, str | None]:
    """How ``words`` match a person: ``(points, kind)``, best first, ``(0, None)`` for none.

    exact address 100; full name 40; first name or surname 30; every word a prefix
    of a name word 20; a substring of the name 10; in the address's local part 10.
    Names compare after ``norm_name``, or folded with every script kept when
    ``norm_name`` would leave nothing (a name in Han, say).
    """
    if words.strip().casefold() == address.casefold():
        return 100, "address"
    query, name = norm_name(words), norm_name(display_name or "")
    if not query:
        query, name = _loose(words), _loose(display_name or "")
    if not query:
        return 0, None
    tokens = name.split()
    wanted = query.split()
    if name and query == name:
        return 40, "full"
    if tokens and len(wanted) == 1 and query in (tokens[0], tokens[-1]):
        return 30, "first_or_last"
    if tokens and all(any(t.startswith(w) for t in tokens) for w in wanted):
        return 20, "prefix"
    if name and query in name:
        return 10, "contains"
    local = norm_name(address.partition("@")[0]) or _loose(address.partition("@")[0])
    if local and query.replace(" ", "") in local.replace(" ", ""):
        return 10, "local_part"
    return 0, None


# ── Warnings (§2.5) ───────────────────────────────────────────────────────

# The order warnings are listed in, so the same inputs always read the same.
WARNING_ORDER = (
    "invalid_address",
    "idn_domain",
    "mixed_script_name",
    "name_contains_other_address",
    "lookalike_name",
    "lookalike_domain",
    "lookalike_address",
    "first_time",
)


def _local_for_compare(local: str, domain: str) -> str:
    return local.replace(".", "") if domain in _DOTLESS else local


def person_warnings(
    address: str,
    display_name: str | None,
    table: SentTable,
    *,
    written: bool,
    contacts: list[tuple[str, str]],
) -> list[dict]:
    """The people warnings of §2.5 for one address, in ``WARNING_ORDER``.

    ``written`` is whether the address was found in Sent Items by scan or search:
    only an unknown address is checked for lookalikes of known ones. Known names,
    addresses and domains come from the sent table; ``contacts`` (``(name,
    address)`` pairs the call has loaded) add names for ``lookalike_name``.
    """
    if not is_valid_address(address):
        return [warn("invalid_address")]
    address = address.lower()
    local, _, domain = address.rpartition("@")
    found: dict[str, dict] = {}

    if any(label.startswith("xn--") for label in domain.split(".")):
        found["idn_domain"] = warn("idn_domain")
    if display_name and mixed_script(display_name):
        found["mixed_script_name"] = warn("mixed_script_name")
    if display_name and contains_other_address(display_name, address):
        found["name_contains_other_address"] = warn("name_contains_other_address")

    if not written:
        name = norm_name(display_name or "")
        if name:
            for known_name, known_address in table.known_names() + [
                (norm_name(n), a.lower()) for n, a in contacts if n and norm_name(n)
            ]:
                if known_name == name and known_address != address:
                    found["lookalike_name"] = warn("lookalike_name", like=known_address)
                    break

        base = registrable(domain)
        for known in sorted({registrable(a.rpartition("@")[2]) for a in table.entries}):
            if known == base or (known in PROVIDERS and base in PROVIDERS):
                continue
            if domain_skeleton(base) == domain_skeleton(known) or (
                len(known) >= 6 and damerau_levenshtein(base, known) <= 1
            ):
                found["lookalike_domain"] = warn("lookalike_domain", like=known)
                break

        mine = _local_for_compare(local, base)
        for known_address in sorted(table.entries):
            known_local, _, known_domain = known_address.rpartition("@")
            if known_domain != domain or known_local == local:
                continue
            distance = damerau_levenshtein(mine, _local_for_compare(known_local, base))
            if 0 < distance <= 1:
                found["lookalike_address"] = warn("lookalike_address", like=known_address)
                break

        found["first_time"] = warn("first_time")

    return [found[code] for code in WARNING_ORDER if code in found]
