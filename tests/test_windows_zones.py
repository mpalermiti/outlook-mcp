"""The generated Windows-to-IANA table (`outlook_mcp._windows_zones`).

Generated from CLDR by `scripts/generate_windows_zones.py`, so these hold the
output to what the calendar conversion relies on: every target is a key this
host's zone database holds, and the names Graph is known to return are there.
"""

from zoneinfo import available_timezones

from outlook_mcp._windows_zones import WINDOWS_TO_IANA


def test_every_iana_target_resolves():
    missing = sorted({zone for zone in WINDOWS_TO_IANA.values()} - available_timezones())
    assert not missing, f"targets this zone database does not hold: {missing}"


def test_the_names_graph_returned_live_are_mapped():
    """`Pacific Standard Time` is what Graph returned for series this server made (2026-11-10)."""
    assert WINDOWS_TO_IANA["Pacific Standard Time"] == "America/Los_Angeles"
    assert WINDOWS_TO_IANA["Eastern Standard Time"] == "America/New_York"
    assert WINDOWS_TO_IANA["GMT Standard Time"] == "Europe/London"


def test_the_table_is_the_whole_cldr_list():
    """CLDR has had over 130 Windows names for years; a short table is a broken generation."""
    assert len(WINDOWS_TO_IANA) > 130
