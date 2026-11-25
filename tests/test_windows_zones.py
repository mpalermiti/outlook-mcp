"""The generated Windows-to-IANA table (`outlook_mcp._windows_zones`).

Generated from CLDR by `scripts/generate_windows_zones.py`, so these hold the
output to what the calendar conversion relies on: every target is a key this
host's zone database holds, and the names Graph is known to return are there.
"""

import importlib.util
import pathlib
from zoneinfo import available_timezones

from outlook_mcp._windows_zones import WINDOWS_TO_IANA

# Loaded by path, as `test_preflight.py` loads its script: `scripts/` is not a package.
_GENERATOR = (
    pathlib.Path(__file__).resolve().parent.parent / "scripts" / "generate_windows_zones.py"
)
_spec = importlib.util.spec_from_file_location("generate_windows_zones", _GENERATOR)
generator = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(generator)


def test_every_iana_target_resolves():
    missing = sorted({zone for zone in WINDOWS_TO_IANA.values()} - available_timezones())
    assert not missing, f"targets this zone database does not hold: {missing}"


def test_the_table_is_the_whole_cldr_list():
    """CLDR has had over 130 Windows names for years; a short table is a broken generation.

    The generator refuses the same floor (`MIN_ROWS`), so the two cannot drift.
    """
    assert len(WINDOWS_TO_IANA) >= generator.MIN_ROWS == 130


# A golden sample of CLDR's 001 rows at the commit the module header names, across
# both hemispheres and the zones users are likeliest to be in. A regeneration that
# changes any of them is a CLDR remap to read before accepting, not noise.
# `Asia/Calcutta` is CLDR's spelling, an alias zoneinfo also holds.
_GOLDEN = {
    "Pacific Standard Time": "America/Los_Angeles",
    "Eastern Standard Time": "America/New_York",
    "GMT Standard Time": "Europe/London",
    "W. Europe Standard Time": "Europe/Berlin",
    "China Standard Time": "Asia/Shanghai",
    "India Standard Time": "Asia/Calcutta",
    "Tokyo Standard Time": "Asia/Tokyo",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "New Zealand Standard Time": "Pacific/Auckland",
    "E. South America Standard Time": "America/Sao_Paulo",
    "South Africa Standard Time": "Africa/Johannesburg",
    "UTC": "Etc/UTC",
}


def test_a_golden_sample_of_rows_is_unchanged():
    assert {name: WINDOWS_TO_IANA.get(name) for name in _GOLDEN} == _GOLDEN
