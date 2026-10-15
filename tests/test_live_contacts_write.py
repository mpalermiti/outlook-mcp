"""Write-tier live guards for contacts: does every field survive a round trip?

Tier 3 of the silent-no-op audit. The mock suite and the wire-level payload
tests prove what we *send*. Only Graph can say what it *kept* — `is_online`
was accepted and dropped on personal accounts with a 201 and no error. Each
test here writes every parameter a tool accepts, reads the object back
through the tool's own read path, and asserts each value persisted.

Contacts have no outward side effect: nothing is emailed, nothing is shared.
Everything created is deleted in a `finally` and named with LIVE_WRITE_NAME.

    OUTLOOK_MCP_LIVE_WRITE=1 uv run pytest -m live_write -v
"""

from __future__ import annotations

from contextlib import asynccontextmanager

import pytest

from outlook_mcp.tools.contacts import create_contact, delete_contact, get_contact, update_contact
from tests.conftest import LIVE_WRITE_NAME

pytestmark = [pytest.mark.live_write, pytest.mark.asyncio]


@asynccontextmanager
async def _temporary_contact(client, config, **kwargs):
    created = await create_contact(client.sdk_client, config=config, **kwargs)
    contact_id = created["id"]
    try:
        yield contact_id
    finally:
        await delete_contact(client.sdk_client, contact_id, config=config)


class TestCreateContactRoundTrip:
    async def test_every_create_field_persists(self, real_graph_client, live_write_config):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="Roundtrip",
            email="live.write.guard@example.com",
            phone="+15555550199",
            company="LiveWriteGuard Co",
            title="Dr.",
            job_title="Probe",
        ) as contact_id:
            got = await get_contact(real_graph_client.sdk_client, contact_id)

            assert got["first_name"] == LIVE_WRITE_NAME
            assert got["last_name"] == "Roundtrip"
            assert "live.write.guard@example.com" in [
                e.get("address") if isinstance(e, dict) else e for e in got["email_addresses"]
            ]
            assert got["mobile_phone"] == "+15555550199"
            assert got["company"] == "LiveWriteGuard Co"
            # #119: a faithful round trip is not enough — the old test wrote a
            # job into `title`, read it back from `title`, and passed while the
            # contact displayed as "Probe <first> <last>". Each value must land
            # in the field that means it.
            assert got["title"] == "Dr."
            assert got["job_title"] == "Probe"
            assert not got["display_name"].startswith("Probe")


class TestUpdateContactRoundTrip:
    async def test_every_update_field_persists(self, real_graph_client, live_write_config):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="Before",
            email="before@example.com",
            phone="+15555550198",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                first_name=f"{LIVE_WRITE_NAME}-After",
                last_name="After",
                email="after@example.com",
                phone="+15555550197",
                job_title="Probe-After",
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["job_title"] == "Probe-After"
            assert got["title"] == ""
            assert got["first_name"] == f"{LIVE_WRITE_NAME}-After"
            assert got["last_name"] == "After"
            assert got["mobile_phone"] == "+15555550197"
            assert "after@example.com" in [
                e.get("address") if isinstance(e, dict) else e for e in got["email_addresses"]
            ]

    async def test_partial_update_leaves_other_fields_alone(
        self, real_graph_client, live_write_config
    ):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="Stable",
            phone="+15555550196",
            company="Keep Me",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                last_name="Changed",
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["last_name"] == "Changed"
            assert got["first_name"] == LIVE_WRITE_NAME
            assert got["mobile_phone"] == "+15555550196"
            assert got["company"] == "Keep Me"


class TestPhoneTypeRoundTrip:
    """#120: each phone type lands in its own field, and only there.

    The issue's live evidence showed Graph accepting ``businessPhones`` on a
    consumer contact; these pin that our own write path gets it there, and that
    the update path can fix a prefix in place (#119's follow-up).
    """

    async def test_create_files_a_business_number_as_business(
        self, real_graph_client, live_write_config
    ):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="BusinessPhone",
            phone="+15555550195",
            phone_type="business",
        ) as contact_id:
            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["business_phones"] == ["+15555550195"]
            assert got["mobile_phone"] == ""
            assert got["home_phones"] == []

    async def test_update_files_a_home_number_and_corrects_the_prefix(
        self, real_graph_client, live_write_config
    ):
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="HomePhone",
            phone="+15555550194",
            title="Probe",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                phone="+15555550193",
                phone_type="home",
                title="Dr.",
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["home_phones"] == ["+15555550193"]
            # Untouched: phone_type picks one field, it does not move the others.
            assert got["mobile_phone"] == "+15555550194"
            assert got["title"] == "Dr."
            assert not got["display_name"].startswith("Probe")

    async def test_update_moves_a_mobile_number_to_business(
        self, real_graph_client, live_write_config
    ):
        """#120's own case on an existing contact: the office number saved as
        mobile moves to business, and no copy stays behind in mobile."""
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="MovePhone",
            phone="+15555550192",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                phone="+15555550192",
                phone_type="business",
                clear_phones=["mobile"],
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["business_phones"] == ["+15555550192"]
            assert got["mobile_phone"] == ""
            assert got["home_phones"] == []

    async def test_update_empties_a_number_list(self, real_graph_client, live_write_config):
        """The lists clear by sending ``[]``, not null — the half #120's
        transcript did not show, so it is pinned here."""
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="ClearPhone",
            phone="+15555550191",
            phone_type="business",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                clear_phones=["business"],
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["business_phones"] == []


class TestAddressRoundTrip:
    """The address half of the same question, and the one that 400'd.

    A partial address is the case this feature exists for — someone knows the
    city and not the street — and it was the case that failed: the omitted parts
    were assigned ``None``, the backing store emitted them onto the *contact*
    under their Python names, and Graph answered ``400 The property
    'country_or_region' does not exist on type 'microsoft.graph.contact'``. The
    full five-part write returned 200 throughout, so nothing short of this tier
    saw it.
    """

    _FULL = {
        "street": "2821 252nd Ave SE",
        "city": "Sammamish",
        "state": "WA",
        "postal_code": "98075",
        "country_or_region": "USA",
    }

    async def test_every_part_of_every_address_persists(self, real_graph_client, live_write_config):
        """All three slots, because all three read back."""
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="Address",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                home_address=self._FULL,
                business_address={"street": "1 Microsoft Way", "city": "Redmond", "state": "WA"},
                other_address={"city": "Bellevue", "country_or_region": "USA"},
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["home_address"] == self._FULL
            assert got["business_address"]["street"] == "1 Microsoft Way"
            assert got["business_address"]["city"] == "Redmond"
            assert got["other_address"]["city"] == "Bellevue"
            assert got["other_address"]["country_or_region"] == "USA"

    async def test_what_get_contact_returns_update_contact_accepts(
        self, real_graph_client, live_write_config
    ):
        """The documented round trip, against Graph rather than against a mock.

        Read an address back and hand it straight to the write path — the thing
        the docstring tells the caller to do to keep the parts it is not
        changing. Any disagreement about a field name between the two halves
        fails here, on the real payload.
        """
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="RoundTripAddress",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                home_address=self._FULL,
                config=live_write_config,
            )
            stored = (await get_contact(real_graph_client.sdk_client, contact_id))["home_address"]

            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                home_address={**stored, "city": "Issaquah"},
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["home_address"] == {**self._FULL, "city": "Issaquah"}

    async def test_a_partial_address_is_accepted(self, real_graph_client, live_write_config):
        """The 400 regression: a city-only patch must reach Graph and be stored."""
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="PartialAddress",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                home_address={"city": "Bothell"},
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["home_address"]["city"] == "Bothell"

    async def test_graph_replaces_the_address_rather_than_merging_it(
        self, real_graph_client, live_write_config
    ):
        """Why the docstring says REPLACES: this is Graph's behaviour, not ours.

        If Graph ever starts merging, this fails and the warning the model reads
        can be softened — which is the only way anyone would find out.
        """
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="ReplacedAddress",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                home_address=self._FULL,
                config=live_write_config,
            )
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                home_address={"city": "Bothell"},
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["home_address"]["city"] == "Bothell"
            assert got["home_address"]["street"] == ""
            assert got["home_address"]["postal_code"] == ""

    async def test_a_name_only_update_leaves_the_address_alone(
        self, real_graph_client, live_write_config
    ):
        """Partial-patch semantics hold at the tool's own level: no address
        argument means the stored addresses are not touched at all."""
        async with _temporary_contact(
            real_graph_client,
            live_write_config,
            first_name=LIVE_WRITE_NAME,
            last_name="KeepsAddress",
        ) as contact_id:
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                home_address={"street": "693 7th St S", "city": "Kirkland"},
                config=live_write_config,
            )
            await update_contact(
                real_graph_client.sdk_client,
                contact_id=contact_id,
                last_name="StillKeepsAddress",
                config=live_write_config,
            )

            got = await get_contact(real_graph_client.sdk_client, contact_id)
            assert got["last_name"] == "StillKeepsAddress"
            assert got["home_address"]["street"] == "693 7th St S"
            assert got["home_address"]["city"] == "Kirkland"
