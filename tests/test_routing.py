"""Tests for capability routing: tool name -> capability -> account."""

from outlook_mcp.routing import (
    COMPOSED_DIGEST_CAPABILITIES,
    capability_for,
    composed_digest_conflict,
)
from outlook_mcp.toolsets import TOOL_GROUPS

# The mail-centric groups fold into "mail"; calendar/contacts/todo map 1:1;
# the delta tools split by name; account tools follow the active account.
EXPECTED_CAPABILITIES = {
    "mail": {"mail", "drafts", "attachments", "folders", "admin", "digest"},
    "calendar": {"calendar"},
    "contacts": {"contacts"},
    "todo": {"todo"},
}


class TestCapabilityFor:
    def test_mail_group_tools_route_to_mail(self):
        for name in ("outlook_list_inbox", "outlook_create_draft", "outlook_download_attachment"):
            assert capability_for(name) == "mail"

    def test_calendar_contacts_todo(self):
        assert capability_for("outlook_list_events") == "calendar"
        assert capability_for("outlook_search_contacts") == "contacts"
        assert capability_for("outlook_list_tasks") == "todo"

    def test_delta_tools_split_by_name(self):
        assert capability_for("outlook_list_inbox_delta") == "mail"
        assert capability_for("outlook_list_events_delta") == "calendar"
        assert capability_for("outlook_list_contacts_delta") == "contacts"

    def test_identity_tools_follow_the_active_account(self):
        for name in ("outlook_whoami", "outlook_auth_status", "outlook_list_accounts"):
            assert capability_for(name) is None

    def test_unknown_and_missing_names(self):
        assert capability_for("outlook_not_a_real_tool") is None
        assert capability_for(None) is None

    def test_every_classified_tool_has_a_routing_answer(self):
        """Drift guard: a new toolset group needs a routing decision, not a silent None.

        capability_for returning None for a real data tool would quietly route
        it to the active account — the exact cross-account leak the routing
        table exists to prevent. Only the "account" group, the name-split
        "delta" group, and genuinely unknown names may answer None.
        """
        mapped_groups = set().union(*EXPECTED_CAPABILITIES.values())
        for name, group in TOOL_GROUPS.items():
            if group in ("account", "delta"):
                continue
            assert group in mapped_groups, (
                f"toolset group '{group}' (e.g. {name}) has no capability mapping"
            )

    def test_every_delta_tool_is_name_routed(self):
        for name, group in TOOL_GROUPS.items():
            if group == "delta":
                assert capability_for(name) in {"mail", "calendar", "contacts"}, (
                    f"delta tool {name} has no capability override"
                )


class TestComposedDigestConflict:
    """changes_since runs mail/events/contacts deltas against one client.

    In a split setup it would silently return one account's calendar under
    another's mail — the guard turns that into a refusal naming the fix.
    """

    def test_uniform_routing_is_fine(self):
        assert (
            composed_digest_conflict({"mail": "net", "calendar": "net", "contacts": "net"}) is None
        )

    def test_split_routing_refuses(self):
        conflict = composed_digest_conflict({"mail": "net", "calendar": "net", "contacts": "neko"})
        assert conflict is not None
        assert "contacts->neko" in conflict
        assert "outlook_list_contacts_delta" in conflict

    def test_the_guard_covers_exactly_the_composed_capabilities(self):
        assert set(COMPOSED_DIGEST_CAPABILITIES) == {"mail", "calendar", "contacts"}
