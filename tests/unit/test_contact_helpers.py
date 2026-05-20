"""Unit tests for contact_helpers."""
import json

import pytest

from outlook_desktop_mcp.utils.contact_helpers import (
    clamp_contact_count,
    normalize_search_query,
    should_cache_contact_result,
    sort_contacts_by_name,
)


def test_clamp_contact_count():
    assert clamp_contact_count(0) == 1
    assert clamp_contact_count(500) == 200
    assert clamp_contact_count(10) == 10


def test_normalize_search_query():
    assert normalize_search_query("  bob  ") == "bob"
    assert normalize_search_query("   ") is None
    assert normalize_search_query("") is None


def test_should_cache_resolve_only_success():
    ok = json.dumps({"resolved": True, "email": "a@b.com"})
    fail = json.dumps({"resolved": False, "name": "x"})
    assert should_cache_contact_result(ok, tool="resolve_recipient")
    assert not should_cache_contact_result(fail, tool="resolve_recipient")


def test_should_cache_empty_list_when_disabled():
    empty = json.dumps([])
    assert should_cache_contact_result(empty, tool="list_contacts")
    assert not should_cache_contact_result(
        empty, tool="mac:list_contacts", allow_empty_list=False
    )


def test_sort_contacts_by_name():
    data = [
        {"full_name": "Zoe"},
        {"full_name": "Alice"},
    ]
    assert sort_contacts_by_name(data)[0]["full_name"] == "Alice"
