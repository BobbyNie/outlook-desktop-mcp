"""In-memory cache for contact/address-book MCP query results."""
from __future__ import annotations

import hashlib
import json
import logging
import time
from typing import Any

logger = logging.getLogger("outlook_desktop_mcp.contact_cache")

# Default: 7 days
CONTACT_CACHE_TTL_SECONDS = 7 * 24 * 3600


class ContactCache:
    """Process-local TTL cache for contact query JSON responses."""

    def __init__(self, ttl_seconds: float = CONTACT_CACHE_TTL_SECONDS):
        self._ttl = ttl_seconds
        self._entries: dict[str, tuple[float, str]] = {}

    def make_key(self, tool_name: str, **params: Any) -> str:
        """Build a stable cache key from tool name and normalized parameters."""
        normalized = json.dumps(params, sort_keys=True, default=str)
        digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
        return f"{tool_name}:{digest}"

    def get(self, key: str) -> str | None:
        """Return cached value if present and not expired, else None."""
        entry = self._entries.get(key)
        if entry is None:
            return None
        expires_at, value = entry
        if time.time() >= expires_at:
            del self._entries[key]
            return None
        logger.info("contact cache hit: %s", key.split(":", 1)[0])
        return value

    def set(self, key: str, value: str) -> None:
        """Store a value with TTL from now."""
        self._entries[key] = (time.time() + self._ttl, value)

    def clear(self) -> None:
        """Remove all entries (mainly for tests)."""
        self._entries.clear()

    def __len__(self) -> int:
        return len(self._entries)
