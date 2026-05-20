"""HTML body + inline-image helpers for draft email tools.

The Outlook draft tools accept any combination of:

- A plain-text ``body`` (always required as text fallback for clients that
  cannot render HTML).
- An optional ``html_body`` for rich formatting.
- An optional ``inline_images`` list: each entry either a string path or a
  dict ``{"path": "/abs/path/to.png", "cid": "img1", "placeholder": "{IMG1}"}``.

The helper normalises the inline image list, generates stable CIDs when none
are provided, substitutes placeholders inside the HTML, and appends any
remaining images at the end of the document. The resulting HTML references
each image as ``<img src="cid:CID">`` so MUAs render them inline once the
matching attachment has its ``PR_ATTACH_CONTENT_ID`` set.

This module deliberately performs no filesystem reads beyond ``os.path.isfile``
so it stays trivially testable in CI without real Outlook.
"""
from __future__ import annotations

import hashlib
import os
import re
from html import escape as html_escape
from typing import Iterable

from outlook_desktop_mcp.utils.attachment_safety import (
    UnsafeAttachmentPath,
    sanitize_attachment_filename,
)


_CID_RE = re.compile(r"^[A-Za-z0-9._\-]{1,60}$")
_DANGEROUS_PATH_PREFIX = ("\\\\", "//")


class InvalidInlineImage(ValueError):
    """Raised when an inline image entry is malformed or references a bad path."""


def _generate_cid(path: str, index: int) -> str:
    digest = hashlib.sha1(os.path.abspath(path).encode("utf-8")).hexdigest()[:10]
    return f"img{index}_{digest}"


def _normalize_image_entry(entry, index: int) -> tuple[str, str, str | None]:
    """Return ``(cid, path, placeholder)`` for one inline_images entry."""
    if isinstance(entry, str):
        path = entry
        cid = None
        placeholder = None
    elif isinstance(entry, dict):
        path = entry.get("path") or entry.get("file")
        cid = entry.get("cid")
        placeholder = entry.get("placeholder")
    else:
        raise InvalidInlineImage(
            f"inline_images[{index}] must be a string path or dict, got {type(entry).__name__}"
        )

    if not path or not isinstance(path, str):
        raise InvalidInlineImage(f"inline_images[{index}].path is required")
    if any(path.startswith(p) for p in _DANGEROUS_PATH_PREFIX):
        raise InvalidInlineImage(
            f"inline_images[{index}].path: UNC paths are not allowed"
        )
    abs_path = os.path.abspath(os.path.expanduser(path))
    if not os.path.isfile(abs_path):
        raise InvalidInlineImage(
            f"inline_images[{index}].path does not exist: {path!r}"
        )
    if cid is None:
        cid = _generate_cid(abs_path, index)
    if not _CID_RE.match(cid):
        raise InvalidInlineImage(
            f"inline_images[{index}].cid must match [A-Za-z0-9._-]{{1,60}}: {cid!r}"
        )
    return cid, abs_path, placeholder


def plain_to_html(body: str) -> str:
    """Convert a plain-text body to safe HTML (escape + wrap paragraphs)."""
    escaped = html_escape(body or "", quote=False)
    paragraphs = [p for p in escaped.split("\n\n")]
    return "".join(
        "<p>" + p.replace("\n", "<br>") + "</p>"
        for p in paragraphs
    ) or "<p></p>"


def prepare_inline_html(
    body: str,
    html_body: str,
    inline_images: Iterable | None,
) -> tuple[str, list[tuple[str, str]]]:
    """Build the final HTML body and a list of ``(cid, abs_path)`` to attach.

    - When ``html_body`` is provided, it is used as-is (caller is responsible
      for producing safe HTML).
    - When only ``body`` is provided, it is converted to HTML with paragraph
      and line-break preservation.
    - Each inline image entry's placeholder (if any) is substituted with the
      ``<img>`` tag. Images whose CID is not already referenced in the HTML
      are appended at the end inside their own ``<p>`` so they are still
      rendered inline.
    """
    images = list(inline_images or [])
    normalized: list[tuple[str, str, str | None]] = []
    for index, entry in enumerate(images):
        normalized.append(_normalize_image_entry(entry, index))

    html = html_body if html_body else plain_to_html(body)

    for cid, _path, placeholder in normalized:
        img_tag = f'<img src="cid:{cid}" alt="">'
        if placeholder and placeholder in html:
            html = html.replace(placeholder, img_tag)
        elif f"cid:{cid}" not in html:
            html = html + f'<p>{img_tag}</p>'

    return html, [(cid, path) for cid, path, _ in normalized]


def suggest_attachment_basename(path: str) -> str:
    """Return a safe basename for a file path used as an inline-image attachment."""
    try:
        return sanitize_attachment_filename(os.path.basename(path) or "image")
    except UnsafeAttachmentPath:
        return "image"
