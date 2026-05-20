"""COM error formatting and structured error JSON helpers."""
import json
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")


def format_com_error(e: Exception) -> str:
    try:
        import pythoncom
        if isinstance(e, pythoncom.com_error):
            hr, msg, exc, arg = e.args
            details = exc[2] if exc else "No details"
            _logger.debug("COM error detail: %s", details)  # internal only
            return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg}"
    except Exception:
        pass
    _logger.warning("Unexpected non-COM exception: %s: %s", type(e).__name__, e)
    return "An unexpected error occurred."


def tool_error_json(
    message: str,
    *,
    code: str = "error",
    retriable: bool = False,
    **extra,
) -> str:
    """Return a JSON-serialized error envelope for MCP tool responses.

    Use this for structured failures (timeouts, COM busy, validation errors)
    so LLM clients can branch on ``code`` / ``retriable`` instead of parsing
    free-form strings.
    """
    payload: dict[str, object] = {"error": message, "code": code, "retriable": retriable}
    payload.update(extra)
    return json.dumps(payload, ensure_ascii=False)


def format_bridge_exception(exc: Exception, *, action: str) -> str:
    """Return a structured JSON error for known bridge exceptions.

    Recognises ``ComBridgeTimeoutError``, ``ComBridgeBusyError``, and
    ``ComBridgeDisconnectedError`` by their ``code``/``retriable`` attributes.
    Falls back to ``format_com_error`` for unrecognised exceptions, preserving
    the legacy ``"Error <action>: <message>"`` string contract.

    ``action`` describes what the caller was trying to do
    (e.g. ``"sending email"``).
    """
    code = getattr(exc, "code", None)
    retriable = bool(getattr(exc, "retriable", False))
    if isinstance(code, str):
        extras: dict[str, object] = {}
        if code == "com_timeout":
            extras["warning"] = (
                "Outlook may have completed the operation after the timeout. "
                "Verify in the Outlook UI before retrying — do not assume the "
                "operation failed."
            )
        return tool_error_json(
            f"Error {action}: {exc}",
            code=code,
            retriable=retriable,
            **extras,
        )
    return f"Error {action}: {format_com_error(exc)}"
