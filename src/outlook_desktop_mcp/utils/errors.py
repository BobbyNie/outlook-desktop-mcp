"""COM error formatting."""
import logging

_logger = logging.getLogger("outlook_desktop_mcp.errors")


def _com_hresult_and_message(e: Exception) -> tuple[int, str] | None:
    """Return (hresult, message) for pywin32 COM errors, or None."""
    try:
        import pythoncom

        if isinstance(e, pythoncom.com_error):
            hr, msg, *_ = e.args
            return int(hr), str(msg)
    except Exception:
        pass
    try:
        import pywintypes

        if isinstance(e, pywintypes.com_error):
            hr, msg, *_ = e.args
            return int(hr), str(msg)
    except Exception:
        pass
    hresult = getattr(e, "hresult", None)
    if hresult is not None:
        return int(hresult), str(e)
    return None


def format_com_error(e: Exception) -> str:
    parsed = _com_hresult_and_message(e)
    if parsed is not None:
        hr, msg = parsed
        _logger.debug("COM error: 0x%08X %s", hr & 0xFFFFFFFF, msg)
        hint = _com_error_hint(hr)
        if hint:
            return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg}. {hint}"
        return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg}"
    _logger.warning("Unexpected non-COM exception: %s: %s", type(e).__name__, e)
    return f"An unexpected error occurred ({type(e).__name__}: {e})."


def _com_error_hint(hresult: int) -> str:
    """Actionable hints for common Outlook automation failures."""
    code = hresult & 0xFFFFFFFF
    if code == 0x80020009:  # DISP_E_EXCEPTION
        return (
            "Outlook rejected the COM call (often DASL Restrict or address-book "
            "access). Allow programmatic access in Trust Center, or retry after "
            "syncing the offline address book."
        )
    if code in (0x80070005, 0x80004005):  # E_ACCESSDENIED, E_FAIL
        return (
            "Access denied — approve Outlook's programmatic access prompt for "
            "address information, or ask IT to allow PromptOOMAddressBookAccess."
        )
    return ""
