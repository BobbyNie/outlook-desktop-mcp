"""Unit tests for COM error formatting."""
from outlook_desktop_mcp.utils.errors import format_com_error


class _FakeComError(Exception):
    def __init__(self, hresult: int, msg: str):
        super().__init__(hresult, msg, None, None)
        self.hresult = hresult


def test_format_com_error_disp_e_exception_hint():
    err = _FakeComError(-2147352567, "Exception occurred.")
    text = format_com_error(err)
    assert "0x80020009" in text
    assert "programmatic access" in text.lower() or "address book" in text.lower()


def test_format_non_com_shows_type():
    text = format_com_error(ValueError("bad account"))
    assert "ValueError" in text
