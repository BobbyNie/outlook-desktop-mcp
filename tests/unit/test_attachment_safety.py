"""Unit tests for attachment path safety."""
import os
import tempfile

import pytest

from outlook_desktop_mcp.utils.attachment_safety import (
    UnsafeAttachmentPath,
    default_attachment_dir,
    resolve_attachment_path,
    sanitize_attachment_filename,
    validate_save_directory,
)


def test_default_directory_under_home():
    p = default_attachment_dir()
    assert p.startswith(os.path.expanduser("~"))


def test_empty_save_directory_uses_default():
    assert validate_save_directory("") == default_attachment_dir()


@pytest.mark.parametrize("unc", ["\\\\server\\share", "//host/share/path"])
def test_unc_paths_rejected(unc):
    with pytest.raises(UnsafeAttachmentPath):
        validate_save_directory(unc)


def test_save_outside_home_or_tmp_rejected():
    with pytest.raises(UnsafeAttachmentPath):
        validate_save_directory("/etc")


def test_save_in_tmp_allowed():
    p = validate_save_directory(tempfile.gettempdir())
    assert os.path.isabs(p)


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("hello.pdf", "hello.pdf"),
        ("../etc/passwd", "passwd"),
        ("..\\..\\boot.ini", "boot.ini"),
        ("safe (1).docx", "safe (1).docx"),
        ("evil.exe", "evil.exe.txt"),
        ("script.bat", "script.bat.txt"),
        ("", "attachment"),
        (".hidden", "hidden"),
        ("...weird", "weird"),
        ("a/b/c.txt", "c.txt"),
    ],
)
def test_sanitize_attachment_filename(raw, expected):
    out = sanitize_attachment_filename(raw)
    assert "/" not in out and "\\" not in out
    assert out == expected


def test_resolve_attachment_path_under_directory(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    subdir = tmp_path / "downloads"
    out = resolve_attachment_path(str(subdir), "Report.pdf")
    assert out.endswith("Report.pdf")
    assert os.path.dirname(out) == os.path.realpath(str(subdir))


def test_resolve_attachment_rejects_traversal(tmp_path):
    p = resolve_attachment_path(str(tmp_path), "../../etc/passwd")
    assert os.path.dirname(p) == os.path.realpath(str(tmp_path))
