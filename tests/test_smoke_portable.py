from __future__ import annotations

from pathlib import Path

from scripts.smoke_portable import (
    _build_multipart_upload,
    _extract_csrf_token,
    contains_text_ignoring_line_wraps,
    parse_windows_listener_endpoints,
    parse_windows_listeners,
    sanitized_windows_environment,
)


def test_sanitized_environment_removes_python_paths() -> None:
    environment = {
        "SystemRoot": r"C:\Windows",
        "PATH": r"C:\Python312;C:\project\.venv\Scripts;C:\Windows\System32",
        "PYTHONPATH": r"C:\project\src",
        "VIRTUAL_ENV": r"C:\project\.venv",
        "PIP_INDEX_URL": "https://example.invalid/simple",
        "TEMP": r"C:\Temp",
    }

    result = sanitized_windows_environment(environment)

    assert "Python" not in result["PATH"]
    assert ".venv" not in result["PATH"]
    assert "PYTHONPATH" not in result
    assert "VIRTUAL_ENV" not in result
    assert "PIP_INDEX_URL" not in result
    assert result["TEMP"] == r"C:\Temp"
    assert result["PATH"].split(";")[0].endswith("System32")
    assert result["SystemRoot"] == r"C:\Windows"
    assert result["WINDIR"] == r"C:\Windows"


def test_matches_non_ascii_path_across_rich_line_wraps() -> None:
    expected = "C:\\Temp\\\u4e2d\u6587 \u7a7a\u683c\\\u8def\u5f84 \u6df7\u5408\\demo"
    output = (
        "Demo complete in C:\\Temp\\\u4e2d\u6587 \r\n"
        "\u7a7a\u683c\\\u8def\u5f84 \u6df7\u5408\\demo\r\n"
    )

    assert contains_text_ignoring_line_wraps(output, expected)


def test_rejects_path_with_missing_non_ascii_characters() -> None:
    path = "C:\\Temp\\\u4e2d\u6587 \u7a7a\u683c\\demo"
    expected = f"Demo complete in {path}"
    output = (
        "Demo complete in C:\\Temp\\\u4e2d? \r\n\u7a7a\u683c\\demo\r\n"
        f"Before: {path}\\before.xlsx\r\n"
    )

    assert not contains_text_ignoring_line_wraps(output, expected)


def test_extracts_csrf_token_from_portable_home_page() -> None:
    token = "a" * 43

    assert (
        _extract_csrf_token(f'<input type="hidden" name="csrf_token" value="{token}">'.encode())
        == token
    )


def test_builds_multipart_workbook_upload(tmp_path: Path) -> None:
    workbook = tmp_path / "demo.xlsx"
    workbook.write_bytes(b"PK\x03\x04workbook")
    token = "b" * 43

    body, content_type = _build_multipart_upload(workbook, token)

    assert content_type.startswith("multipart/form-data; boundary=")
    assert b'name="csrf_token"' in body
    assert token.encode() in body
    assert b'name="workbook"; filename="demo.xlsx"' in body
    assert workbook.read_bytes() in body


def test_parse_windows_listeners_reports_address_and_owner() -> None:
    output = """
  Proto  Local Address          Foreign Address        State           PID
  TCP    127.0.0.1:8765         0.0.0.0:0              LISTENING       1234
  TCP    0.0.0.0:8766           0.0.0.0:0              LISTENING       4321
  TCP    [::1]:8765             [::]:0                 LISTENING       5678
"""

    assert parse_windows_listeners(output, 8765) == [
        ("127.0.0.1", 1234),
        ("::1", 5678),
    ]
    assert parse_windows_listener_endpoints(output) == [
        ("127.0.0.1", 8765, 1234),
        ("0.0.0.0", 8766, 4321),  # noqa: S104 - parser fixture, not a bind.
        ("::1", 8765, 5678),
    ]
