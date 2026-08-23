from __future__ import annotations

import errno
import io
import shutil
from pathlib import Path

import pytest
from scripts.smoke_portable import (
    PortableSmokeError,
    _build_multipart_upload,
    _extract_csrf_token,
    _is_retryable_temporary_delete_error,
    _log_command,
    _remove_temporary_tree,
    configure_utf8_stdio,
    contains_text_ignoring_line_wraps,
    parse_windows_listener_endpoints,
    parse_windows_listeners,
    sanitized_windows_environment,
)


def _windows_delete_error(winerror: int) -> OSError:
    error = OSError(f"Windows delete error {winerror}")
    error.winerror = winerror  # type: ignore[attr-defined]
    return error


def test_configures_legacy_stdio_for_non_ascii_command_logs() -> None:
    stdout_buffer = io.BytesIO()
    stderr_buffer = io.BytesIO()
    stdout = io.TextIOWrapper(stdout_buffer, encoding="cp1252", errors="strict")
    stderr = io.TextIOWrapper(stderr_buffer, encoding="cp1252", errors="strict")
    try:
        configure_utf8_stdio(stdout=stdout, stderr=stderr)
        _log_command([r"C:\workbooklens-测试\WorkbookLensCLI.exe", "--version"], stream=stdout)
        stdout.flush()

        assert stdout.encoding == "utf-8"
        assert stderr.encoding == "utf-8"
        assert "测试" in stdout_buffer.getvalue().decode("utf-8")
    finally:
        stdout.detach()
        stderr.detach()


def test_command_log_escapes_non_ascii_for_unconfigured_cp1252_stream() -> None:
    buffer = io.BytesIO()
    stream = io.TextIOWrapper(buffer, encoding="cp1252", errors="strict")
    try:
        _log_command([r"C:\workbooklens-测试\WorkbookLensCLI.exe", "--version"], stream=stream)
        stream.flush()

        output = buffer.getvalue().decode("cp1252")
        assert "\\u6d4b\\u8bd5" in output
        assert "WorkbookLensCLI.exe --version" in output
    finally:
        stream.detach()


@pytest.mark.parametrize("winerror", [5, 32, 33, 145])
def test_classifies_retryable_windows_delete_errors(winerror: int) -> None:
    assert _is_retryable_temporary_delete_error(_windows_delete_error(winerror))


@pytest.mark.parametrize("error_number", [errno.EACCES, errno.EBUSY, errno.ENOTEMPTY])
def test_classifies_retryable_delete_errnos(error_number: int) -> None:
    assert _is_retryable_temporary_delete_error(OSError(error_number, "locked"))


def test_rejects_nonretryable_windows_delete_error() -> None:
    assert not _is_retryable_temporary_delete_error(_windows_delete_error(87))


def test_temporary_tree_cleanup_retries_windows_file_lock(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workspace = tmp_path / "workspace"
    workspace.mkdir()
    (workspace / "locked.log").write_text("pending close", encoding="utf-8")
    original_rmtree = shutil.rmtree
    attempts = 0

    def flaky_rmtree(path: Path) -> None:
        nonlocal attempts
        attempts += 1
        if attempts == 1:
            raise _windows_delete_error(32)
        original_rmtree(path)

    monkeypatch.setattr("scripts.smoke_portable.shutil.rmtree", flaky_rmtree)

    _remove_temporary_tree(workspace, timeout=1.0, retry_delay=0.0)

    assert attempts == 2
    assert not workspace.exists()


def test_temporary_tree_cleanup_still_fails_after_timeout(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workspace = tmp_path / "workspace"
    workspace.mkdir()
    moments = iter([0.0, 0.0, 15.0])
    attempts = 0

    def locked_rmtree(_path: Path) -> None:
        nonlocal attempts
        attempts += 1
        raise _windows_delete_error(32)

    monkeypatch.setattr("scripts.smoke_portable.shutil.rmtree", locked_rmtree)
    monkeypatch.setattr("scripts.smoke_portable.time.monotonic", lambda: next(moments))
    monkeypatch.setattr("scripts.smoke_portable.time.sleep", lambda _delay: None)

    with pytest.raises(PortableSmokeError, match="could not be removed within 15s"):
        _remove_temporary_tree(workspace, timeout=15.0, retry_delay=0.25)

    assert attempts == 2


def test_temporary_tree_cleanup_propagates_unrelated_error(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workspace = tmp_path / "workspace"
    workspace.mkdir()
    error = _windows_delete_error(87)
    attempts = 0

    def failing_rmtree(_path: Path) -> None:
        nonlocal attempts
        attempts += 1
        raise error

    monkeypatch.setattr("scripts.smoke_portable.shutil.rmtree", failing_rmtree)

    with pytest.raises(OSError) as raised:
        _remove_temporary_tree(workspace, timeout=15.0, retry_delay=0.0)

    assert raised.value is error
    assert attempts == 1


def test_temporary_tree_cleanup_accepts_already_removed_path(tmp_path: Path) -> None:
    _remove_temporary_tree(tmp_path / "missing", timeout=0.0, retry_delay=0.0)


def test_temporary_tree_cleanup_retries_missing_child_while_root_remains(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workspace = tmp_path / "workspace"
    workspace.mkdir()
    original_rmtree = shutil.rmtree
    attempts = 0

    def transient_missing_child(path: Path) -> None:
        nonlocal attempts
        attempts += 1
        if attempts == 1:
            raise FileNotFoundError(
                errno.ENOENT,
                "child disappeared during cleanup",
                str(path / "vanished.tmp"),
            )
        original_rmtree(path)

    monkeypatch.setattr(
        "scripts.smoke_portable.shutil.rmtree",
        transient_missing_child,
    )

    _remove_temporary_tree(workspace, timeout=1.0, retry_delay=0.0)

    assert attempts == 2
    assert not workspace.exists()


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
    path = "C:\\Temp\\\u4e2d\u6587 \u7a7a\u683c\\\u8def\u5f84 \u6df7\u5408\\demo"
    expected = f"Demo complete {path}"
    output = (
        "Demo complete C:\\Temp\\\u4e2d\u6587 \r\n\u7a7a\u683c\\\u8def\u5f84 \u6df7\u5408\\demo\r\n"
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
