from __future__ import annotations

import asyncio
import re
from pathlib import Path

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from httpx import Response
from starlette import formparsers
from starlette.types import Message, Scope

import workbooklens.conversion as conversion
from workbooklens import __version__
from workbooklens.conversion import ConversionProvider, ConversionResult
from workbooklens.demo.workflow import generate_demo_workbook
from workbooklens.exceptions import UsageError
from workbooklens.i18n import require_translation
from workbooklens.web import create_app
from workbooklens.web.app import (
    CSRF_COOKIE_NAME,
    LANGUAGE_COOKIE_NAME,
    MAX_MULTIPART_OVERHEAD_BYTES,
)
from workbooklens.web.templates import TEMPLATES

OLE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")


def _csrf_token(html: str) -> str:
    match = re.search(r'name="csrf_token" value="([A-Za-z0-9_-]+)"', html)
    assert match
    token = match.group(1)
    assert len(token) >= 43
    return token


def _client(app: FastAPI, *, raise_server_exceptions: bool = True) -> TestClient:
    return TestClient(
        app,
        base_url="http://127.0.0.1",
        raise_server_exceptions=raise_server_exceptions,
    )


def _run_asgi_post(
    app: FastAPI,
    headers: list[tuple[bytes, bytes]],
    chunks: list[bytes],
) -> tuple[list[Message], int]:
    async def run_request() -> tuple[list[Message], int]:
        messages: list[Message] = []
        receive_calls = 0

        async def receive() -> Message:
            nonlocal receive_calls
            if receive_calls >= len(chunks):
                return {"type": "http.disconnect"}
            index = receive_calls
            receive_calls += 1
            return {
                "type": "http.request",
                "body": chunks[index],
                "more_body": index < len(chunks) - 1,
            }

        async def send(message: Message) -> None:
            messages.append(message)

        scope: Scope = {
            "type": "http",
            "asgi": {"version": "3.0", "spec_version": "2.3"},
            "http_version": "1.1",
            "method": "POST",
            "scheme": "http",
            "path": "/scan",
            "raw_path": b"/scan",
            "query_string": b"",
            "root_path": "",
            "headers": headers,
            "client": ("127.0.0.1", 54321),
            "server": ("127.0.0.1", 80),
            "app": app,
        }
        await app(scope, receive, send)
        return messages, receive_calls

    return asyncio.run(run_request())


def _response_status(messages: list[Message]) -> int:
    start = next(message for message in messages if message["type"] == "http.response.start")
    return start["status"]


def _response_headers(messages: list[Message]) -> dict[bytes, bytes]:
    start = next(message for message in messages if message["type"] == "http.response.start")
    return dict(start["headers"])


def _assert_security_headers(response: Response) -> None:
    headers = response.headers
    assert headers["cache-control"] == "no-store"
    assert headers["x-content-type-options"] == "nosniff"
    assert headers["x-frame-options"] == "DENY"
    assert headers["referrer-policy"] == "no-referrer"
    assert headers["cross-origin-opener-policy"] == "same-origin"
    assert headers["cross-origin-resource-policy"] == "same-origin"
    policy = headers["content-security-policy"]
    assert "default-src 'none'" in policy
    assert "style-src 'unsafe-inline'" in policy
    assert "script-src 'unsafe-inline'" in policy
    assert "form-action 'self'" in policy
    assert "frame-ancestors 'none'" in policy


def test_local_web_scan_apply_and_download_workflow(tmp_path: Path) -> None:
    workbook = tmp_path / "demo.xlsx"
    generate_demo_workbook(workbook)
    app = create_app(max_file_bytes=5 * 1024 * 1024)
    with _client(app) as client:
        health = client.get("/health")
        assert health.json() == {"status": "ok"}
        _assert_security_headers(health)
        assert client.get("/openapi.json").json()["info"]["version"] == __version__
        home = client.get("/")
        assert home.status_code == 200
        assert "removed during normal shutdown" in home.text
        token = _csrf_token(home.text)
        assert client.cookies.get(CSRF_COOKIE_NAME) == token
        set_cookie = home.headers["set-cookie"].lower()
        assert "httponly" in set_cookie
        assert "samesite=strict" in set_cookie
        with workbook.open("rb") as handle:
            response = client.post(
                "/scan",
                data={"csrf_token": token},
                files={
                    "workbook": (
                        "demo.xlsx",
                        handle,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                },
            )
        assert response.status_code == 200, response.text
        session_match = re.search(r"/sessions/([^/]+)/apply", response.text)
        assert session_match
        session_id = session_match.group(1)
        result_token = _csrf_token(response.text)
        patch_ids = re.findall(r'name="patch_id" value="([^"]+)" data-risk="safe"', response.text)
        assert len(patch_ids) == 4
        assert 'data-risk="layout_review"' in response.text
        report = client.get(f"/sessions/{session_id}/report")
        assert report.status_code == 200
        assert report.headers["content-disposition"].startswith("inline")
        _assert_security_headers(report)
        assert client.get(f"/sessions/{session_id}/plan").status_code == 200
        rejected_apply = client.post(f"/sessions/{session_id}/apply", data={"patch_id": patch_ids})
        assert rejected_apply.status_code == 403
        applied = client.post(
            f"/sessions/{session_id}/apply",
            data={"csrf_token": result_token, "patch_id": patch_ids},
        )
        assert applied.status_code == 200, applied.text
        assert "Repairs completed" in applied.text
        fixed = client.get(f"/sessions/{session_id}/fixed")
        assert fixed.status_code == 200
        assert fixed.content.startswith(b"PK")
        diff = client.get(f"/sessions/{session_id}/diff")
        assert diff.status_code == 200
        assert diff.headers["content-disposition"].startswith("inline")
        assert client.get(f"/sessions/{session_id}/apply-report").status_code == 200


def test_web_language_choice_persists_and_can_be_changed() -> None:
    app = create_app()
    with _client(app) as client:
        chinese = client.get("/?lang=zh-CN")
        assert chinese.status_code == 200
        assert '<html lang="zh-CN">' in chinese.text
        assert "检查并修复 Excel 工作簿" in chinese.text
        assert "本地扫描" in chinese.text
        assert "选择文件" in chinese.text
        assert "未选择文件" in chinese.text
        assert "No file selected" not in chinese.text
        assert "Scan locally" not in chinese.text
        assert client.cookies.get(LANGUAGE_COOKIE_NAME) == "zh-CN"
        language_cookie = next(
            value
            for value in chinese.headers.get_list("set-cookie")
            if value.startswith(f"{LANGUAGE_COOKIE_NAME}=")
        ).lower()
        assert "max-age=31536000" in language_cookie
        assert "httponly" in language_cookie
        assert "samesite=lax" in language_cookie

        persisted = client.get("/")
        assert '<html lang="zh-CN">' in persisted.text
        assert "选择工作簿" in persisted.text

        english = client.get("/?lang=en")
        assert '<html lang="en">' in english.text
        assert "Inspect and repair an Excel workbook" in english.text
        assert "Choose file" in english.text
        assert "No file selected" in english.text
        assert "未选择文件" not in english.text
        assert "检查并修复 Excel 工作簿" not in english.text
        assert client.cookies.get(LANGUAGE_COOKIE_NAME) == "en"


def test_web_file_pickers_use_localized_text_instead_of_native_browser_labels() -> None:
    app = create_app()
    with _client(app) as client:
        response = client.get("/?lang=en")

    assert response.text.count('class="file-picker-control"') == 2
    assert len(re.findall(r'<input[^>]+type="file"', response.text)) == 2
    assert "file-selector-button" not in response.text
    assert "bindFileSelection('workbook'" in response.text
    assert "bindFileSelection('legacy-workbook'" in response.text
    assert "status.textContent = input.files" in response.text


def test_web_uses_configured_or_browser_language_on_first_open() -> None:
    configured = create_app(language="zh-CN")
    with _client(configured) as client:
        response = client.get("/", headers={"accept-language": "en-US,en;q=0.9"})
    assert '<html lang="zh-CN">' in response.text
    assert "选择工作簿" in response.text

    browser_selected = create_app()
    with _client(browser_selected) as client:
        response = client.get("/", headers={"accept-language": "zh-CN,zh;q=0.9,en;q=0.8"})
    assert '<html lang="zh-CN">' in response.text
    assert "选择工作簿" in response.text


def test_web_template_catalog_is_complete() -> None:
    keys = {
        match for template in TEMPLATES.values() for match in re.findall(r"web\.[a-z_]+", template)
    }
    assert keys
    for key in keys:
        assert require_translation(key, "en") != key
        assert require_translation(key, "zh-CN") != key


def test_web_results_can_switch_language_without_resubmitting_upload(tmp_path: Path) -> None:
    workbook = tmp_path / "demo.xlsx"
    generate_demo_workbook(workbook)
    app = create_app(max_file_bytes=5 * 1024 * 1024)
    with _client(app) as client:
        home = client.get("/?lang=en")
        with workbook.open("rb") as handle:
            result = client.post(
                "/scan",
                data={"csrf_token": _csrf_token(home.text), "language": "en"},
                files={
                    "workbook": (
                        "demo.xlsx",
                        handle,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                },
            )
        assert result.status_code == 200
        session_match = re.search(r"/sessions/([^/]+)/apply", result.text)
        assert session_match
        session_id = session_match.group(1)
        assert "Inspection results" in result.text

        chinese = client.get(f"/sessions/{session_id}?lang=zh-CN")
        assert chinese.status_code == 200
        assert '<html lang="zh-CN">' in chinese.text
        assert "检查结果" in chinese.text
        assert "检查建议修复" in chinese.text
        assert "Inspection results" not in chinese.text

        persisted = client.get(f"/sessions/{session_id}")
        assert '<html lang="zh-CN">' in persisted.text
        assert "发现的问题" in persisted.text


def test_web_converts_legacy_xls_with_an_installed_local_provider(
    tmp_path: Path,
    monkeypatch,
) -> None:
    provider = ConversionProvider("excel", "Microsoft Excel", Path("powershell.exe"))
    monkeypatch.setattr(conversion, "available_conversion_providers", lambda: (provider,))

    def fake_convert(
        source: Path,
        output: Path,
        *,
        max_output_bytes: int,
    ) -> ConversionResult:
        assert source.read_bytes().startswith(OLE_SIGNATURE)
        assert max_output_bytes == 5 * 1024 * 1024
        generate_demo_workbook(output)
        return ConversionResult(output, provider)

    monkeypatch.setattr(conversion, "convert_xls_to_xlsx", fake_convert)
    app = create_app(max_file_bytes=5 * 1024 * 1024)
    with _client(app) as client:
        home = client.get("/")
        token = _csrf_token(home.text)
        assert "Available locally: Microsoft Excel" in home.text
        assert "Scanning and repair do not execute formulas" in home.text
        assert "Trusted files only" in home.text
        assert "may recalculate formulas or process workbook-defined behavior" in home.text
        response = client.post(
            "/convert",
            data={"csrf_token": token},
            headers={"origin": "http://localhost"},
            files={
                "legacy_workbook": (
                    "附件5：贵州省2025年研究生科研基金项目立项推荐汇总表_杨双艳.xls",  # noqa: RUF001
                    OLE_SIGNATURE + b"legacy workbook",
                    "application/vnd.ms-excel",
                )
            },
        )
        assert response.status_code == 200, response.text
        assert response.content.startswith(b"PK")
        assert response.headers["content-type"].startswith(conversion.XLSX_MEDIA_TYPE)
        assert response.headers["x-workbooklens-converter"] == "Microsoft Excel"
        assert ".xlsx" in response.headers["content-disposition"]
        assert not list(Path(app.state.root).glob("convert-*"))


def test_web_disables_conversion_button_without_a_local_provider(monkeypatch) -> None:
    monkeypatch.setattr(conversion, "available_conversion_providers", lambda: ())
    app = create_app()

    with _client(app) as client:
        home = client.get("/")

    assert "Converter unavailable" in home.text
    assert "Install Microsoft Excel or LibreOffice" in home.text


@pytest.mark.parametrize(
    ("provider_message", "error_code"),
    [
        ("Upload is not a recognized binary Excel .xls workbook", "WL-CNV-001"),
        ("No local .xls converter is available", "WL-CNV-002"),
        ("Every available local converter failed. provider detail", "WL-CNV-003"),
        ("The local converter did not produce a verified macro-free .xlsx workbook", "WL-CNV-004"),
        ("Every available local converter failed. timed out after 180 seconds", "WL-CNV-005"),
    ],
)
def test_web_conversion_errors_are_stable_localized_and_sanitized(
    monkeypatch,
    provider_message: str,
    error_code: str,
) -> None:
    provider = ConversionProvider("excel", "Microsoft Excel", Path("powershell.exe"))
    monkeypatch.setattr(conversion, "available_conversion_providers", lambda: (provider,))

    def failed_conversion(*_args, **_kwargs):
        raise UsageError(provider_message)

    monkeypatch.setattr(conversion, "convert_xls_to_xlsx", failed_conversion)
    app = create_app(max_file_bytes=1024)
    with _client(app) as client:
        home = client.get("/?lang=zh-CN")
        response = client.post(
            "/convert",
            data={"csrf_token": _csrf_token(home.text), "language": "zh-CN"},
            files={
                "legacy_workbook": (
                    "legacy.xls",
                    OLE_SIGNATURE + b"local test",
                    "application/vnd.ms-excel",
                )
            },
        )

    assert response.status_code == 400
    assert '<html lang="zh-CN">' in response.text
    assert error_code in response.text
    assert "诊断编号" in response.text
    assert provider_message not in response.text
    assert "CLIXML" not in response.text


def test_web_rejects_wrong_or_oversized_legacy_upload(monkeypatch) -> None:
    monkeypatch.setattr(conversion, "available_conversion_providers", lambda: ())
    app = create_app(max_file_bytes=10)
    with _client(app) as client:
        token = _csrf_token(client.get("/").text)
        wrong = client.post(
            "/convert",
            data={"csrf_token": token},
            files={"legacy_workbook": ("book.xlsx", b"PK", "application/octet-stream")},
        )
        assert wrong.status_code == 400
        oversized = client.post(
            "/convert",
            data={"csrf_token": token},
            files={
                "legacy_workbook": (
                    "book.xls",
                    OLE_SIGNATURE + b"x" * 11,
                    "application/vnd.ms-excel",
                )
            },
        )
        assert oversized.status_code == 413


def test_web_rejects_wrong_extension_and_oversized_upload() -> None:
    app = create_app(max_file_bytes=10)
    with _client(app) as client:
        token = _csrf_token(client.get("/").text)
        wrong = client.post(
            "/scan",
            data={"csrf_token": token},
            files={"workbook": ("notes.txt", b"hello", "text/plain")},
        )
        assert wrong.status_code == 400
        oversized = client.post(
            "/scan",
            data={"csrf_token": token},
            files={"workbook": ("book.xlsx", b"x" * 11, "application/octet-stream")},
        )
        assert oversized.status_code == 413


def test_xlsm_web_scan_does_not_offer_repairs(tmp_path: Path) -> None:
    xlsx = tmp_path / "demo.xlsx"
    generate_demo_workbook(xlsx)
    app = create_app(max_file_bytes=5 * 1024 * 1024)
    with _client(app) as client, xlsx.open("rb") as handle:
        token = _csrf_token(client.get("/").text)
        response = client.post(
            "/scan",
            data={"csrf_token": token},
            files={
                "workbook": ("demo.xlsm", handle, "application/vnd.ms-excel.sheet.macroEnabled.12")
            },
        )
    assert response.status_code == 200
    assert 'name="patch_id"' not in response.text
    assert "No reviewable repairs were proposed" in response.text


def test_web_accepts_unicode_filename_with_loopback_alias_and_null_origin(tmp_path: Path) -> None:
    workbook = tmp_path / "source.xlsx"
    generate_demo_workbook(workbook)
    app = create_app(max_file_bytes=5 * 1024 * 1024)
    filename = "附件5：贵州省2025年研究生科研基金项目立项推荐汇总表_杨双艳.xlsx"  # noqa: RUF001
    with _client(app) as client:
        token = _csrf_token(client.get("/").text)
        for origin in ("http://localhost", "null"):
            with workbook.open("rb") as handle:
                response = client.post(
                    "/scan",
                    data={"csrf_token": token},
                    headers={"origin": origin},
                    files={
                        "workbook": (
                            filename,
                            handle,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    },
                )
            assert response.status_code == 200, response.text
            assert filename in response.text


def test_web_errors_are_recoverable_html_pages() -> None:
    app = create_app(max_file_bytes=10)
    with _client(app) as client:
        token = _csrf_token(client.get("/").text)
        response = client.post(
            "/scan",
            data={"csrf_token": token},
            files={"workbook": ("notes.txt", b"hello", "text/plain")},
        )
    assert response.status_code == 400
    assert response.headers["content-type"].startswith("text/html")
    assert "Unsupported file type" in response.text
    assert "source workbook was not modified" in response.text.lower()
    assert "Diagnostic ID" in response.text
    assert "Upload must be .xlsx or .xlsm" not in response.text


def test_web_error_page_never_exposes_raw_provider_output(monkeypatch, caplog) -> None:
    raw_provider_output = (
        '#< CLIXML <S S="Error">C:\\Users\\person\\private.xls at ConvertWorkbook.ps1:42</S>'
    )

    def failed_scan(*_args, **_kwargs):
        raise RuntimeError(raw_provider_output)

    monkeypatch.setattr("workbooklens.web.app.scan_workbook", failed_scan)
    app = create_app(max_file_bytes=1024)
    with _client(app, raise_server_exceptions=False) as client:
        home = client.get("/?lang=en")
        response = client.post(
            "/scan",
            data={"csrf_token": _csrf_token(home.text), "language": "en"},
            files={
                "workbook": (
                    "private.xlsx",
                    b"not parsed by the injected failure",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )

    assert response.status_code == 500
    assert "Diagnostic ID" in response.text
    assert "WL-" in response.text
    assert "CLIXML" not in response.text
    assert "ConvertWorkbook.ps1" not in response.text
    assert "C:\\Users\\person" not in response.text
    assert "RuntimeError" not in response.text
    assert "diagnostic_id=WL-" in caplog.text
    assert "exception_type=RuntimeError" in caplog.text
    assert "CLIXML" not in caplog.text
    assert "C:\\Users\\person" not in caplog.text


def test_web_accepts_only_exact_loopback_host_headers() -> None:
    app = create_app()
    with _client(app) as client:
        for host in ("localhost", "localhost:8123", "127.0.0.1", "127.0.0.1:65535"):
            response = client.get("/health", headers={"host": host})
            assert response.status_code == 200
            _assert_security_headers(response)
        for host in (
            "testserver",
            "localhost.example",
            "127.0.0.2",
            "localhost:0",
            "localhost:65536",
            "localhost:not-a-port",
        ):
            response = client.get("/health", headers={"host": host})
            assert response.status_code == 400
            assert response.text.startswith("WL-REQ-001: ")
            _assert_security_headers(response)


def test_web_rejects_external_and_loopback_lookalike_origins() -> None:
    app = create_app()
    with _client(app) as client:
        token = _csrf_token(client.get("/").text)
        for origin in ("http://example.com", "http://localhost.example"):
            response = client.post(
                "/scan",
                data={"csrf_token": token},
                headers={"origin": origin},
                files={"workbook": ("book.xlsx", b"not parsed", "application/octet-stream")},
            )
            assert response.status_code == 403
            assert response.text.startswith("WL-SEC-001: ")
            _assert_security_headers(response)


def test_early_request_guard_uses_the_selected_language() -> None:
    app = create_app()
    with _client(app) as client:
        home = client.get("/?lang=zh-CN")
        response = client.post(
            "/scan",
            data={"csrf_token": _csrf_token(home.text), "language": "zh-CN"},
            headers={"origin": "https://example.com"},
            files={"workbook": ("book.xlsx", b"not parsed", "application/octet-stream")},
        )

    assert response.status_code == 403
    assert response.text.startswith("WL-SEC-001: ")
    assert "此表单来自不受信任的来源" in response.text
    assert "untrusted origin" not in response.text


def test_web_rejects_missing_invalid_and_cross_origin_csrf() -> None:
    app = create_app()
    with _client(app) as client:
        home = client.get("/")
        token = _csrf_token(home.text)
        for data, headers in (
            ({}, {}),
            ({"csrf_token": "not-the-server-token"}, {}),
            ({"csrf_token": token}, {"origin": "https://example.com"}),
            ({"csrf_token": token}, {"origin": "http://127.0.0.1:81"}),
            ({"csrf_token": token}, {"origin": "https://127.0.0.1"}),
            ({"csrf_token": token}, {"referer": "http://localhost.example/form"}),
        ):
            response = client.post(
                "/scan",
                data=data,
                headers=headers,
                files={"workbook": ("book.xlsx", b"not parsed", "application/octet-stream")},
            )
            assert response.status_code == 403
            _assert_security_headers(response)
        same_origin = client.post(
            "/scan",
            data={"csrf_token": token},
            headers={"origin": "http://127.0.0.1", "referer": "http://127.0.0.1/"},
            files={"workbook": ("notes.txt", b"hello", "text/plain")},
        )
        assert same_origin.status_code == 400
        assert "Upload must be .xlsx or .xlsm" not in same_origin.text
        assert "Diagnostic ID" in same_origin.text

        client.cookies.clear()
        client.cookies.set(CSRF_COOKIE_NAME, "wrong-cookie")
        wrong_cookie = client.post(
            "/scan",
            data={"csrf_token": token},
            headers={"origin": "null"},
            files={"workbook": ("notes.txt", b"hello", "text/plain")},
        )
        assert wrong_cookie.status_code == 403


def test_early_guard_rejects_before_multipart_tempfile(
    monkeypatch,
) -> None:
    app = create_app(max_file_bytes=1024)
    with _client(app) as client:
        home = client.get("/")
        token = _csrf_token(home.text)

        def unexpected_spool(*_args, **_kwargs):
            raise AssertionError("multipart parser created a temporary file before early rejection")

        monkeypatch.setattr(formparsers, "SpooledTemporaryFile", unexpected_spool)

        for origin in ("https://example.com", "http://127.0.0.1:81"):
            rejected = client.post(
                "/scan",
                data={"csrf_token": token},
                headers={"origin": origin},
                files={"workbook": ("book.xlsx", b"small", "application/octet-stream")},
            )
            assert rejected.status_code == 403

        oversized = client.post(
            "/scan",
            data={"csrf_token": token},
            files={
                "workbook": (
                    "book.xlsx",
                    b"x" * (1024 + MAX_MULTIPART_OVERHEAD_BYTES + 1024),
                    "application/octet-stream",
                )
            },
        )
        assert oversized.status_code == 413

        client.cookies.clear()
        missing_cookie = client.post(
            "/scan",
            data={"csrf_token": token},
            files={"workbook": ("book.xlsx", b"small", "application/octet-stream")},
        )
        assert missing_cookie.status_code == 403


def test_declared_oversized_body_is_rejected_without_receive() -> None:
    app = create_app(max_file_bytes=1024)
    token = str(app.state.csrf_token)
    messages, receive_calls = _run_asgi_post(
        app,
        [
            (b"host", b"127.0.0.1"),
            (b"content-type", b"multipart/form-data; boundary=guard"),
            (b"content-length", b"999999"),
            (b"cookie", f"{CSRF_COOKIE_NAME}={token}".encode()),
        ],
        [b"must not be consumed"],
    )

    assert _response_status(messages) == 413
    assert receive_calls == 0

    missing_length_messages, missing_length_calls = _run_asgi_post(
        app,
        [
            (b"host", b"127.0.0.1"),
            (b"content-type", b"multipart/form-data; boundary=guard"),
            (b"cookie", f"{CSRF_COOKIE_NAME}={token}".encode()),
        ],
        [b"must not be consumed"],
    )
    assert _response_status(missing_length_messages) == 411
    assert missing_length_calls == 0


def test_actual_streaming_body_limit_stops_before_full_consumption() -> None:
    app = create_app(max_file_bytes=1)
    token = str(app.state.csrf_token)
    prefix = f"csrf_token={token}&padding=".encode()
    chunks = [prefix + b"x" * 8192, *([b"x" * 8192] * 7)]
    messages, receive_calls = _run_asgi_post(
        app,
        [
            (b"host", b"127.0.0.1"),
            (b"content-type", b"application/x-www-form-urlencoded"),
            (b"cookie", f"{CSRF_COOKIE_NAME}={token}".encode()),
        ],
        chunks,
    )

    assert _response_status(messages) == 413
    headers = _response_headers(messages)
    assert headers[b"cache-control"] == b"no-store"
    assert headers[b"x-content-type-options"] == b"nosniff"
    assert receive_calls == 2
    assert sum(len(chunk) for chunk in chunks[receive_calls:]) == 49_152


def test_security_headers_apply_to_html_json_and_errors() -> None:
    app = create_app()
    with _client(app) as client:
        responses = [
            client.get("/"),
            client.get("/health"),
            client.get("/missing"),
            client.post(
                "/scan",
                files={"workbook": ("book.xlsx", b"not parsed", "application/octet-stream")},
            ),
        ]
    for response in responses:
        _assert_security_headers(response)


def test_serve_command_hardcodes_loopback_binding() -> None:
    cli_path = Path(__file__).parents[2] / "src" / "workbooklens" / "cli.py"
    source = cli_path.read_text(encoding="utf-8")
    assert 'host="127.0.0.1"' in source
