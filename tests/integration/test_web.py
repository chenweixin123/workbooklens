from __future__ import annotations

import asyncio
import re
from pathlib import Path

from fastapi import FastAPI
from fastapi.testclient import TestClient
from httpx import Response
from starlette import formparsers
from starlette.types import Message, Scope

from workbooklens import __version__
from workbooklens.demo.workflow import generate_demo_workbook
from workbooklens.web import create_app
from workbooklens.web.app import CSRF_COOKIE_NAME, MAX_MULTIPART_OVERHEAD_BYTES


def _csrf_token(html: str) -> str:
    match = re.search(r'name="csrf_token" value="([A-Za-z0-9_-]+)"', html)
    assert match
    token = match.group(1)
    assert len(token) >= 43
    return token


def _client(app: FastAPI) -> TestClient:
    return TestClient(app, base_url="http://127.0.0.1")


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
        assert "removed on normal server shutdown" in home.text
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
        assert "Validated copy created" in applied.text
        fixed = client.get(f"/sessions/{session_id}/fixed")
        assert fixed.status_code == 200
        assert fixed.content.startswith(b"PK")
        diff = client.get(f"/sessions/{session_id}/diff")
        assert diff.status_code == 200
        assert diff.headers["content-disposition"].startswith("inline")
        assert client.get(f"/sessions/{session_id}/apply-report").status_code == 200


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
    assert "No safe deterministic patches are available" in response.text


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
    assert "could not continue" in response.text
    assert "source workbook was not modified" in response.text


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
            assert response.text == "Invalid Host header"
            _assert_security_headers(response)


def test_web_rejects_missing_invalid_and_cross_origin_csrf() -> None:
    app = create_app()
    with _client(app) as client:
        home = client.get("/")
        token = _csrf_token(home.text)
        for data, headers in (
            ({}, {}),
            ({"csrf_token": "not-the-server-token"}, {}),
            ({"csrf_token": token}, {"origin": "https://example.com"}),
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
        assert "Upload must be .xlsx or .xlsm" in same_origin.text


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

        cross_site = client.post(
            "/scan",
            data={"csrf_token": token},
            headers={"origin": "https://example.com"},
            files={"workbook": ("book.xlsx", b"small", "application/octet-stream")},
        )
        assert cross_site.status_code == 403

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
