from __future__ import annotations

import errno
import socket
import urllib.error
from collections.abc import Iterator
from types import SimpleNamespace
from typing import Any

import pytest
from typer.testing import CliRunner

import workbooklens.cli as cli
import workbooklens.web.launcher as launcher
from workbooklens.exceptions import UsageError


def _occupied_loopback_port() -> Iterator[int]:
    listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
        listener.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
    listener.bind((launcher.LOOPBACK_HOST, 0))
    listener.listen()
    try:
        yield int(listener.getsockname()[1])
    finally:
        listener.close()


def test_port_zero_reserves_an_automatic_loopback_port() -> None:
    binding = launcher.bind_loopback_socket(0)
    try:
        assert binding.port > 0
        assert binding.requested_port == 0
        assert not binding.fell_back
        assert binding.socket.getsockname() == (launcher.LOOPBACK_HOST, binding.port)
    finally:
        binding.socket.close()


def test_occupied_port_is_a_usage_error_without_fallback() -> None:
    for port in _occupied_loopback_port():
        with pytest.raises(UsageError, match=rf"Local port {port} is already in use"):
            launcher.bind_loopback_socket(port)


def test_occupied_port_falls_back_only_when_enabled() -> None:
    for port in _occupied_loopback_port():
        binding = launcher.bind_loopback_socket(port, fallback_port=True)
        try:
            assert binding.fell_back
            assert binding.requested_port == port
            assert binding.port != port
            assert binding.socket.getsockname() == (launcher.LOOPBACK_HOST, binding.port)
        finally:
            binding.socket.close()


def test_non_occupancy_bind_error_does_not_fall_back(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[int] = []

    def fail(port: int) -> socket.socket:
        calls.append(port)
        raise OSError(errno.EACCES, "permission denied")

    monkeypatch.setattr(launcher, "_new_loopback_socket", fail)
    with pytest.raises(UsageError, match="Unable to bind"):
        launcher.bind_loopback_socket(8765, fallback_port=True)
    assert calls == [8765]


def test_run_local_ui_passes_prebound_socket_to_uvicorn(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    app = object()
    captured: dict[str, Any] = {}

    def fake_config(received_app: object, **kwargs: Any) -> object:
        captured["app"] = received_app
        captured["config"] = kwargs
        return object()

    class FakeServer:
        def __init__(self, config: object) -> None:
            captured["server_config"] = config

        def run(self, *, sockets: list[socket.socket]) -> None:
            captured["sockets"] = sockets
            captured["bound"] = sockets[0].getsockname()

    monkeypatch.setattr(launcher, "create_app", lambda **_kwargs: app)
    monkeypatch.setattr(launcher.uvicorn, "Config", fake_config)
    monkeypatch.setattr(launcher.uvicorn, "Server", FakeServer)

    messages: list[str] = []
    launcher.run_local_ui(port=0, status=messages.append)

    config = captured["config"]
    assert captured["app"] is app
    assert config["host"] == launcher.LOOPBACK_HOST
    assert config["port"] == captured["bound"][1]
    assert config["loop"] == "asyncio"
    assert config["http"] == "h11"
    assert config["ws"] == "none"
    assert config["lifespan"] == "on"
    assert len(captured["sockets"]) == 1
    assert captured["sockets"][0].fileno() == -1
    assert messages == [
        f"WorkbookLens local UI: http://127.0.0.1:{config['port']}",
        "Press Ctrl+C in this window to stop WorkbookLens.",
    ]


def test_run_local_ui_closes_socket_when_app_creation_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    binding = launcher.bind_loopback_socket(0)
    monkeypatch.setattr(launcher, "bind_loopback_socket", lambda *_args, **_kwargs: binding)

    def fail_app(**_kwargs: object) -> object:
        raise RuntimeError("app failed")

    monkeypatch.setattr(launcher, "create_app", fail_app)

    with pytest.raises(RuntimeError, match="app failed"):
        launcher.run_local_ui(port=0)
    assert binding.socket.fileno() == -1


class _HealthResponse:
    status = 200

    def __enter__(self) -> _HealthResponse:
        return self

    def __exit__(self, *_args: object) -> None:
        return None

    def read(self, _limit: int) -> bytes:
        return b'{"status":"ok"}'


def test_browser_opens_only_after_health_and_without_system_proxy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    proxy_settings: list[dict[str, str]] = []
    health_calls: list[str] = []
    browser_calls: list[str] = []

    class FakeOpener:
        def open(self, request: Any, *, timeout: float) -> _HealthResponse:
            health_calls.append(request.full_url)
            assert timeout > 0
            if len(health_calls) == 1:
                raise urllib.error.URLError("not ready")
            return _HealthResponse()

    sentinel_handler = object()
    monkeypatch.setattr(
        launcher.urllib.request,
        "ProxyHandler",
        lambda proxies: proxy_settings.append(proxies) or sentinel_handler,
    )
    monkeypatch.setattr(
        launcher.urllib.request,
        "build_opener",
        lambda handler: FakeOpener() if handler is sentinel_handler else None,
    )
    monkeypatch.setattr(launcher.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(
        launcher.webbrowser,
        "open",
        lambda url, **_kwargs: browser_calls.append(url) or True,
    )

    launcher._open_browser_when_ready(
        "http://127.0.0.1:8765",
        status=lambda _message: None,
        timeout=1,
    )

    assert proxy_settings == [{}]
    assert health_calls == [
        "http://127.0.0.1:8765/health",
        "http://127.0.0.1:8765/health",
    ]
    assert browser_calls == ["http://127.0.0.1:8765"]


def test_browser_failure_is_nonfatal(monkeypatch: pytest.MonkeyPatch) -> None:
    opener = SimpleNamespace(open=lambda *_args, **_kwargs: _HealthResponse())
    monkeypatch.setattr(launcher.urllib.request, "ProxyHandler", lambda _proxies: object())
    monkeypatch.setattr(launcher.urllib.request, "build_opener", lambda _handler: opener)

    def fail_browser(*_args: object, **_kwargs: object) -> bool:
        raise RuntimeError("no browser")

    monkeypatch.setattr(launcher.webbrowser, "open", fail_browser)
    messages: list[str] = []

    launcher._open_browser_when_ready(
        "http://127.0.0.1:8765",
        status=messages.append,
        timeout=1,
    )

    assert len(messages) == 1
    assert "no browser" in messages[0]
    assert "http://127.0.0.1:8765" in messages[0]


def test_serve_cli_defaults_and_boolean_flags(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[dict[str, Any]] = []
    monkeypatch.setattr(cli, "run_local_ui", lambda **kwargs: calls.append(kwargs))
    runner = CliRunner()

    default_result = runner.invoke(cli.app, ["serve", "--port", "0"])
    enabled_result = runner.invoke(
        cli.app,
        ["serve", "--port", "0", "--open-browser", "--fallback-port"],
    )

    assert default_result.exit_code == 0, default_result.stdout
    assert enabled_result.exit_code == 0, enabled_result.stdout
    assert calls[0]["open_browser"] is False
    assert calls[0]["fallback_port"] is False
    assert calls[1]["open_browser"] is True
    assert calls[1]["fallback_port"] is True


def test_serve_cli_reports_occupied_port_as_exit_two() -> None:
    runner = CliRunner()
    for port in _occupied_loopback_port():
        result = runner.invoke(cli.app, ["serve", "--port", str(port)])
    assert result.exit_code == 2
    assert f"Local port {port} is already in use" in result.stdout
