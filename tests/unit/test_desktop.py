from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

import workbooklens.desktop as desktop


class _FakeShownEvent:
    def __init__(self) -> None:
        self.handlers: list[Any] = []

    def __iadd__(self, handler: Any) -> _FakeShownEvent:
        self.handlers.append(handler)
        return self


class _FakeWindow:
    def __init__(self) -> None:
        self.events = SimpleNamespace(shown=_FakeShownEvent())
        self.destroyed = False

    def destroy(self) -> None:
        self.destroyed = True


class _FakeWebview:
    def __init__(
        self,
        *,
        renderer: str = "edgechromium",
        renderer_after_start: str | None = None,
    ) -> None:
        self.settings = {"ALLOW_DOWNLOADS": False}
        self.window = _FakeWindow()
        self.created: dict[str, Any] = {}
        self.started: dict[str, Any] = {}
        self.platforms = SimpleNamespace(
            winforms=SimpleNamespace(renderer=renderer),
        )
        self.renderer_after_start = renderer_after_start

    def create_window(self, title: str, url: str, **kwargs: Any) -> _FakeWindow:
        self.created = {"title": title, "url": url, **kwargs}
        return self.window

    def start(self, **kwargs: Any) -> None:
        self.started = kwargs
        if self.renderer_after_start is not None:
            self.platforms.winforms.renderer = self.renderer_after_start


def test_load_webview_imports_winforms_backend(monkeypatch: pytest.MonkeyPatch) -> None:
    webview = _FakeWebview()
    imported: list[str] = []

    def fake_import(name: str) -> Any:
        imported.append(name)
        if name == "webview":
            return webview
        if name == "webview.platforms.winforms":
            return webview.platforms.winforms
        raise AssertionError(name)

    monkeypatch.setattr(desktop.importlib, "import_module", fake_import)

    assert desktop._load_webview() is webview
    assert imported == ["webview", "webview.platforms.winforms"]


def test_native_window_uses_edge_webview_and_enables_downloads(tmp_path: Path) -> None:
    webview = _FakeWebview()
    server = SimpleNamespace(url="http://127.0.0.1:8765")

    desktop._run_window(webview, server, tmp_path, smoke=False)

    assert webview.settings["ALLOW_DOWNLOADS"] is True
    assert webview.created["title"] == "WorkbookLens"
    assert webview.created["url"] == server.url
    assert webview.created["min_size"] == (900, 620)
    assert webview.started["gui"] == "edgechromium"
    assert webview.started["private_mode"] is False
    assert Path(webview.started["storage_path"]) == tmp_path / "webview"


def test_native_window_rejects_renderer_fallback_after_start(tmp_path: Path) -> None:
    webview = _FakeWebview(renderer_after_start="mshtml")
    server = SimpleNamespace(url="http://127.0.0.1:8765")

    with pytest.raises(RuntimeError, match="EdgeChromium"):
        desktop._run_window(webview, server, tmp_path, smoke=False)

    assert webview.started["gui"] == "edgechromium"
    assert webview.platforms.winforms.renderer == "mshtml"


def test_desktop_owns_server_for_window_lifetime(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    calls: list[str] = []
    server = SimpleNamespace(stop=lambda: calls.append("stop"))
    webview = _FakeWebview()
    monkeypatch.setattr(desktop, "_application_data_root", lambda: tmp_path)
    monkeypatch.setattr(desktop, "_configure_logging", lambda _root: tmp_path / "desktop.log")
    monkeypatch.setattr(desktop, "_load_webview", lambda: webview)
    monkeypatch.setattr(
        desktop,
        "start_local_ui_server",
        lambda **_kwargs: calls.append("start") or server,
    )
    monkeypatch.setattr(
        desktop,
        "_run_window",
        lambda received, owned, root, *, smoke: calls.append(
            f"window:{received is webview}:{owned is server}:{root == tmp_path}:{smoke}"
        ),
    )

    assert desktop.run_desktop([]) == 0
    assert calls == ["start", "window:True:True:True:False", "stop"]


def test_desktop_rejects_mshtml_before_starting_local_server(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    calls: list[str] = []
    shown: list[tuple[str, Path | None]] = []
    webview = _FakeWebview(renderer="mshtml")
    monkeypatch.setattr(desktop, "_application_data_root", lambda: tmp_path)
    monkeypatch.setattr(desktop, "_configure_logging", lambda _root: tmp_path / "desktop.log")
    monkeypatch.setattr(desktop, "_load_webview", lambda: webview)
    monkeypatch.setattr(
        desktop,
        "start_local_ui_server",
        lambda **_kwargs: calls.append("start"),
    )
    monkeypatch.setattr(
        desktop,
        "_show_native_error",
        lambda diagnostic_id, path: shown.append((diagnostic_id, path)),
    )

    assert desktop.run_desktop([]) == 1
    assert calls == []
    assert len(shown) == 1
    assert shown[0][1] == tmp_path / "desktop.log"


def test_desktop_stops_server_and_reports_post_start_mshtml_fallback(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    calls: list[str] = []
    shown: list[tuple[str, Path | None]] = []
    server = SimpleNamespace(
        url="http://127.0.0.1:8765",
        stop=lambda: calls.append("stop"),
    )
    webview = _FakeWebview(renderer_after_start="mshtml")
    monkeypatch.setattr(desktop, "_application_data_root", lambda: tmp_path)
    monkeypatch.setattr(desktop, "_configure_logging", lambda _root: tmp_path / "desktop.log")
    monkeypatch.setattr(desktop, "_load_webview", lambda: webview)
    monkeypatch.setattr(desktop, "start_local_ui_server", lambda **_kwargs: server)
    monkeypatch.setattr(
        desktop,
        "_show_native_error",
        lambda diagnostic_id, path: shown.append((diagnostic_id, path)),
    )

    assert desktop.run_desktop([]) == 1
    assert calls == ["stop"]
    assert len(shown) == 1
    assert shown[0][1] == tmp_path / "desktop.log"


def test_desktop_rejects_unadvertised_arguments() -> None:
    assert desktop.run_desktop(["--version"]) == 2


def test_desktop_failure_shows_only_diagnostic_reference(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    shown: list[tuple[str, Path | None]] = []
    monkeypatch.setattr(desktop, "_application_data_root", lambda: tmp_path)
    monkeypatch.setattr(desktop, "_configure_logging", lambda _root: tmp_path / "desktop.log")
    monkeypatch.setattr(
        desktop,
        "_load_webview",
        lambda: (_ for _ in ()).throw(RuntimeError("sensitive internal detail")),
    )
    monkeypatch.setattr(
        desktop,
        "_show_native_error",
        lambda diagnostic_id, path: shown.append((diagnostic_id, path)),
    )

    assert desktop.run_desktop([]) == 1
    assert len(shown) == 1
    assert len(shown[0][0]) == 12
    assert shown[0][1] == tmp_path / "desktop.log"
