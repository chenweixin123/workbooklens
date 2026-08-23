"""Native Windows shell for the loopback-only WorkbookLens interface."""

from __future__ import annotations

import ctypes
import importlib
import locale
import logging
import logging.handlers
import os
import sys
import tempfile
import threading
import time
import uuid
from collections.abc import Sequence
from pathlib import Path
from typing import Any

from workbooklens.web.launcher import DEFAULT_PORT, RunningLocalUI, start_local_ui_server

_SMOKE_ARGUMENT = "--workbooklens-smoke-test"
_WINDOW_TITLE = "WorkbookLens"
_LOG_MAX_BYTES = 2 * 1024 * 1024
_LOG_BACKUPS = 3


def _application_data_root() -> Path:
    raw_root = os.environ.get("LOCALAPPDATA")
    root = Path(raw_root) if raw_root else Path(tempfile.gettempdir())
    return root.expanduser().resolve() / "WorkbookLens"


def _configure_logging(root: Path) -> Path:
    log_directory = root / "logs"
    log_directory.mkdir(parents=True, exist_ok=True)
    log_path = log_directory / "desktop.log"
    handler = logging.handlers.RotatingFileHandler(
        log_path,
        maxBytes=_LOG_MAX_BYTES,
        backupCount=_LOG_BACKUPS,
        encoding="utf-8",
        delay=True,
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
    for logger_name in ("workbooklens", "uvicorn", "uvicorn.error", "pywebview"):
        logger = logging.getLogger(logger_name)
        logger.handlers.clear()
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.propagate = False
    return log_path


def _prefers_chinese() -> bool:
    language = locale.getlocale()[0] or os.environ.get("LANG", "")
    return language.casefold().startswith("zh")


def _show_native_error(diagnostic_id: str, log_path: Path | None) -> None:
    if _prefers_chinese():
        message = (
            "WorkbookLens 无法启动。请确认已安装 Microsoft Edge WebView2 Runtime, "
            "然后重新打开软件。"
        )
        if log_path is not None:
            message += f"\n\n诊断编号: {diagnostic_id}\n日志: {log_path}"
        title = "WorkbookLens 启动失败"
    else:
        message = (
            "WorkbookLens could not start. Make sure Microsoft Edge WebView2 Runtime is "
            "installed, then open the application again."
        )
        if log_path is not None:
            message += f"\n\nDiagnostic ID: {diagnostic_id}\nLog: {log_path}"
        title = "WorkbookLens startup failed"

    if sys.platform != "win32":
        return
    try:
        user32 = ctypes.windll.user32
        user32.MessageBoxW(None, message, title, 0x10)
    except (AttributeError, OSError):
        return


def _load_webview() -> Any:
    try:
        webview = importlib.import_module("webview")
        importlib.import_module("webview.platforms.winforms")
        return webview
    except (ImportError, OSError) as exc:
        raise RuntimeError("The native webview runtime could not be loaded") from exc


def _require_edgechromium_renderer(webview: Any) -> None:
    try:
        renderer = webview.platforms.winforms.renderer
    except AttributeError as exc:
        raise RuntimeError("The native webview renderer could not be identified") from exc
    if renderer != "edgechromium":
        raise RuntimeError("WorkbookLens requires the EdgeChromium webview renderer")


def _schedule_smoke_close(window: Any, shown: threading.Event) -> None:
    shown.set()

    def close_window() -> None:
        time.sleep(0.5)
        window.destroy()

    threading.Thread(
        target=close_window,
        name="workbooklens-desktop-smoke-close",
        daemon=True,
    ).start()


def _run_window(webview: Any, server: RunningLocalUI, root: Path, *, smoke: bool) -> None:
    webview.settings["ALLOW_DOWNLOADS"] = True
    window = webview.create_window(
        _WINDOW_TITLE,
        server.url,
        width=1280,
        height=820,
        min_size=(900, 620),
        resizable=True,
        background_color="#F6F8FB",
        text_select=True,
        zoomable=True,
    )
    if window is None:
        raise RuntimeError("The native WorkbookLens window could not be created")

    shown = threading.Event()
    if smoke:
        window.events.shown += lambda: _schedule_smoke_close(window, shown)

    storage_path = root / "webview"
    storage_path.mkdir(parents=True, exist_ok=True)
    webview.start(
        gui="edgechromium",
        debug=False,
        private_mode=False,
        storage_path=str(storage_path),
    )
    _require_edgechromium_renderer(webview)
    if smoke and not shown.is_set():
        raise RuntimeError("The native WorkbookLens window never became visible")


def run_desktop(argv: Sequence[str] | None = None) -> int:
    """Run the native desktop window and own its local server until the window closes."""

    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments not in ([], [_SMOKE_ARGUMENT]):
        return 2
    smoke = arguments == [_SMOKE_ARGUMENT]
    root = _application_data_root()
    log_path: Path | None = None
    server: RunningLocalUI | None = None
    try:
        log_path = _configure_logging(root)
        logger = logging.getLogger("workbooklens.desktop")
        webview = _load_webview()
        _require_edgechromium_renderer(webview)
        server = start_local_ui_server(
            port=DEFAULT_PORT,
            fallback_port=True,
            status=logger.info,
        )
        _run_window(webview, server, root, smoke=smoke)
        return 0
    except BaseException:
        diagnostic_id = uuid.uuid4().hex[:12]
        logging.getLogger("workbooklens.desktop").exception(
            "Desktop startup/runtime failure diagnostic_id=%s",
            diagnostic_id,
        )
        if not smoke:
            _show_native_error(diagnostic_id, log_path)
        return 1
    finally:
        if server is not None:
            try:
                server.stop()
            except BaseException:
                logging.getLogger("workbooklens.desktop").exception(
                    "Desktop local server shutdown failed"
                )


def main() -> None:
    raise SystemExit(run_desktop())


__all__ = ["main", "run_desktop"]
