"""Loopback-only server launcher shared by the CLI and portable builds."""

from __future__ import annotations

import errno
import socket
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from collections.abc import Callable
from dataclasses import dataclass

import uvicorn

from workbooklens.exceptions import UsageError
from workbooklens.web.app import create_app

LOOPBACK_HOST = "127.0.0.1"
DEFAULT_PORT = 8765
HEALTH_TIMEOUT_SECONDS = 10.0
HEALTH_POLL_SECONDS = 0.1

StatusCallback = Callable[[str], None]


@dataclass(frozen=True, slots=True)
class LocalBinding:
    """A reserved loopback socket and its user-facing port details."""

    socket: socket.socket
    port: int
    requested_port: int
    fell_back: bool


def _new_loopback_socket(port: int) -> socket.socket:
    listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            listener.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        listener.bind((LOOPBACK_HOST, port))
    except BaseException:
        listener.close()
        raise
    return listener


def _address_in_use(exc: OSError) -> bool:
    return exc.errno == errno.EADDRINUSE or getattr(exc, "winerror", None) == 10048


def bind_loopback_socket(port: int, *, fallback_port: bool = False) -> LocalBinding:
    """Reserve a loopback port, optionally falling back only after an in-use error."""

    try:
        listener = _new_loopback_socket(port)
    except OSError as exc:
        if fallback_port and port != 0 and _address_in_use(exc):
            try:
                listener = _new_loopback_socket(0)
            except OSError as fallback_exc:
                raise UsageError(
                    f"Unable to reserve a fallback loopback port: {fallback_exc}"
                ) from fallback_exc
            actual_port = int(listener.getsockname()[1])
            return LocalBinding(
                socket=listener,
                port=actual_port,
                requested_port=port,
                fell_back=True,
            )
        if _address_in_use(exc):
            raise UsageError(
                f"Local port {port} is already in use. Close the other process, choose --port 0, "
                "or enable --fallback-port."
            ) from exc
        raise UsageError(f"Unable to bind the local UI to {LOOPBACK_HOST}:{port}: {exc}") from exc

    actual_port = int(listener.getsockname()[1])
    return LocalBinding(
        socket=listener,
        port=actual_port,
        requested_port=port,
        fell_back=False,
    )


def _health_ready(opener: urllib.request.OpenerDirector, url: str, timeout: float) -> bool:
    # The caller constructs this URL from the fixed loopback host and a bound numeric port.
    request = urllib.request.Request(f"{url}/health", method="GET")  # noqa: S310
    try:
        with opener.open(request, timeout=timeout) as response:
            return bool(response.status == 200 and response.read(256).strip() == b'{"status":"ok"}')
    except (OSError, urllib.error.URLError):
        return False


def _open_browser_when_ready(
    url: str,
    *,
    status: StatusCallback,
    timeout: float = HEALTH_TIMEOUT_SECONDS,
    poll_interval: float = HEALTH_POLL_SECONDS,
) -> None:
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    deadline = time.monotonic() + timeout
    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            status(f"The local UI did not become ready in time. Open it manually: {url}")
            return
        if _health_ready(opener, url, min(remaining, 0.5)):
            break
        time.sleep(min(poll_interval, remaining))

    try:
        opened = webbrowser.open(url, new=2, autoraise=True)
    except Exception as exc:
        status(f"The browser could not be opened automatically ({exc}). Open this URL: {url}")
        return
    if not opened:
        status(f"The browser could not be opened automatically. Open this URL: {url}")


def run_local_ui(
    *,
    port: int = DEFAULT_PORT,
    max_file_bytes: int = 100 * 1024 * 1024,
    open_browser: bool = False,
    fallback_port: bool = False,
    status: StatusCallback = print,
) -> None:
    """Run the local UI on a pre-bound loopback socket until shutdown."""

    binding = bind_loopback_socket(port, fallback_port=fallback_port)
    try:
        local_app = create_app(max_file_bytes=max_file_bytes)
        url = f"http://{LOOPBACK_HOST}:{binding.port}"
        if binding.fell_back:
            status(f"Local port {binding.requested_port} is in use; using {binding.port} instead.")
        status(f"WorkbookLens local UI: {url}")
        status("Press Ctrl+C in this window to stop WorkbookLens.")

        config = uvicorn.Config(
            local_app,
            host=LOOPBACK_HOST,
            port=binding.port,
            log_level="info",
            loop="asyncio",
            http="h11",
            ws="none",
            lifespan="on",
            proxy_headers=False,
        )
        server = uvicorn.Server(config)
        if open_browser:
            threading.Thread(
                target=_open_browser_when_ready,
                kwargs={"url": url, "status": status},
                name="workbooklens-browser",
                daemon=True,
            ).start()
        server.run(sockets=[binding.socket])
    finally:
        binding.socket.close()
