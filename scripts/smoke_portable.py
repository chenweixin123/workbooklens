from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import TextIO

if __package__:
    from .check_portable_artifact import (
        PortableArtifactError,
        extract_checked_artifact,
    )
else:
    from check_portable_artifact import (
        PortableArtifactError,
        extract_checked_artifact,
    )


class PortableSmokeError(RuntimeError):
    pass


def configure_utf8_stdio(
    *,
    stdout: TextIO | None = None,
    stderr: TextIO | None = None,
) -> None:
    streams = (
        sys.stdout if stdout is None else stdout,
        sys.stderr if stderr is None else stderr,
    )
    for stream in streams:
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="backslashreplace")
        except (AttributeError, ValueError):
            continue


def _safe_print(message: str, *, stream: TextIO, flush: bool = False) -> None:
    try:
        print(message, file=stream, flush=flush)
    except UnicodeEncodeError:
        encoding = stream.encoding or "ascii"
        escaped = message.encode(encoding, errors="backslashreplace").decode(encoding)
        print(escaped, file=stream, flush=flush)


def _log_command(command: list[str], *, stream: TextIO | None = None) -> None:
    _safe_print(
        f"+ {subprocess.list2cmdline(command)}",
        stream=sys.stdout if stream is None else stream,
        flush=True,
    )


def sanitized_windows_environment(source: dict[str, str] | None = None) -> dict[str, str]:
    source = dict(os.environ if source is None else source)
    system_root = source.get("SystemRoot") or source.get("WINDIR")
    if not system_root:
        raise PortableSmokeError("SystemRoot/WINDIR is required for the portable smoke test")
    system = Path(system_root)
    source["SystemRoot"] = str(system)
    source["WINDIR"] = str(system)
    source["PATH"] = ";".join(
        [
            str(system / "System32"),
            str(system),
            str(system / "System32" / "Wbem"),
        ]
    )
    for key in tuple(source):
        folded = key.casefold()
        if folded.startswith(("python", "pip_", "uv_", "conda")) or folded in {
            "_old_virtual_path",
            "pyenv",
            "virtual_env",
            "virtual_env_prompt",
        }:
            source.pop(key, None)
    source["PYTHONNOUSERSITE"] = "1"
    return source


def _run_cli(
    executable: Path,
    arguments: list[str],
    *,
    cwd: Path,
    env: dict[str, str],
    timeout: float,
) -> str:
    command = [str(executable), *arguments]
    _log_command(command)
    try:
        completed = subprocess.run(  # noqa: S603 - executable is the validated artifact.
            command,
            cwd=cwd,
            env=env,
            check=True,
            timeout=timeout,
            text=True,
            encoding="utf-8",
            errors="strict",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
    except subprocess.CalledProcessError as exc:
        raise PortableSmokeError(
            f"portable command failed ({exc.returncode}): {command!r}\n{exc.stdout or ''}"
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise PortableSmokeError(f"portable command timed out: {command!r}") from exc
    except UnicodeError as exc:
        raise PortableSmokeError(
            f"portable command did not emit valid UTF-8: {command!r}: {exc}"
        ) from exc
    except OSError as exc:
        raise PortableSmokeError(f"cannot run portable command {command!r}: {exc}") from exc
    return completed.stdout


def contains_text_ignoring_line_wraps(output: str, expected: str) -> bool:
    """Match text that Rich may wrap across redirected-output lines."""

    return expected in output.replace("\r", "").replace("\n", "")


def _assert_no_python_on_path(env: dict[str, str], cwd: Path) -> None:
    where = Path(env["SystemRoot"]) / "System32" / "where.exe"
    completed = subprocess.run(  # noqa: S603 - where.exe comes from SystemRoot.
        [where, "python.exe"],
        cwd=cwd,
        env=env,
        check=False,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    if completed.returncode == 0:
        raise PortableSmokeError(
            f"sanitized PATH still resolves python.exe: {completed.stdout.strip()}"
        )


def _find_free_loopback_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


def parse_windows_listeners(output: str, port: int) -> list[tuple[str, int]]:
    return [
        (host, process_id)
        for host, local_port, process_id in parse_windows_listener_endpoints(output)
        if local_port == port
    ]


def parse_windows_listener_endpoints(output: str) -> list[tuple[str, int, int]]:
    listeners: list[tuple[str, int, int]] = []
    for line in output.splitlines():
        fields = line.split()
        if len(fields) < 5 or fields[0].upper() != "TCP":
            continue
        if fields[3].upper() != "LISTENING":
            continue
        local = fields[1]
        try:
            host, port_text = local.rsplit(":", 1)
            local_port = int(port_text)
            process_id = int(fields[4])
        except (ValueError, IndexError):
            continue
        listeners.append((host.strip("[]"), local_port, process_id))
    return listeners


def _netstat_listeners(env: dict[str, str], cwd: Path, port: int) -> list[tuple[str, int]]:
    netstat = Path(env["SystemRoot"]) / "System32" / "netstat.exe"
    completed = subprocess.run(  # noqa: S603 - netstat.exe comes from SystemRoot.
        [netstat, "-ano", "-p", "TCP"],
        cwd=cwd,
        env=env,
        check=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    return parse_windows_listeners(completed.stdout, port)


def _netstat_listener_endpoints(
    env: dict[str, str],
    cwd: Path,
) -> list[tuple[str, int, int]]:
    netstat = Path(env["SystemRoot"]) / "System32" / "netstat.exe"
    completed = subprocess.run(  # noqa: S603 - netstat.exe comes from SystemRoot.
        [netstat, "-ano", "-p", "TCP"],
        cwd=cwd,
        env=env,
        check=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    return parse_windows_listener_endpoints(completed.stdout)


def _http_open(
    opener: urllib.request.OpenerDirector,
    url: str,
    *,
    data: bytes | None = None,
    headers: dict[str, str] | None = None,
    timeout: float = 2.0,
) -> tuple[int, bytes, str]:
    if not url.startswith("http://127.0.0.1:"):
        raise PortableSmokeError(f"refusing non-loopback smoke URL: {url!r}")
    request_headers = {"User-Agent": "WorkbookLens-smoke"}
    if headers:
        request_headers.update(headers)
    request = urllib.request.Request(  # noqa: S310 - URL is constrained to loopback HTTP.
        url,
        data=data,
        headers=request_headers,
    )
    with opener.open(request, timeout=timeout) as response:
        return (
            int(response.status),
            response.read(),
            response.headers.get_content_type(),
        )


def _http_get(url: str, timeout: float = 2.0) -> tuple[int, bytes, str]:
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    return _http_open(opener, url, timeout=timeout)


def _extract_csrf_token(html: bytes) -> str:
    match = re.search(rb'name="csrf_token" value="([A-Za-z0-9_-]{43,})"', html)
    if match is None:
        raise PortableSmokeError("portable server home page did not expose a valid CSRF token")
    return match.group(1).decode("ascii")


def _build_multipart_upload(workbook: Path, csrf_token: str) -> tuple[bytes, str]:
    if re.fullmatch(r"[A-Za-z0-9_-]{43,}", csrf_token) is None:
        raise PortableSmokeError("refusing to build multipart upload with an invalid CSRF token")
    payload = workbook.read_bytes()
    boundary = "----WorkbookLensSmokeBoundary7MA4YWxkTrZu0gW"
    boundary_bytes = boundary.encode("ascii")
    if boundary_bytes in payload:
        raise PortableSmokeError("multipart boundary unexpectedly occurs in the demo workbook")
    body = b"".join(
        [
            b"--" + boundary_bytes + b"\r\n",
            b'Content-Disposition: form-data; name="csrf_token"\r\n\r\n',
            csrf_token.encode("ascii") + b"\r\n",
            b"--" + boundary_bytes + b"\r\n",
            b'Content-Disposition: form-data; name="workbook"; filename="demo.xlsx"\r\n',
            (
                b"Content-Type: application/vnd.openxmlformats-officedocument."
                b"spreadsheetml.sheet\r\n\r\n"
            ),
            payload,
            b"\r\n--" + boundary_bytes + b"--\r\n",
        ]
    )
    return body, f"multipart/form-data; boundary={boundary}"


def _wait_for_server(
    process: subprocess.Popen[str],
    *,
    env: dict[str, str],
    cwd: Path,
    port: int,
    timeout: float,
) -> None:
    deadline = time.monotonic() + timeout
    last_error = "server did not respond"
    while time.monotonic() < deadline:
        if process.poll() is not None:
            output = process.communicate()[0]
            raise PortableSmokeError(
                f"portable server exited before becoming ready ({process.returncode})\n{output}"
            )
        try:
            status, body, _ = _http_get(f"http://127.0.0.1:{port}/health")
            payload = json.loads(body)
            listeners = _netstat_listeners(env, cwd, port)
            if (
                status == 200
                and payload == {"status": "ok"}
                and listeners == [("127.0.0.1", process.pid)]
            ):
                return
            last_error = (
                f"health={status}/{payload!r}, listeners={listeners!r}, expected pid={process.pid}"
            )
        except (
            OSError,
            ValueError,
            json.JSONDecodeError,
            urllib.error.URLError,
        ) as exc:
            last_error = str(exc)
        time.sleep(0.25)
    raise PortableSmokeError(f"portable server did not become ready: {last_error}")


def _wait_for_port_release(port: int, timeout: float) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            try:
                probe.bind(("127.0.0.1", port))
            except OSError:
                time.sleep(0.25)
                continue
            return
    raise PortableSmokeError(f"loopback port {port} was not released")


def _directory_snapshot(root: Path) -> dict[str, tuple[int, str]]:
    snapshot: dict[str, tuple[int, str]] = {}
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            raise PortableSmokeError(f"portable installation contains a symlink: {path}")
        if not path.is_file():
            continue
        digest = hashlib.sha256()
        try:
            with path.open("rb") as stream:
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(chunk)
            size = path.stat().st_size
        except OSError as exc:
            raise PortableSmokeError(
                f"cannot snapshot portable installation file {path}: {exc}"
            ) from exc
        snapshot[path.relative_to(root).as_posix()] = (size, digest.hexdigest())
    return snapshot


def _stop_server(
    process: subprocess.Popen[str],
    *,
    port: int,
) -> None:
    process.send_signal(signal.CTRL_BREAK_EVENT)  # type: ignore[attr-defined]
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired as exc:
        raise PortableSmokeError("portable server did not stop after Ctrl+Break") from exc
    _wait_for_port_release(port, timeout=10)


def _smoke_server(
    executable: Path,
    *,
    cwd: Path,
    env: dict[str, str],
    workbook: Path,
    expected_version: str,
    timeout: float,
) -> None:
    port = _find_free_loopback_port()
    command = [str(executable), "serve", "--port", str(port)]
    creationflags = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
    process = subprocess.Popen(  # noqa: S603 - executable is the validated artifact.
        command,
        cwd=cwd,
        env=env,
        creationflags=creationflags,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    stopped_by_signal = False
    try:
        _wait_for_server(
            process,
            env=env,
            cwd=cwd,
            port=port,
            timeout=timeout,
        )
        base_url = f"http://127.0.0.1:{port}"
        opener = urllib.request.build_opener(
            urllib.request.ProxyHandler({}),
            urllib.request.HTTPCookieProcessor(),
        )
        status, body, content_type = _http_open(opener, f"{base_url}/")
        if status != 200 or b"WorkbookLens" not in body or content_type != "text/html":
            raise PortableSmokeError(
                "portable server home page did not return the WorkbookLens HTML UI"
            )
        csrf_token = _extract_csrf_token(body)
        upload, upload_type = _build_multipart_upload(workbook, csrf_token)
        scan_status, scan_body, scan_type = _http_open(
            opener,
            f"{base_url}/scan",
            data=upload,
            headers={
                "Content-Type": upload_type,
                "Origin": base_url,
            },
            timeout=timeout,
        )
        session_match = re.search(rb"/sessions/([A-Za-z0-9_-]+)/apply", scan_body)
        if scan_status != 200 or scan_type != "text/html" or session_match is None:
            raise PortableSmokeError(
                "portable server did not scan a multipart workbook upload into a review session"
            )
        session_id = session_match.group(1).decode("ascii")
        plan_status, plan_body, plan_type = _http_open(
            opener,
            f"{base_url}/sessions/{session_id}/plan",
            timeout=timeout,
        )
        try:
            plan_payload = json.loads(plan_body)
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise PortableSmokeError("portable server returned an invalid repair plan") from exc
        if (
            plan_status != 200
            or plan_type != "application/json"
            or plan_payload.get("tool_version") != expected_version
        ):
            raise PortableSmokeError(
                f"portable server returned an unexpected repair plan: {plan_payload!r}"
            )
        report_status, report_body, report_type = _http_open(
            opener,
            f"{base_url}/sessions/{session_id}/report",
            timeout=timeout,
        )
        if report_status != 200 or report_type != "text/html" or b"WorkbookLens" not in report_body:
            raise PortableSmokeError("portable server did not return the uploaded workbook report")
        listeners = _netstat_listeners(env, cwd, port)
        if listeners != [("127.0.0.1", process.pid)]:
            raise PortableSmokeError(
                f"portable server is not loopback-only or has unexpected owner: {listeners!r}"
            )

        _stop_server(process, port=port)
        stopped_by_signal = True
    finally:
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
        output = process.communicate()[0]
        if not stopped_by_signal:
            print("portable server output:\n" + output, file=sys.stderr)


def _assert_occupied_port_fails_without_fallback(
    executable: Path,
    *,
    cwd: Path,
    env: dict[str, str],
    port: int,
    timeout: float,
) -> None:
    command = [str(executable), "serve", "--port", str(port)]
    try:
        completed = subprocess.run(  # noqa: S603 - executable is the validated artifact.
            command,
            cwd=cwd,
            env=env,
            check=False,
            timeout=timeout,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
    except subprocess.TimeoutExpired as exc:
        raise PortableSmokeError(
            "portable server hung on an occupied port without fallback"
        ) from exc
    except OSError as exc:
        raise PortableSmokeError(f"cannot run portable server: {exc}") from exc
    if completed.returncode == 0:
        raise PortableSmokeError(
            "portable server unexpectedly succeeded on an occupied port without --fallback-port"
        )


def _wait_for_fallback_server(
    process: subprocess.Popen[str],
    *,
    env: dict[str, str],
    cwd: Path,
    occupied_port: int,
    timeout: float,
) -> int:
    deadline = time.monotonic() + timeout
    last_error = "fallback server did not create a listener"
    while time.monotonic() < deadline:
        if process.poll() is not None:
            output = process.communicate()[0]
            raise PortableSmokeError(
                "portable fallback server exited before becoming ready "
                f"({process.returncode})\n{output}"
            )
        try:
            owned = [
                endpoint
                for endpoint in _netstat_listener_endpoints(env, cwd)
                if endpoint[2] == process.pid
            ]
            if any(host != "127.0.0.1" for host, _, _ in owned):
                raise PortableSmokeError(f"fallback server is not loopback-only: {owned!r}")
            candidates = [
                port for host, port, _ in owned if host == "127.0.0.1" and port != occupied_port
            ]
            if len(candidates) == 1 and len(owned) == 1:
                port = candidates[0]
                health_status, health_body, _ = _http_get(f"http://127.0.0.1:{port}/health")
                home_status, home_body, home_type = _http_get(f"http://127.0.0.1:{port}/")
                if (
                    health_status == 200
                    and json.loads(health_body) == {"status": "ok"}
                    and home_status == 200
                    and b"WorkbookLens" in home_body
                    and home_type == "text/html"
                ):
                    return port
            last_error = f"owned listeners={owned!r}, expected pid={process.pid}"
        except (
            OSError,
            ValueError,
            json.JSONDecodeError,
            urllib.error.URLError,
        ) as exc:
            last_error = str(exc)
        time.sleep(0.25)
    raise PortableSmokeError(f"portable fallback server did not become ready: {last_error}")


def _smoke_occupied_port_fallback(
    executable: Path,
    *,
    cwd: Path,
    env: dict[str, str],
    timeout: float,
) -> None:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as occupied:
        occupied.bind(("127.0.0.1", 0))
        occupied.listen(1)
        occupied_port = int(occupied.getsockname()[1])
        _assert_occupied_port_fails_without_fallback(
            executable,
            cwd=cwd,
            env=env,
            port=occupied_port,
            timeout=timeout,
        )

        command = [
            str(executable),
            "serve",
            "--port",
            str(occupied_port),
            "--fallback-port",
        ]
        creationflags = subprocess.CREATE_NEW_PROCESS_GROUP  # type: ignore[attr-defined]
        process = subprocess.Popen(  # noqa: S603 - executable is the validated artifact.
            command,
            cwd=cwd,
            env=env,
            creationflags=creationflags,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
        stopped_by_signal = False
        try:
            fallback_port = _wait_for_fallback_server(
                process,
                env=env,
                cwd=cwd,
                occupied_port=occupied_port,
                timeout=timeout,
            )
            _stop_server(process, port=fallback_port)
            stopped_by_signal = True
        finally:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=5)
            output = process.communicate()[0]
            if not stopped_by_signal:
                print("portable fallback server output:\n" + output, file=sys.stderr)


def smoke_portable(
    archive: Path,
    *,
    expected_version: str,
    repository_root: Path,
    timeout: float,
    keep_temp: bool,
) -> Path | None:
    if os.name != "nt":
        raise PortableSmokeError("portable executable smoke tests require Windows")
    temporary = Path(
        tempfile.mkdtemp(prefix="WorkbookLens portable \u4e2d\u6587 \u7a7a\u683c ")
    ).resolve()
    try:
        root = extract_checked_artifact(
            archive,
            temporary,
            expected_version=expected_version,
            repository_root=repository_root,
        )
        executable = root / "WorkbookLens.exe"
        env = sanitized_windows_environment()
        _assert_no_python_on_path(env, root)
        installation_snapshot = _directory_snapshot(root)

        version_output = _run_cli(executable, ["--version"], cwd=root, env=env, timeout=timeout)
        if expected_version not in version_output:
            raise PortableSmokeError(
                f"--version output does not contain {expected_version!r}: {version_output!r}"
            )
        _run_cli(executable, ["--help"], cwd=root, env=env, timeout=timeout)

        work = temporary / "\u6d4b\u8bd5 \u8f93\u51fa with spaces"
        work.mkdir()
        demo = work / "demo"
        demo_output = _run_cli(
            executable,
            ["demo", "--out", str(demo)],
            cwd=root,
            env=env,
            timeout=timeout,
        )
        expected_demo_message = f"Demo complete in {demo}"
        if not contains_text_ignoring_line_wraps(demo_output, expected_demo_message):
            raise PortableSmokeError(
                f"demo output did not preserve its non-ASCII path: {demo_output!r}"
            )
        before = demo / "before.xlsx"
        after = demo / "after.xlsx"
        for required in (
            before,
            after,
            demo / "repair-plan.json",
            demo / "diff.html",
        ):
            if not required.is_file():
                raise PortableSmokeError(f"demo did not create expected file: {required}")

        scan_dir = work / "scan report"
        _run_cli(
            executable,
            ["scan", str(before), "--out", str(scan_dir)],
            cwd=root,
            env=env,
            timeout=timeout,
        )
        if not any(scan_dir.rglob("*")):
            raise PortableSmokeError("scan command produced no report files")

        plan = work / "repair plan.json"
        _run_cli(
            executable,
            ["plan", str(before), "--out", str(plan)],
            cwd=root,
            env=env,
            timeout=timeout,
        )
        if not plan.is_file():
            raise PortableSmokeError("plan command did not create a plan")

        applied = work / "applied workbook.xlsx"
        _run_cli(
            executable,
            [
                "apply",
                str(before),
                str(plan),
                "--out",
                str(applied),
                "--safe-only",
            ],
            cwd=root,
            env=env,
            timeout=timeout,
        )
        if not applied.is_file():
            raise PortableSmokeError("apply command did not create a workbook")

        diff = work / "semantic diff.html"
        _run_cli(
            executable,
            ["diff", str(before), str(applied), "--out", str(diff)],
            cwd=root,
            env=env,
            timeout=timeout,
        )
        if not diff.is_file():
            raise PortableSmokeError("diff command did not create an HTML report")

        _smoke_server(
            executable,
            cwd=root,
            env=env,
            workbook=before,
            expected_version=expected_version,
            timeout=timeout,
        )
        _smoke_occupied_port_fallback(
            executable,
            cwd=root,
            env=env,
            timeout=timeout,
        )
        if _directory_snapshot(root) != installation_snapshot:
            raise PortableSmokeError(
                "portable commands modified files in the installation directory"
            )
        if keep_temp:
            print(f"portable smoke workspace retained at {temporary}")
            return temporary
        return None
    finally:
        if not keep_temp and temporary.exists():
            shutil.rmtree(temporary)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run frozen-process smoke tests against a WorkbookLens portable ZIP."
    )
    parser.add_argument("archive", type=Path)
    parser.add_argument("--expected-version", required=True)
    parser.add_argument(
        "--repository-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    parser.add_argument("--timeout", type=float, default=60.0)
    parser.add_argument("--keep-temp", action="store_true")
    return parser


def main() -> int:
    configure_utf8_stdio()
    args = _build_parser().parse_args()
    try:
        smoke_portable(
            args.archive.resolve(),
            expected_version=args.expected_version,
            repository_root=args.repository_root.resolve(),
            timeout=args.timeout,
            keep_temp=args.keep_temp,
        )
    except (PortableArtifactError, PortableSmokeError, OSError) as exc:
        print(f"portable smoke failed: {exc}", file=sys.stderr)
        return 1
    print("portable smoke passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
