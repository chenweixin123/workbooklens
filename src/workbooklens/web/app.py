"""Accessible server-rendered local workflow with bounded untrusted uploads."""

from __future__ import annotations

import json
import logging
import secrets
import shutil
import tempfile
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from functools import partial
from pathlib import Path
from typing import Any, Literal
from urllib.parse import urlsplit

from fastapi import FastAPI, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, PlainTextResponse, RedirectResponse
from jinja2 import Environment, select_autoescape
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import MutableHeaders
from starlette.types import ASGIApp, Message, Receive, Scope, Send

from workbooklens import __version__, conversion
from workbooklens.diff import compare_workbooks, write_diff_report
from workbooklens.exceptions import WorkbookLensError
from workbooklens.i18n import (
    DEFAULT_LANGUAGE,
    Language,
    UserFacingError,
    localize_exception,
    localize_finding,
    localize_patch,
    localized_error,
    new_diagnostic_id,
    normalize_language,
    translate,
)
from workbooklens.models import Finding, PatchPlan, PatchResult, PatchRisk, WorkbookDiff
from workbooklens.ooxml.safety import PackageLimits
from workbooklens.repair import apply_patch_plan, build_patch_plan
from workbooklens.repair.planning import write_patch_plan
from workbooklens.reports import write_scan_report
from workbooklens.scanner import ScanResult, scan_workbook
from workbooklens.utils import write_json
from workbooklens.web.templates import template_loader

LOCAL_HOSTS = frozenset({"localhost", "127.0.0.1"})
CSRF_COOKIE_NAME = "workbooklens_csrf"
LANGUAGE_COOKIE_NAME = "workbooklens_language"
LANGUAGE_COOKIE_MAX_AGE = 365 * 24 * 60 * 60
MAX_MULTIPART_OVERHEAD_BYTES = 16 * 1024
MAX_FORM_BODY_BYTES = 1024 * 1024
UNSAFE_METHODS = frozenset({"POST", "PUT", "PATCH", "DELETE"})
SECURITY_HEADERS = {
    "Cache-Control": "no-store",
    "Content-Security-Policy": (
        "default-src 'none'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'; "
        "style-src 'unsafe-inline'; script-src 'unsafe-inline'; img-src 'self' data:"
    ),
    "Cross-Origin-Opener-Policy": "same-origin",
    "Cross-Origin-Resource-Policy": "same-origin",
    "Permissions-Policy": "camera=(), microphone=(), geolocation=(), payment=(), usb=()",
    "Referrer-Policy": "no-referrer",
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
}

LOGGER = logging.getLogger(__name__)


def _request_operation(
    request: Request,
) -> Literal["conversion", "scan", "repair", "download", "request"]:
    path = request.url.path
    if path == "/convert":
        return "conversion"
    if path == "/scan":
        return "scan"
    if path.endswith("/apply"):
        return "repair"
    if path.startswith("/sessions/"):
        return "download"
    return "request"


class _LocalizedHTTPException(HTTPException):
    def __init__(self, status_code: int, detail: str, error_key: str) -> None:
        super().__init__(status_code=status_code, detail=detail)
        self.error_key = error_key


def _http_error(status_code: int, detail: str, error_key: str) -> HTTPException:
    return _LocalizedHTTPException(status_code, detail, error_key)


def _conversion_error_key(exc: BaseException) -> str:
    explicit = getattr(exc, "error_key", None)
    if isinstance(explicit, str) and explicit.startswith("conversion."):
        return explicit
    message = str(exc).lower()
    if any(
        marker in message
        for marker in (
            "accepts only .xls",
            "not a recognized binary excel .xls",
            "legacy workbook does not exist",
        )
    ):
        return "conversion.invalid_input"
    if "no local .xls converter is available" in message:
        return "conversion.unavailable"
    if "timed out" in message or "did not terminate in time" in message:
        return "conversion.timeout"
    if any(
        marker in message
        for marker in (
            "verified macro-free .xlsx",
            "did not create the expected .xlsx",
            "output is missing",
            "output contains",
            "output has no cell style table",
        )
    ):
        return "conversion.output_invalid"
    return "conversion.all_providers_failed"


@dataclass(slots=True)
class WebSession:
    """Server-owned paths and immutable scan/plan data for one local upload."""

    session_id: str
    filename: str
    source: Path
    plan_path: Path
    plan: PatchPlan
    scan: ScanResult
    findings: list[Finding]
    language: Language
    reports: dict[Language, Path]
    fixed: Path | None = None
    result: PatchResult | None = None
    semantic_diff: WorkbookDiff | None = None
    diff_reports: dict[Language, Path] | None = None
    apply_report: Path | None = None


@dataclass(slots=True)
class _LimitedReceive:
    receive: Receive
    maximum: int
    received: int = 0
    exceeded: bool = False

    async def __call__(self) -> Message:
        if self.exceeded:
            return {"type": "http.disconnect"}
        message = await self.receive()
        if message["type"] == "http.request":
            self.received += len(message.get("body", b""))
            if self.received > self.maximum:
                self.exceeded = True
                return {"type": "http.disconnect"}
        return message


def _environment() -> Environment:
    return Environment(
        loader=template_loader(),
        autoescape=select_autoescape(enabled_extensions=("html",), default=True),
    )


def _preferred_language(request: Request) -> Language:
    for item in request.headers.get("accept-language", "").split(","):
        tag = item.partition(";")[0].strip().lower()
        if tag.startswith("zh"):
            return "zh-CN"
        if tag.startswith("en"):
            return "en"
    return DEFAULT_LANGUAGE


def _request_language(
    request: Request,
    explicit: str | None = None,
    *,
    fallback: Language | None = None,
) -> Language:
    configured = getattr(request.app.state, "default_language", None)
    browser_default = fallback or normalize_language(
        configured,
        fallback=_preferred_language(request),
    )
    cookie_value = request.cookies.get(LANGUAGE_COOKIE_NAME)
    cookie_language = normalize_language(cookie_value, fallback=browser_default)
    return normalize_language(explicit, fallback=cookie_language)


def _set_language_cookie(response: HTMLResponse | RedirectResponse, language: Language) -> None:
    response.set_cookie(
        LANGUAGE_COOKIE_NAME,
        language,
        max_age=LANGUAGE_COOKIE_MAX_AGE,
        httponly=True,
        samesite="lax",
        path="/",
    )


def _display_value(value: Any, *, empty_label: str) -> str:
    if value is None:
        return empty_label
    if isinstance(value, str):
        return value
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    except (TypeError, ValueError):
        return str(value)


def _render_page(
    environment: Environment,
    template_name: str,
    language: Language,
    *,
    page_title: str,
    language_action: str,
    view: str,
    **context: Any,
) -> str:
    translator = partial(translate, language=language)
    return environment.get_template(template_name).render(
        language=language,
        page_title=page_title,
        language_action=language_action,
        view=view,
        version=__version__,
        t=translator,
        display_value=partial(
            _display_value,
            empty_label=translator("web.value_empty"),
        ),
        **context,
    )


def _language_response(
    html: str,
    language: Language,
    *,
    status_code: int = 200,
) -> HTMLResponse:
    response = HTMLResponse(html, status_code=status_code)
    _set_language_cookie(response, language)
    return response


def _parse_local_authority(authority: str) -> tuple[str, int | None] | None:
    if not authority or authority != authority.strip():
        return None
    host, separator, port_text = authority.partition(":")
    host = host.lower()
    if host not in LOCAL_HOSTS:
        return None
    if not separator:
        return host, None
    if not port_text or not port_text.isascii() or not port_text.isdigit():
        return None
    port = int(port_text)
    if not 1 <= port <= 65535:
        return None
    return host, port


def _origin(value: str) -> tuple[str, str, int] | None:
    try:
        parsed = urlsplit(value)
        port = parsed.port
    except ValueError:
        return None
    scheme = parsed.scheme.lower()
    host = (parsed.hostname or "").lower()
    if (
        scheme not in {"http", "https"}
        or host not in LOCAL_HOSTS
        or parsed.username is not None
        or parsed.password is not None
    ):
        return None
    if port is None:
        port = 443 if scheme == "https" else 80
    if not 1 <= port <= 65535:
        return None
    return scheme, host, port


def _allowed_form_source(value: str, request_origin: tuple[str, str, int] | None) -> bool:
    if value == "null":
        return True
    submitted_origin = _origin(value)
    if request_origin is None or submitted_origin is None:
        return False
    request_scheme, _request_host, request_port = request_origin
    submitted_scheme, _submitted_host, submitted_port = submitted_origin
    return submitted_scheme == request_scheme and submitted_port == request_port


def _validate_post_request(request: Request, csrf_token: str) -> None:
    request_origin = _origin(str(request.url))
    for header in ("origin", "referer"):
        value = request.headers.get(header)
        if value is not None and not _allowed_form_source(value, request_origin):
            raise _http_error(
                403,
                "Cross-origin form submission rejected",
                "security.cross_origin",
            )
    expected = str(request.app.state.csrf_token)
    if not csrf_token or not secrets.compare_digest(csrf_token, expected):
        raise _http_error(
            403,
            "Invalid or missing form security token",
            "security.csrf",
        )


def _content_length(scope: Scope) -> int | None:
    values = [
        value.strip()
        for name, value in scope.get("headers", [])
        if name.lower() == b"content-length"
    ]
    if not values:
        return None
    if len(values) != 1 or not values[0] or not values[0].isascii() or not values[0].isdigit():
        raise ValueError("Invalid Content-Length header")
    return int(values[0])


def _request_body_limit(path: str, max_file_bytes: int) -> int:
    if path in {"/convert", "/scan"}:
        return max_file_bytes + MAX_MULTIPART_OVERHEAD_BYTES
    return MAX_FORM_BODY_BYTES


class _LocalRequestGuard:
    """Reject unsafe local requests before FastAPI parses form or multipart bodies."""

    def __init__(self, app: ASGIApp, *, max_file_bytes: int, csrf_token: str) -> None:
        self.app = app
        self.max_file_bytes = max_file_bytes
        self.csrf_token = csrf_token

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        request = Request(scope)
        response_started = False
        limited_receive: _LimitedReceive | None = None

        async def secure_send(message: Message) -> None:
            nonlocal response_started
            if message["type"] == "http.response.start":
                response_started = True
                headers = MutableHeaders(scope=message)
                for name, value in SECURITY_HEADERS.items():
                    headers[name] = value
            await send(message)

        async def guarded_send(message: Message) -> None:
            if limited_receive is not None and limited_receive.exceeded:
                return
            await secure_send(message)

        async def reject(status_code: int, error_key: str) -> None:
            language = normalize_language(
                request.cookies.get(LANGUAGE_COOKIE_NAME),
                fallback=_preferred_language(request),
            )
            error = localized_error(error_key, language)
            response = PlainTextResponse(
                f"{error.code}: {error.message}",
                status_code=status_code,
            )
            await response(scope, receive, secure_send)

        if _parse_local_authority(request.headers.get("host", "")) is None:
            await reject(400, "request.invalid")
            return

        if request.method not in UNSAFE_METHODS:
            await self.app(scope, receive, secure_send)
            return

        request_origin = _origin(str(request.url))
        for header in ("origin", "referer"):
            value = request.headers.get(header)
            if value is not None and not _allowed_form_source(value, request_origin):
                await reject(403, "security.cross_origin")
                return

        csrf_cookie = request.cookies.get(CSRF_COOKIE_NAME, "")
        if not csrf_cookie or not secrets.compare_digest(csrf_cookie, self.csrf_token):
            await reject(403, "security.csrf")
            return

        try:
            declared_length = _content_length(scope)
        except ValueError:
            await reject(400, "request.invalid")
            return

        body_limit = _request_body_limit(request.url.path, self.max_file_bytes)
        content_type = request.headers.get("content-type", "").partition(";")[0].strip().lower()
        if (
            request.url.path in {"/convert", "/scan"}
            and content_type == "multipart/form-data"
            and declared_length is None
        ):
            await reject(411, "request.invalid")
            return
        if declared_length is not None and declared_length > body_limit:
            await reject(413, "upload.too_large")
            return

        limited_receive = _LimitedReceive(receive, body_limit)
        try:
            await self.app(scope, limited_receive, guarded_send)
        except Exception:
            if not limited_receive.exceeded:
                raise
        if limited_receive.exceeded:
            if response_started:
                raise RuntimeError("Request body limit was exceeded after the response started")
            await reject(413, "upload.too_large")


async def _store_upload(upload: UploadFile, target: Path, maximum: int) -> None:
    size = 0
    with target.open("xb") as handle:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > maximum:
                raise _http_error(
                    413,
                    "Workbook exceeds the configured upload limit",
                    "upload.too_large",
                )
            handle.write(chunk)
    if size == 0:
        raise _http_error(400, "Uploaded workbook is empty", "upload.empty")


def create_app(
    *,
    max_file_bytes: int = 100 * 1024 * 1024,
    language: Language | str | None = None,
) -> FastAPI:
    """Create the local UI. Callers must still bind Uvicorn to 127.0.0.1."""

    sessions: dict[str, WebSession] = {}

    @asynccontextmanager
    async def lifespan(app: FastAPI) -> AsyncIterator[None]:
        temporary = tempfile.TemporaryDirectory(prefix="workbooklens-web-")
        app.state.root = Path(temporary.name)
        yield
        sessions.clear()
        temporary.cleanup()

    app = FastAPI(
        title="WorkbookLens",
        version=__version__,
        docs_url=None,
        redoc_url=None,
        lifespan=lifespan,
    )
    environment = _environment()
    app.state.csrf_token = secrets.token_urlsafe(32)
    app.state.default_language = language
    app.add_middleware(
        _LocalRequestGuard,
        max_file_bytes=max_file_bytes,
        csrf_token=app.state.csrf_token,
    )

    def error_page(
        request: Request,
        exc: Exception,
        status_code: int,
        *,
        diagnostic_id: str | None = None,
    ) -> HTMLResponse:
        language = _request_language(request)
        diagnostic_id = diagnostic_id or new_diagnostic_id()
        error: UserFacingError = localize_exception(
            exc,
            language,
            operation=_request_operation(request),
            diagnostic_id=diagnostic_id,
        )
        log_level = logging.ERROR if status_code >= 500 else logging.WARNING
        LOGGER.log(
            log_level,
            "Local UI request failed [diagnostic_id=%s error_code=%s error_key=%s exception_type=%s]",
            diagnostic_id,
            error.code,
            error.key,
            type(exc).__name__,
        )
        html = _render_page(
            environment,
            "error.html",
            language,
            page_title=error.title,
            language_action="/",
            view="error",
            error=error,
        )
        return _language_response(html, language, status_code=status_code)

    @app.exception_handler(WorkbookLensError)
    async def workbook_error(request: Request, exc: WorkbookLensError) -> HTMLResponse:
        return error_page(request, exc, 400)

    @app.exception_handler(HTTPException)
    async def http_error(request: Request, exc: HTTPException) -> HTMLResponse:
        return error_page(request, exc, exc.status_code)

    @app.exception_handler(Exception)
    async def unexpected_error(request: Request, exc: Exception) -> HTMLResponse:
        diagnostic_id = new_diagnostic_id()
        return error_page(
            request,
            exc,
            500,
            diagnostic_id=diagnostic_id,
        )

    @app.get("/health")
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    @app.get("/", response_class=HTMLResponse)
    async def index(request: Request, lang: str | None = None) -> HTMLResponse:
        language = _request_language(request, lang)
        translator = partial(translate, language=language)
        converter_names = [
            provider.label for provider in conversion.available_conversion_providers()
        ]
        html = _render_page(
            environment,
            "index.html",
            language,
            page_title=translator("web.home_title"),
            language_action="/",
            view="home",
            max_mb=max_file_bytes // 1024**2,
            csrf_token=app.state.csrf_token,
            converter_names=converter_names,
            provider_separator=translator("web.provider_separator"),
        )
        response = _language_response(html, language)
        response.set_cookie(
            CSRF_COOKIE_NAME,
            app.state.csrf_token,
            httponly=True,
            samesite="strict",
            path="/",
        )
        return response

    @app.post("/convert")
    async def convert_upload(
        request: Request,
        legacy_workbook: UploadFile,
        csrf_token: str = Form(default=""),
        language: str = Form(default=""),
    ) -> FileResponse:
        _validate_post_request(request, csrf_token)
        selected_language = _request_language(request, language)
        filename = Path(legacy_workbook.filename or "workbook.xls").name
        if Path(filename).suffix.lower() != ".xls":
            raise _http_error(
                400,
                "Upload must be a binary .xls workbook",
                "conversion.invalid_input",
            )
        conversion_id = secrets.token_urlsafe(18)
        conversion_root = request.app.state.root / f"convert-{conversion_id}"
        conversion_root.mkdir(parents=True)
        source = conversion_root / "input.xls"
        output = conversion_root / "converted.xlsx"
        try:
            await _store_upload(legacy_workbook, source, max_file_bytes)
            result = await run_in_threadpool(
                conversion.convert_xls_to_xlsx,
                source,
                output,
                max_output_bytes=max_file_bytes,
            )
        except BaseException as exc:
            shutil.rmtree(conversion_root, ignore_errors=True)
            if isinstance(exc, HTTPException):
                raise
            if isinstance(exc, WorkbookLensError):
                exc.error_key = _conversion_error_key(exc)
            elif isinstance(exc, Exception):
                raise WorkbookLensError(
                    error_key=_conversion_error_key(exc),
                ) from exc
            raise
        download_name = f"{Path(filename).stem or 'workbook'}.xlsx"
        response = FileResponse(
            result.output,
            media_type=conversion.XLSX_MEDIA_TYPE,
            filename=download_name,
            headers={"X-WorkbookLens-Converter": result.provider.label},
            background=BackgroundTask(shutil.rmtree, conversion_root, ignore_errors=True),
        )
        response.set_cookie(
            LANGUAGE_COOKIE_NAME,
            selected_language,
            max_age=LANGUAGE_COOKIE_MAX_AGE,
            httponly=True,
            samesite="lax",
            path="/",
        )
        return response

    @app.post("/scan")
    async def scan_upload(
        request: Request,
        workbook: UploadFile,
        csrf_token: str = Form(default=""),
        language: str = Form(default=""),
    ) -> RedirectResponse:
        _validate_post_request(request, csrf_token)
        selected_language = _request_language(request, language)
        filename = Path(workbook.filename or "workbook.xlsx").name
        suffix = Path(filename).suffix.lower()
        if suffix not in {".xlsx", ".xlsm"}:
            raise _http_error(
                400,
                "Upload must be .xlsx or .xlsm",
                "upload.invalid_type",
            )
        if len(sessions) >= 20:
            raise _http_error(
                429,
                "Session limit reached; restart the local server",
                "session.limit",
            )
        session_id = secrets.token_urlsafe(18)
        session_root = request.app.state.root / session_id
        session_root.mkdir(parents=True)
        source = session_root / f"input{suffix}"
        try:
            await _store_upload(workbook, source, max_file_bytes)
            limits = PackageLimits(max_file_bytes=max_file_bytes)
            scan = await run_in_threadpool(scan_workbook, source, limits=limits)
            plan = build_patch_plan(scan)
            plan_path = session_root / "repair-plan.json"
            write_patch_plan(plan_path, plan)
        except BaseException:
            shutil.rmtree(session_root, ignore_errors=True)
            raise
        sessions[session_id] = WebSession(
            session_id=session_id,
            filename=filename,
            source=source,
            plan_path=plan_path,
            plan=plan,
            scan=scan,
            findings=scan.findings,
            language=selected_language,
            reports={},
            diff_reports={},
        )
        response = RedirectResponse(
            f"/sessions/{session_id}",
            status_code=303,
        )
        _set_language_cookie(response, selected_language)
        return response

    def session_or_404(session_id: str) -> WebSession:
        session = sessions.get(session_id)
        if session is None:
            raise _http_error(
                404,
                "Local session not found or expired",
                "session.not_found",
            )
        return session

    def reviewable_patches(session: WebSession) -> list[Any]:
        if session.source.suffix.lower() == ".xlsm":
            return []
        return [
            patch
            for patch in session.plan.patches
            if patch.safe_only_eligible or patch.risk == PatchRisk.LAYOUT_REVIEW
        ]

    @app.get("/sessions/{session_id}", response_class=HTMLResponse)
    async def results(
        session_id: str,
        request: Request,
        lang: str | None = None,
    ) -> HTMLResponse:
        session = session_or_404(session_id)
        language = _request_language(request, lang, fallback=session.language)
        session.language = language
        findings = [localize_finding(finding, language) for finding in session.findings]
        patches = [localize_patch(patch, language) for patch in reviewable_patches(session)]
        severity_counts = {"info": 0, "warning": 0, "error": 0, "critical": 0}
        for finding in findings:
            severity_counts[finding.severity.value] += 1
        translator = partial(translate, language=language)
        html = _render_page(
            environment,
            "results.html",
            language,
            page_title=translator("web.results_page_title"),
            language_action=f"/sessions/{session_id}",
            view="results",
            session_id=session_id,
            filename=session.filename,
            findings=findings,
            patches=patches,
            severity_counts=severity_counts,
            csrf_token=app.state.csrf_token,
            has_layout_review=any(patch.risk == PatchRisk.LAYOUT_REVIEW for patch in patches),
        )
        return _language_response(html, language)

    @app.get("/sessions/{session_id}/report")
    async def report(
        session_id: str,
        request: Request,
        lang: str | None = None,
    ) -> FileResponse:
        session = session_or_404(session_id)
        language = _request_language(request, lang, fallback=session.language)
        session.language = language
        report_path = session.reports.get(language)
        if report_path is None:
            report_paths = await run_in_threadpool(
                write_scan_report,
                session.scan,
                session.source.parent / f"report-{language}",
                language=language,
            )
            report_path = report_paths["html"]
            session.reports[language] = report_path
        response = FileResponse(
            report_path,
            media_type="text/html",
            filename="workbooklens-report.html",
            content_disposition_type="inline",
        )
        response.set_cookie(
            LANGUAGE_COOKIE_NAME,
            language,
            max_age=LANGUAGE_COOKIE_MAX_AGE,
            httponly=True,
            samesite="lax",
            path="/",
        )
        return response

    @app.get("/sessions/{session_id}/plan")
    async def plan(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        return FileResponse(
            session.plan_path, media_type="application/json", filename="repair-plan.json"
        )

    @app.post("/sessions/{session_id}/apply")
    async def apply_selected(
        session_id: str,
        request: Request,
        csrf_token: str = Form(default=""),
        language: str = Form(default=""),
        patch_id: list[str] = Form(default=[]),
        accept_layout_risk: bool = Form(default=False),
    ) -> RedirectResponse:
        _validate_post_request(request, csrf_token)
        session = session_or_404(session_id)
        selected_language = _request_language(request, language, fallback=session.language)
        if not patch_id:
            raise _http_error(
                400,
                "Select at least one reviewed patch",
                "repair.selection_required",
            )
        fixed = session.source.parent / "fixed.xlsx"
        fixed.unlink(missing_ok=True)
        result = await run_in_threadpool(
            apply_patch_plan,
            session.source,
            session.plan,
            fixed,
            selected_ids=set(patch_id),
            accept_layout_risk=accept_layout_risk,
        )
        apply_report = session.source.parent / "apply-report.json"
        write_json(apply_report, result.model_dump(mode="json"))
        semantic_diff = await run_in_threadpool(compare_workbooks, session.source, fixed)
        session.fixed = fixed
        session.result = result
        session.semantic_diff = semantic_diff
        session.diff_reports = {}
        session.apply_report = apply_report
        session.language = selected_language
        response = RedirectResponse(
            f"/sessions/{session_id}/completed",
            status_code=303,
        )
        _set_language_cookie(response, selected_language)
        return response

    @app.get("/sessions/{session_id}/completed", response_class=HTMLResponse)
    async def completed(
        session_id: str,
        request: Request,
        lang: str | None = None,
    ) -> HTMLResponse:
        session = session_or_404(session_id)
        if session.result is None:
            raise _http_error(
                404,
                "No fixed workbook has been created",
                "download.not_ready",
            )
        language = _request_language(request, lang, fallback=session.language)
        session.language = language
        translator = partial(translate, language=language)
        html = _render_page(
            environment,
            "applied.html",
            language,
            page_title=translator("web.applied_title"),
            language_action=f"/sessions/{session_id}/completed",
            view="completed",
            session_id=session_id,
            result=session.result,
        )
        return _language_response(html, language)

    @app.get("/sessions/{session_id}/fixed")
    async def fixed(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        if session.fixed is None:
            raise _http_error(
                404,
                "No fixed workbook has been created",
                "download.not_ready",
            )
        filename = f"{Path(session.filename).stem}.fixed.xlsx"
        return FileResponse(
            session.fixed,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )

    @app.get("/sessions/{session_id}/diff")
    async def diff(
        session_id: str,
        request: Request,
        lang: str | None = None,
    ) -> FileResponse:
        session = session_or_404(session_id)
        if session.semantic_diff is None:
            raise _http_error(
                404,
                "No semantic diff has been created",
                "download.not_ready",
            )
        language = _request_language(request, lang, fallback=session.language)
        session.language = language
        diff_reports = session.diff_reports
        if diff_reports is None:
            diff_reports = {}
            session.diff_reports = diff_reports
        diff_path = diff_reports.get(language)
        if diff_path is None:
            diff_path = session.source.parent / f"diff-{language}.html"
            await run_in_threadpool(
                write_diff_report,
                session.semantic_diff,
                diff_path,
                language=language,
            )
            diff_reports[language] = diff_path
        response = FileResponse(
            diff_path,
            media_type="text/html",
            filename="workbooklens-diff.html",
            content_disposition_type="inline",
        )
        response.set_cookie(
            LANGUAGE_COOKIE_NAME,
            language,
            max_age=LANGUAGE_COOKIE_MAX_AGE,
            httponly=True,
            samesite="lax",
            path="/",
        )
        return response

    @app.get("/sessions/{session_id}/apply-report")
    async def apply_report(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        if session.apply_report is None:
            raise _http_error(
                404,
                "No apply report has been created",
                "download.not_ready",
            )
        return FileResponse(
            session.apply_report,
            media_type="application/json",
            filename="apply-report.json",
        )

    return app
