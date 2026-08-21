"""Accessible server-rendered local workflow with bounded untrusted uploads."""

from __future__ import annotations

import secrets
import shutil
import tempfile
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlsplit

from fastapi import FastAPI, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, PlainTextResponse
from jinja2 import Environment, select_autoescape
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import MutableHeaders
from starlette.types import ASGIApp, Message, Receive, Scope, Send

from workbooklens.diff import compare_workbooks, write_diff_report
from workbooklens.exceptions import WorkbookLensError
from workbooklens.models import Finding, PatchPlan, PatchRisk
from workbooklens.ooxml.safety import PackageLimits
from workbooklens.repair import apply_patch_plan, build_patch_plan
from workbooklens.repair.planning import write_patch_plan
from workbooklens.reports import write_scan_report
from workbooklens.scanner import scan_workbook
from workbooklens.utils import write_json

LOCAL_HOSTS = frozenset({"localhost", "127.0.0.1"})
CSRF_COOKIE_NAME = "workbooklens_csrf"
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

INDEX_TEMPLATE = """<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>WorkbookLens</title>
<style>:root{color-scheme:light dark;--bg:#f5f7fb;--card:#fff;--ink:#172033;--line:#dce3ee;--accent:#3157d5}
@media(prefers-color-scheme:dark){:root{--bg:#10131a;--card:#181d27;--ink:#edf2f7;--line:#303849;--accent:#8da2ff}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.55 Arial,sans-serif}main{max-width:920px;margin:auto;padding:32px}.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px;margin:18px 0}button{font:inherit;background:var(--accent);color:white;border:0;border-radius:8px;padding:10px 16px;font-weight:700}input[type=file]{display:block;margin:16px 0}.privacy{border-left:5px solid var(--accent)}</style></head>
<body><main><h1>WorkbookLens</h1><p>Lint, review, and safely repair an Excel workbook on this computer.</p>
<section class="card privacy"><b>Private by design.</b> Files remain in a process-owned temporary local directory and are removed on normal server shutdown. An abrupt crash or power loss can leave temporary files until they are removed manually or by operating-system cleanup. Formulas, macros, external links, and embedded objects are never executed.</section>
<form id="scan-form" class="card" action="/scan" method="post" enctype="multipart/form-data"><input type="hidden" name="csrf_token" value="{{ csrf_token }}"><h2>1. Choose a workbook</h2>
<label for="workbook">Supported: .xlsx; .xlsm is scan-only</label><input id="workbook" name="workbook" type="file" accept=".xlsx,.xlsm" required>
<button id="scan-button" type="submit">Scan locally</button><p id="scan-status" aria-live="polite">The upload limit is {{ max_mb }} MB.</p></form></main>
<script>document.querySelector('#scan-form').addEventListener('submit',()=>{const button=document.querySelector('#scan-button');button.disabled=true;button.textContent='Scanning…';document.querySelector('#scan-status').textContent='Validating the package and running deterministic rules locally…'});</script></body></html>"""

RESULT_TEMPLATE = """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>WorkbookLens findings</title>
<style>:root{color-scheme:light dark;--bg:#f5f7fb;--card:#fff;--ink:#172033;--line:#dce3ee;--accent:#3157d5;--error:#c24135;--warning:#a15c00;--info:#1769aa}
@media(prefers-color-scheme:dark){:root{--bg:#10131a;--card:#181d27;--ink:#edf2f7;--line:#303849;--accent:#8da2ff}}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 Arial,sans-serif}main{max-width:1100px;margin:auto;padding:28px}.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin:16px 0}.finding{border-left:5px solid var(--info)}.finding.error,.finding.critical{border-left-color:var(--error)}.finding.warning{border-left-color:var(--warning)}button,.button{display:inline-block;font:inherit;background:var(--accent);color:white;border:0;border-radius:8px;padding:9px 14px;text-decoration:none;font-weight:700}code{overflow-wrap:anywhere}.actions{display:flex;gap:10px;flex-wrap:wrap}label.patch{display:block;padding:10px;border:1px solid var(--line);border-radius:8px;margin:8px 0}</style></head>
<body><main><a href="/">← New scan</a><h1>Findings for {{ filename }}</h1><div class="actions"><a class="button" href="/sessions/{{ session_id }}/report">Open full HTML report</a><a class="button" href="/sessions/{{ session_id }}/plan">Download JSON plan</a></div>
<form class="card" action="/sessions/{{ session_id }}/apply" method="post"><input type="hidden" name="csrf_token" value="{{ csrf_token }}"><h2>2. Review patches</h2>
{% if patches %}<p>Select only changes you have reviewed. The original file is never modified.</p>{% for patch in patches %}<label class="patch"><input type="checkbox" name="patch_id" value="{{ patch.id }}" data-risk="{{ patch.risk.value }}"> <b>{{ patch.sheet }}!{{ patch.cell }}</b> · {{ patch.kind.value }} · {{ '%.0f'|format(patch.confidence.root*100) }}% confidence · {{ patch.risk.value }}<br><code>{{ patch.before }}</code> → <code>{{ patch.after }}</code><br>{{ patch.description }}</label>{% endfor %}
{% if has_layout_review %}<label class="patch"><input type="checkbox" name="accept_layout_risk" value="true"> <b>I reviewed the layout risk.</b> Row heights, column widths, saved view, wrapping, borders, or print pagination may change. Atomic patch groups are always applied together.</label>{% endif %}
<button type="submit">Apply selected patches to a copy</button>{% else %}<p>No safe deterministic patches are available, and no layout-review patches were offered. Scan and report downloads remain available.</p>{% endif %}</form>
<section><h2>All findings ({{ findings|length }})</h2>{% for finding in findings %}<article class="card finding {{ finding.severity.value }}"><b>{{ finding.severity.value|upper }} · {{ finding.rule_id }} · {{ finding.sheet or 'Workbook' }}{% if finding.location %}!{{ finding.location }}{% endif %}</b><h3>{{ finding.title }}</h3><p>{{ finding.explanation }}</p><details><summary>Evidence</summary><p>{{ finding.evidence.summary }}</p><code>{{ finding.evidence.observed }}</code></details></article>{% endfor %}{% if not findings %}<div class="card">No findings.</div>{% endif %}</section></main></body></html>"""

APPLIED_TEMPLATE = """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>WorkbookLens repair complete</title><style>:root{color-scheme:light dark;--bg:#f5f7fb;--card:#fff;--ink:#172033;--line:#dce3ee;--accent:#3157d5}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.55 Arial,sans-serif}main{max-width:850px;margin:auto;padding:32px}.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;margin:16px 0}.button{display:inline-block;background:var(--accent);color:white;border-radius:8px;padding:10px 15px;text-decoration:none;font-weight:700;margin:5px}</style></head><body><main><h1>Validated copy created</h1><section class="card"><p>Applied {{ result.applied_patch_ids|length }} reviewed patches. The source hash remained {{ result.source_sha256 }}.</p><p>{{ result.resolved_finding_ids|length }} findings resolved; {{ result.new_finding_ids|length }} new informational/warning findings.</p>{% for message in result.validation_messages %}<p>✓ {{ message }}</p>{% endfor %}</section><div><a class="button" href="/sessions/{{ session_id }}/fixed">Download fixed workbook</a><a class="button" href="/sessions/{{ session_id }}/diff">Preview semantic diff</a><a class="button" href="/sessions/{{ session_id }}/apply-report">Download apply report</a></div></main></body></html>"""
ERROR_TEMPLATE = """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>WorkbookLens could not continue</title><style>:root{color-scheme:light dark;--bg:#f5f7fb;--card:#fff;--ink:#172033;--line:#dce3ee;--accent:#3157d5;--error:#c24135}@media(prefers-color-scheme:dark){:root{--bg:#10131a;--card:#181d27;--ink:#edf2f7;--line:#303849;--accent:#8da2ff}}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.55 Arial,sans-serif}main{max-width:760px;margin:auto;padding:32px}.card{background:var(--card);border:1px solid var(--line);border-left:5px solid var(--error);border-radius:12px;padding:20px;margin:16px 0}.button{display:inline-block;background:var(--accent);color:white;border-radius:8px;padding:10px 15px;text-decoration:none;font-weight:700}</style></head><body><main><h1>WorkbookLens could not continue</h1><section class="card"><p>{{ message }}</p><p>The source workbook was not modified.</p></section><a class="button" href="/">Choose another workbook</a></main></body></html>"""


@dataclass(slots=True)
class WebSession:
    """Server-owned paths and immutable scan/plan data for one local upload."""

    session_id: str
    filename: str
    source: Path
    report: Path
    plan_path: Path
    plan: PatchPlan
    findings: list[Finding]
    fixed: Path | None = None
    diff: Path | None = None
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
    return Environment(autoescape=select_autoescape(default=True))


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


def _validate_post_request(request: Request, csrf_token: str) -> None:
    request_origin = _origin(str(request.url))
    for header in ("origin", "referer"):
        value = request.headers.get(header)
        if value is not None and _origin(value) != request_origin:
            raise HTTPException(status_code=403, detail="Cross-origin form submission rejected")
    expected = str(request.app.state.csrf_token)
    if not csrf_token or not secrets.compare_digest(csrf_token, expected):
        raise HTTPException(status_code=403, detail="Invalid or missing form security token")


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
    if path == "/scan":
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

        async def reject(status_code: int, detail: str) -> None:
            response = PlainTextResponse(detail, status_code=status_code)
            await response(scope, receive, secure_send)

        if _parse_local_authority(request.headers.get("host", "")) is None:
            await reject(400, "Invalid Host header")
            return

        if request.method not in UNSAFE_METHODS:
            await self.app(scope, receive, secure_send)
            return

        request_origin = _origin(str(request.url))
        for header in ("origin", "referer"):
            value = request.headers.get(header)
            if value is not None and _origin(value) != request_origin:
                await reject(403, "Cross-origin form submission rejected")
                return

        csrf_cookie = request.cookies.get(CSRF_COOKIE_NAME, "")
        if not csrf_cookie or not secrets.compare_digest(csrf_cookie, self.csrf_token):
            await reject(403, "Invalid or missing form security cookie")
            return

        try:
            declared_length = _content_length(scope)
        except ValueError:
            await reject(400, "Invalid Content-Length header")
            return

        body_limit = _request_body_limit(request.url.path, self.max_file_bytes)
        content_type = request.headers.get("content-type", "").partition(";")[0].strip().lower()
        if (
            request.url.path == "/scan"
            and content_type == "multipart/form-data"
            and declared_length is None
        ):
            await reject(411, "Multipart uploads require a Content-Length header")
            return
        if declared_length is not None and declared_length > body_limit:
            await reject(413, "Request body exceeds the configured upload limit")
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
            await reject(413, "Request body exceeds the configured upload limit")


async def _store_upload(upload: UploadFile, target: Path, maximum: int) -> None:
    size = 0
    with target.open("xb") as handle:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > maximum:
                raise HTTPException(
                    status_code=413, detail="Workbook exceeds the configured upload limit"
                )
            handle.write(chunk)
    if size == 0:
        raise HTTPException(status_code=400, detail="Uploaded workbook is empty")


def create_app(*, max_file_bytes: int = 100 * 1024 * 1024) -> FastAPI:
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
        docs_url=None,
        redoc_url=None,
        lifespan=lifespan,
    )
    environment = _environment()
    app.state.csrf_token = secrets.token_urlsafe(32)
    app.add_middleware(
        _LocalRequestGuard,
        max_file_bytes=max_file_bytes,
        csrf_token=app.state.csrf_token,
    )

    def error_page(message: str, status_code: int) -> HTMLResponse:
        html = environment.from_string(ERROR_TEMPLATE).render(message=message)
        return HTMLResponse(html, status_code=status_code)

    @app.exception_handler(WorkbookLensError)
    async def workbook_error(_request: Request, exc: WorkbookLensError) -> HTMLResponse:
        return error_page(str(exc), 400)

    @app.exception_handler(HTTPException)
    async def http_error(_request: Request, exc: HTTPException) -> HTMLResponse:
        return error_page(str(exc.detail), exc.status_code)

    @app.get("/health")
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    @app.get("/", response_class=HTMLResponse)
    async def index() -> HTMLResponse:
        response = HTMLResponse(
            environment.from_string(INDEX_TEMPLATE).render(
                max_mb=max_file_bytes // 1024**2,
                csrf_token=app.state.csrf_token,
            )
        )
        response.set_cookie(
            CSRF_COOKIE_NAME,
            app.state.csrf_token,
            httponly=True,
            samesite="strict",
            path="/",
        )
        return response

    @app.post("/scan", response_class=HTMLResponse)
    async def scan_upload(
        request: Request,
        workbook: UploadFile,
        csrf_token: str = Form(default=""),
    ) -> str:
        _validate_post_request(request, csrf_token)
        filename = Path(workbook.filename or "workbook.xlsx").name
        suffix = Path(filename).suffix.lower()
        if suffix not in {".xlsx", ".xlsm"}:
            raise HTTPException(status_code=400, detail="Upload must be .xlsx or .xlsm")
        if len(sessions) >= 20:
            raise HTTPException(
                status_code=429, detail="Session limit reached; restart the local server"
            )
        session_id = secrets.token_urlsafe(18)
        session_root = request.app.state.root / session_id
        session_root.mkdir(parents=True)
        source = session_root / f"input{suffix}"
        try:
            await _store_upload(workbook, source, max_file_bytes)
            limits = PackageLimits(max_file_bytes=max_file_bytes)
            scan = await run_in_threadpool(scan_workbook, source, limits=limits)
            report_paths = await run_in_threadpool(write_scan_report, scan, session_root / "report")
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
            report=report_paths["html"],
            plan_path=plan_path,
            plan=plan,
            findings=scan.findings,
        )
        reviewable_patches = (
            []
            if suffix == ".xlsm"
            else [
                patch
                for patch in plan.patches
                if patch.safe_only_eligible or patch.risk == PatchRisk.LAYOUT_REVIEW
            ]
        )
        return environment.from_string(RESULT_TEMPLATE).render(
            session_id=session_id,
            filename=filename,
            findings=scan.findings,
            patches=reviewable_patches,
            csrf_token=app.state.csrf_token,
            has_layout_review=any(
                patch.risk == PatchRisk.LAYOUT_REVIEW for patch in reviewable_patches
            ),
        )

    def session_or_404(session_id: str) -> WebSession:
        session = sessions.get(session_id)
        if session is None:
            raise HTTPException(status_code=404, detail="Local session not found or expired")
        return session

    @app.get("/sessions/{session_id}/report")
    async def report(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        return FileResponse(
            session.report,
            media_type="text/html",
            filename="workbooklens-report.html",
            content_disposition_type="inline",
        )

    @app.get("/sessions/{session_id}/plan")
    async def plan(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        return FileResponse(
            session.plan_path, media_type="application/json", filename="repair-plan.json"
        )

    @app.post("/sessions/{session_id}/apply", response_class=HTMLResponse)
    async def apply_selected(
        session_id: str,
        request: Request,
        csrf_token: str = Form(default=""),
        patch_id: list[str] = Form(default=[]),
        accept_layout_risk: bool = Form(default=False),
    ) -> str:
        _validate_post_request(request, csrf_token)
        session = session_or_404(session_id)
        if not patch_id:
            raise HTTPException(status_code=400, detail="Select at least one reviewed patch")
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
        diff = session.source.parent / "diff.html"
        semantic_diff = await run_in_threadpool(compare_workbooks, session.source, fixed)
        await run_in_threadpool(write_diff_report, semantic_diff, diff)
        session.fixed = fixed
        session.diff = diff
        session.apply_report = apply_report
        return environment.from_string(APPLIED_TEMPLATE).render(
            session_id=session_id,
            result=result,
        )

    @app.get("/sessions/{session_id}/fixed")
    async def fixed(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        if session.fixed is None:
            raise HTTPException(status_code=404, detail="No fixed workbook has been created")
        filename = f"{Path(session.filename).stem}.fixed.xlsx"
        return FileResponse(
            session.fixed,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )

    @app.get("/sessions/{session_id}/diff")
    async def diff(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        if session.diff is None:
            raise HTTPException(status_code=404, detail="No semantic diff has been created")
        return FileResponse(
            session.diff,
            media_type="text/html",
            filename="workbooklens-diff.html",
            content_disposition_type="inline",
        )

    @app.get("/sessions/{session_id}/apply-report")
    async def apply_report(session_id: str) -> FileResponse:
        session = session_or_404(session_id)
        if session.apply_report is None:
            raise HTTPException(status_code=404, detail="No apply report has been created")
        return FileResponse(
            session.apply_report,
            media_type="application/json",
            filename="apply-report.json",
        )

    return app
