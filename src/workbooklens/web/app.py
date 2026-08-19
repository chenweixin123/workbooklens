"""Accessible server-rendered local workflow with bounded untrusted uploads."""

from __future__ import annotations

import secrets
import shutil
import tempfile
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from pathlib import Path

from fastapi import FastAPI, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from jinja2 import Environment, select_autoescape
from starlette.concurrency import run_in_threadpool

from workbooklens.diff import compare_workbooks, write_diff_report
from workbooklens.exceptions import WorkbookLensError
from workbooklens.models import Finding, PatchPlan
from workbooklens.ooxml.safety import PackageLimits
from workbooklens.repair import apply_patch_plan, build_patch_plan
from workbooklens.repair.planning import write_patch_plan
from workbooklens.reports import write_scan_report
from workbooklens.scanner import scan_workbook
from workbooklens.utils import write_json

INDEX_TEMPLATE = """<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>WorkbookLens</title>
<style>:root{color-scheme:light dark;--bg:#f5f7fb;--card:#fff;--ink:#172033;--line:#dce3ee;--accent:#3157d5}
@media(prefers-color-scheme:dark){:root{--bg:#10131a;--card:#181d27;--ink:#edf2f7;--line:#303849;--accent:#8da2ff}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.55 Arial,sans-serif}main{max-width:920px;margin:auto;padding:32px}.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px;margin:18px 0}button{font:inherit;background:var(--accent);color:white;border:0;border-radius:8px;padding:10px 16px;font-weight:700}input[type=file]{display:block;margin:16px 0}.privacy{border-left:5px solid var(--accent)}</style></head>
<body><main><h1>WorkbookLens</h1><p>Lint, review, and safely repair an Excel workbook on this computer.</p>
<section class="card privacy"><b>Private by design.</b> Files remain in a temporary local directory and are removed when this server stops. Formulas, macros, external links, and embedded objects are never executed.</section>
<form id="scan-form" class="card" action="/scan" method="post" enctype="multipart/form-data"><h2>1. Choose a workbook</h2>
<label for="workbook">Supported: .xlsx; .xlsm is scan-only</label><input id="workbook" name="workbook" type="file" accept=".xlsx,.xlsm" required>
<button id="scan-button" type="submit">Scan locally</button><p id="scan-status" aria-live="polite">The upload limit is {{ max_mb }} MB.</p></form></main>
<script>document.querySelector('#scan-form').addEventListener('submit',()=>{const button=document.querySelector('#scan-button');button.disabled=true;button.textContent='Scanning…';document.querySelector('#scan-status').textContent='Validating the package and running deterministic rules locally…'});</script></body></html>"""

RESULT_TEMPLATE = """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>WorkbookLens findings</title>
<style>:root{color-scheme:light dark;--bg:#f5f7fb;--card:#fff;--ink:#172033;--line:#dce3ee;--accent:#3157d5;--error:#c24135;--warning:#a15c00;--info:#1769aa}
@media(prefers-color-scheme:dark){:root{--bg:#10131a;--card:#181d27;--ink:#edf2f7;--line:#303849;--accent:#8da2ff}}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 Arial,sans-serif}main{max-width:1100px;margin:auto;padding:28px}.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin:16px 0}.finding{border-left:5px solid var(--info)}.finding.error,.finding.critical{border-left-color:var(--error)}.finding.warning{border-left-color:var(--warning)}button,.button{display:inline-block;font:inherit;background:var(--accent);color:white;border:0;border-radius:8px;padding:9px 14px;text-decoration:none;font-weight:700}code{overflow-wrap:anywhere}.actions{display:flex;gap:10px;flex-wrap:wrap}label.patch{display:block;padding:10px;border:1px solid var(--line);border-radius:8px;margin:8px 0}</style></head>
<body><main><a href="/">← New scan</a><h1>Findings for {{ filename }}</h1><div class="actions"><a class="button" href="/sessions/{{ session_id }}/report">Open full HTML report</a><a class="button" href="/sessions/{{ session_id }}/plan">Download JSON plan</a></div>
<form class="card" action="/sessions/{{ session_id }}/apply" method="post"><h2>2. Review safe patches</h2>
{% if patches %}<p>Select only changes you have reviewed. The original file is never modified.</p>{% for patch in patches %}<label class="patch"><input type="checkbox" name="patch_id" value="{{ patch.id }}"> <b>{{ patch.sheet }}!{{ patch.cell }}</b> · {{ patch.kind.value }} · {{ '%.0f'|format(patch.confidence.root*100) }}% confidence<br><code>{{ patch.before }}</code> → <code>{{ patch.after }}</code><br>{{ patch.description }}</label>{% endfor %}<button type="submit">Apply selected patches to a copy</button>{% else %}<p>No safe deterministic patches are available. Scan and report downloads remain available.</p>{% endif %}</form>
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


def _environment() -> Environment:
    return Environment(autoescape=select_autoescape(default=True))


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
    async def index() -> str:
        return environment.from_string(INDEX_TEMPLATE).render(max_mb=max_file_bytes // 1024**2)

    @app.post("/scan", response_class=HTMLResponse)
    async def scan_upload(request: Request, workbook: UploadFile) -> str:
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
        safe_patches = [] if suffix == ".xlsm" else [patch for patch in plan.patches if patch.safe]
        return environment.from_string(RESULT_TEMPLATE).render(
            session_id=session_id,
            filename=filename,
            findings=scan.findings,
            patches=safe_patches,
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
        patch_id: list[str] = Form(default=[]),
    ) -> str:
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
