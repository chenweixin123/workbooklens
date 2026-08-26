from __future__ import annotations

import asyncio
import re
from pathlib import Path
from types import SimpleNamespace

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from httpx import Response
from starlette import formparsers
from starlette.types import Message, Scope

import workbooklens.conversion as conversion
import workbooklens.web.app as web_app_module
from workbooklens import __version__
from workbooklens.conversion import ConversionProvider, ConversionResult
from workbooklens.demo.workflow import generate_demo_workbook
from workbooklens.exceptions import UsageError
from workbooklens.i18n import require_translation
from workbooklens.models import (
    PatchDerivation,
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchResult,
    PatchRisk,
    RecalculationProvider,
    ValidationStatus,
    WorkbookDiff,
)
from workbooklens.scanner import scan_workbook
from workbooklens.web import create_app
from workbooklens.web.app import (
    CSRF_COOKIE_NAME,
    LANGUAGE_COOKIE_NAME,
    MAX_MULTIPART_OVERHEAD_BYTES,
    MAX_PROFILE_BYTES,
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
        all_patch_ids = re.findall(
            r'name="patch_id" value="([^"]+)" data-risk="[^"]+"', response.text
        )
        assert len(patch_ids) >= 2
        assert len(all_patch_ids) > len(patch_ids)
        assert 'data-risk="formula_derived"' in response.text
        assert 'data-risk="layout_review"' in response.text
        assert re.search(
            r'<button id="select-all-patches" class="secondary" type="button" '
            r'aria-controls="patch-list">Select safe \+ formulas</button>',
            response.text,
        )
        assert re.search(
            r'<button id="clear-all-patches" class="secondary" type="button" '
            r'aria-controls="patch-list" disabled>Clear all</button>',
            response.text,
        )
        assert f"0 of {len(all_patch_ids)} repairs selected" in response.text
        assert (
            "applyForm.querySelectorAll('#patch-list input[name=\"patch_id\"]:not(:disabled)')"
            in response.text
        )
        assert "checkbox.dataset.autoSelectable === 'true'" in response.text
        assert "autoSelectableCheckboxes.forEach" in response.text
        assert "document.activeElement === selectAllButton" in response.text
        assert "clearAllButton.focus()" in response.text
        assert "document.activeElement === clearAllButton" in response.text
        assert "selectAllButton.focus()" in response.text
        assert "querySelectorAll('input[type=\"checkbox\"]')" not in response.text
        assert 'name="accept_layout_risk"' in response.text
        report = client.get(f"/sessions/{session_id}/report")
        assert report.status_code == 200
        assert report.headers["content-disposition"].startswith("inline")
        _assert_security_headers(report)
        assert client.get(f"/sessions/{session_id}/plan").status_code == 200
        rejected_apply = client.post(f"/sessions/{session_id}/apply", data={"patch_id": patch_ids})
        assert rejected_apply.status_code == 403
        rejected_layout_apply = client.post(
            f"/sessions/{session_id}/apply",
            data={"csrf_token": result_token, "patch_id": all_patch_ids},
        )
        assert rejected_layout_apply.status_code == 400
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


def _v3_web_plan() -> PatchPlan:
    return PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="abc123",
        patches=[
            PatchOperation(
                id="safe-patch",
                kind=PatchKind.NORMALIZE_TEXT,
                sheet="Sheet1",
                cell="A1",
                before="1,000 ",
                after=1000,
                confidence=0.99,
                safe=True,
                description="Normalize a uniquely parsed numeric value.",
                precondition=PatchPrecondition(cell_fingerprint="safe"),
            ),
            PatchOperation(
                id="formula-patch",
                kind=PatchKind.SET_FORMULA,
                sheet="Sheet1",
                cell="B1",
                before="=A1+A2",
                after="=SUM(A1:A2)",
                confidence=0.99,
                safe=False,
                risk=PatchRisk.FORMULA_DERIVED,
                description="Restore the unique peer formula.",
                precondition=PatchPrecondition(cell_fingerprint="formula"),
                derivation=PatchDerivation(
                    strategy="r1c1_peer_consensus",
                    sources=["peer_before:Sheet1!A1", "peer_after:Sheet1!A3"],
                    candidate_count=1,
                    invariants=["same_table_region"],
                    requires_recalculation=True,
                ),
            ),
            PatchOperation(
                id="semantic-patch",
                kind=PatchKind.SET_NUMERIC,
                sheet="Sheet1",
                cell="C1",
                before="八万九千",
                after=89000,
                confidence=0.99,
                safe=False,
                risk=PatchRisk.SEMANTIC_REVIEW,
                description="Interpret a business value after review.",
                precondition=PatchPrecondition(cell_fingerprint="semantic"),
                derivation=PatchDerivation(
                    strategy="profile_numeric_semantic_candidate",
                    sources=[
                        "profile_column_role:Sheet1!C:number",
                        "native_peer_consensus:27/28",
                    ],
                    candidate_count=1,
                    invariants=["same_table_region", "unique_numeric_parse"],
                    requires_recalculation=True,
                ),
            ),
            PatchOperation(
                id="layout-patch",
                kind=PatchKind.SET_COLUMN_WIDTH,
                sheet="Sheet1",
                cell="D1",
                before=8.43,
                after=16.0,
                confidence=0.99,
                safe=False,
                risk=PatchRisk.LAYOUT_REVIEW,
                description="Widen the reviewed column.",
                precondition=PatchPrecondition(cell_fingerprint="layout"),
            ),
        ],
    )


def test_web_v3_groups_auto_repair_and_completion_details(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan = _v3_web_plan()
    scan_captured: dict[str, object] = {}
    profile_path_checks: list[tuple[str, bool]] = []
    real_load_test_config = web_app_module.load_test_config

    def fake_scan(*_args: object, **kwargs: object) -> SimpleNamespace:
        scan_captured.update(kwargs)
        return SimpleNamespace(findings=[])

    def capture_profile_path(path: Path):
        profile_path_checks.append(
            (
                path.name,
                path.resolve().is_relative_to(Path(app.state.root).resolve()),
            )
        )
        return real_load_test_config(path)

    monkeypatch.setattr(web_app_module, "scan_workbook", fake_scan)
    monkeypatch.setattr(web_app_module, "build_patch_plan", lambda _scan: plan)
    captured: dict[str, object] = {}

    def fake_apply(*args: object, **kwargs: object) -> PatchResult:
        captured.update(kwargs)
        output = args[2]
        assert isinstance(output, Path)
        output.write_bytes(b"PK local repaired workbook")
        return PatchResult(
            source_sha256="sourcehash",
            output_sha256="outputhash",
            output_path=str(output),
            applied_patch_ids=["safe-patch"],
            package_changes=[],
            resolved_finding_ids=["finding-1"],
            recalculation_provider=RecalculationProvider.EXCEL,
            formula_errors_before=["Sheet1!B1=#VALUE!"],
            formula_errors_after=[],
            downgraded_patch_ids=["formula-patch"],
            skipped_patch_ids=["semantic-patch"],
            validation_status=ValidationStatus.PASSED,
            rollback_performed=False,
        )

    monkeypatch.setattr(web_app_module, "apply_patch_plan", fake_apply)
    monkeypatch.setattr(
        web_app_module,
        "compare_workbooks",
        lambda *_args, **_kwargs: WorkbookDiff(
            before_sha256="sourcehash",
            after_sha256="outputhash",
        ),
    )
    app = create_app(max_file_bytes=1024)
    monkeypatch.setattr(web_app_module, "load_test_config", capture_profile_path)
    with _client(app) as client:
        home = client.get("/?lang=en")
        response = client.post(
            "/scan",
            data={"csrf_token": _csrf_token(home.text), "language": "en"},
            files={
                "workbook": (
                    "v3.xlsx",
                    b"PK local test",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                "profile": (
                    "../../evil.yaml",
                    b"version: 3\nprofile:\n  infer_semantics: true\n  sheets: []\n",
                    "application/yaml",
                ),
            },
        )
        assert response.status_code == 200, response.text
        assert "Lossless normalization and safe repairs" in response.text
        assert "Formula-derived repairs" in response.text
        assert "Semantic confirmation required" in response.text
        assert "Layout confirmation required" in response.text
        assert 'data-risk="safe" data-auto-selectable="true"' in response.text
        assert 'data-risk="formula_derived" data-auto-selectable="true"' in response.text
        assert 'data-risk="semantic_review" data-auto-selectable="false"' in response.text
        assert 'data-risk="layout_review" data-auto-selectable="false"' in response.text
        assert "Derivation evidence" in response.text
        assert "profile_numeric_semantic_candidate" in response.text
        assert "native_peer_consensus:27/28" in response.text
        assert "same_table_region" in response.text
        assert "Requires recalculation" in response.text
        assert 'name="accept_semantic_risk"' in response.text
        assert 'name="accept_layout_risk"' in response.text
        assert 'name="trust_workbook_for_recalculation"' in response.text
        assert "I trust this workbook for isolated local recalculation" in response.text
        assert 'id="repair-mode" type="hidden" name="auto_repair" value="false"' in response.text
        assert 'id="auto-repair-button" type="submit">One-click safe repair' in response.text
        assert "button.id === 'auto-repair-button' ? 'true' : 'false'" in response.text

        session_match = re.search(r"/sessions/([^/]+)/apply", response.text)
        assert session_match
        session_id = session_match.group(1)
        applied = client.post(
            f"/sessions/{session_id}/apply",
            data={
                "csrf_token": _csrf_token(response.text),
                "auto_repair": "true",
                "trust_workbook_for_recalculation": "true",
                "patch_id": [patch.id for patch in plan.patches],
                "accept_semantic_risk": "true",
                "accept_layout_risk": "true",
            },
        )

    assert applied.status_code == 200, applied.text
    assert profile_path_checks == [("profile.yml", True)]
    assert scan_captured["config"] == {
        "version": 3,
        "workbook": {"max_critical_findings": None, "max_error_findings": None},
        "assertions": [],
        "keys": [],
        "profile": {
            "infer_semantics": True,
            "report_trailing_whitespace": True,
            "review_trailing_whitespace_patches": False,
            "sheets": [],
        },
        "suppressions": [],
    }
    assert captured["selected_ids"] == {"semantic-patch"}
    assert captured["auto_repair"] is True
    assert captured["recalc_provider"] == "auto"
    assert captured["trust_workbook_for_recalculation"] is True
    assert captured["accept_semantic_risk"] is True
    assert captured["accept_layout_risk"] is False
    assert captured["config"] == scan_captured["config"]
    assert "Recalculation provider" in applied.text
    assert "Microsoft Excel" in applied.text
    assert "Formula errors before" in applied.text
    assert "formula-patch" in applied.text
    assert "semantic-patch" in applied.text
    assert "Sheet1!A1" in applied.text


@pytest.mark.parametrize(
    ("filename", "payload", "expected_status"),
    [
        ("profile.json", b"version: 3\n", 400),
        ("profile.yml", b"", 400),
        ("profile.yml", b"\xff\xfe", 400),
        ("profile.yml", b"!!python/object:os.system {}\n", 400),
        ("profile.yml", b"version: 3\nunknown: true\n", 400),
        ("profile.yml", b"x" * (MAX_PROFILE_BYTES + 1), 413),
    ],
    ids=(
        "wrong-extension",
        "empty",
        "invalid-utf8",
        "unsafe-yaml-tag",
        "invalid-schema",
        "too-large",
    ),
)
def test_web_profile_upload_rejects_unsafe_or_invalid_content(
    monkeypatch: pytest.MonkeyPatch,
    filename: str,
    payload: bytes,
    expected_status: int,
) -> None:
    monkeypatch.setattr(
        web_app_module,
        "scan_workbook",
        lambda *_args, **_kwargs: SimpleNamespace(findings=[]),
    )
    monkeypatch.setattr(
        web_app_module,
        "build_patch_plan",
        lambda _scan: PatchPlan(
            tool_version="2.4.0",
            source_name="input.xlsx",
            source_sha256="abc123",
            patches=[],
        ),
    )
    app = create_app(max_file_bytes=2 * 1024 * 1024)
    with _client(app) as client:
        home = client.get("/?lang=en")
        response = client.post(
            "/scan",
            data={"csrf_token": _csrf_token(home.text), "language": "en"},
            files={
                "workbook": (
                    "input.xlsx",
                    b"PK local test",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                "profile": (filename, payload, "application/yaml"),
            },
        )

    assert response.status_code == expected_status


def test_web_repair_failure_shows_validation_state_and_safe_report_download(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    plan = _v3_web_plan()
    monkeypatch.setattr(
        web_app_module,
        "scan_workbook",
        lambda *_args, **_kwargs: SimpleNamespace(findings=[]),
    )
    monkeypatch.setattr(web_app_module, "build_patch_plan", lambda _scan: plan)

    def failed_apply(*args: object, **_kwargs: object) -> PatchResult:
        output = args[2]
        assert isinstance(output, Path)
        report = output.parent / "fixed.workbooklens-failure-test.txt"
        report.write_text("failure report / 失败报告\n", encoding="utf-8")
        result = PatchResult(
            source_sha256="sourcehash",
            output_sha256="",
            output_path=str(output),
            applied_patch_ids=[],
            package_changes=[],
            recalculation_provider=RecalculationProvider.EXCEL,
            formula_errors_before=["Sheet1!B1=#VALUE!"],
            formula_errors_after=["Sheet1!B1=#VALUE!"],
            validation_status=ValidationStatus.FAILED,
            rollback_performed=True,
            failure_report_path=str(report),
        )
        error = UsageError("validation failed")
        error.__dict__["patch_result"] = result
        error.__dict__["failure_report_path"] = report
        raise error

    monkeypatch.setattr(web_app_module, "apply_patch_plan", failed_apply)
    app = create_app(max_file_bytes=1024)
    with _client(app) as client:
        home = client.get("/?lang=zh-CN")
        scanned = client.post(
            "/scan",
            data={"csrf_token": _csrf_token(home.text), "language": "zh-CN"},
            files={
                "workbook": (
                    "failed.xlsx",
                    b"PK local test",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )
        session_match = re.search(r"/sessions/([^/]+)/apply", scanned.text)
        assert session_match
        session_id = session_match.group(1)
        failed = client.post(
            f"/sessions/{session_id}/apply",
            data={
                "csrf_token": _csrf_token(scanned.text),
                "language": "zh-CN",
                "patch_id": "safe-patch",
            },
        )
        downloaded = client.get(f"/sessions/{session_id}/failure-report")

    assert failed.status_code == 400
    assert "修复验证结果" in failed.text
    assert "重算提供者" in failed.text
    assert "Microsoft Excel" in failed.text
    assert "失败并已回滚" in failed.text
    assert "是否执行回滚" in failed.text
    assert "下载中英文失败报告" in failed.text
    assert downloaded.status_code == 200
    assert downloaded.text.splitlines() == ["failure report / 失败报告"]
    assert "attachment" in downloaded.headers["content-disposition"]


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

    assert response.text.count('class="file-picker-control"') == 3
    assert len(re.findall(r'<input[^>]+type="file"', response.text)) == 3
    assert "file-selector-button" not in response.text
    assert "bindFileSelection('workbook'" in response.text
    assert "bindFileSelection('profile'" in response.text
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
        assert ">全选安全项与公式项</button>" in chinese.text
        assert ">取消全选</button>" in chinese.text
        assert "Select safe + formulas</button>" not in chinese.text
        assert "Clear all</button>" not in chinese.text
        assert "Inspection results" not in chinese.text

        persisted = client.get(f"/sessions/{session_id}")
        assert '<html lang="zh-CN">' in persisted.text
        assert "发现的问题" in persisted.text


def test_web_localizes_structured_evidence_without_changing_canonical_scan(
    tmp_path: Path,
    monkeypatch,
) -> None:
    workbook = tmp_path / "demo.xlsx"
    generate_demo_workbook(workbook)
    scan = scan_workbook(workbook)
    assert scan.findings
    canonical_observed = {
        "source_proof": [{"font_size": 8}],
        "unknown_key": "用户原文",
    }
    canonical_details = {
        "proof": "propagated_formula_error",
        "fixed_total_pages": 3,
    }
    finding = scan.findings[0]
    scan.findings[0] = finding.model_copy(
        update={
            "evidence": finding.evidence.model_copy(
                update={
                    "observed": canonical_observed,
                    "details": canonical_details,
                }
            )
        }
    )
    monkeypatch.setattr(
        "workbooklens.web.app.scan_workbook",
        lambda *_args, **_kwargs: scan,
    )

    app = create_app(max_file_bytes=5 * 1024 * 1024)
    with _client(app) as client:
        home = client.get("/?lang=en")
        with workbook.open("rb") as handle:
            english = client.post(
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
        assert english.status_code == 200
        assert "Source proof" in english.text
        assert "Font size" in english.text
        assert "Proof" in english.text
        assert "Propagated formula error" in english.text
        assert "Fixed total pages" in english.text
        session_match = re.search(r"/sessions/([^/]+)/apply", english.text)
        assert session_match

        chinese = client.get(f"/sessions/{session_match.group(1)}?lang=zh-CN")
        assert chinese.status_code == 200
        assert "源证明" in chinese.text
        assert "字号" in chinese.text
        assert "证明" in chinese.text
        assert "传播的公式错误" in chinese.text
        assert "固定总页数" in chinese.text
        assert "用户原文" in chinese.text

    assert scan.findings[0].evidence.observed == canonical_observed
    assert scan.findings[0].evidence.details == canonical_details


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
        assert "Scanning and direct OOXML repair do not execute formulas" in home.text
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
    assert not re.search(r'<input[^>]+name="patch_id"', response.text)
    assert 'id="select-all-patches"' not in response.text
    assert 'id="clear-all-patches"' not in response.text
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
                    b"x" * (1024 + MAX_PROFILE_BYTES + MAX_MULTIPART_OVERHEAD_BYTES + 1024),
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
            (
                b"content-length",
                str(1024 + MAX_PROFILE_BYTES + MAX_MULTIPART_OVERHEAD_BYTES + 1).encode(),
            ),
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
    boundary = b"stream-guard"
    body_limit = 1 + MAX_PROFILE_BYTES + MAX_MULTIPART_OVERHEAD_BYTES
    body = (
        b"--"
        + boundary
        + b'\r\nContent-Disposition: form-data; name="workbook"; filename="book.xlsx"\r\n'
        + b"Content-Type: application/octet-stream\r\n\r\n"
        + b"x" * (body_limit + 512 * 1024)
        + b"\r\n--"
        + boundary
        + b"--\r\n"
    )
    chunks = [body[index : index + 262_144] for index in range(0, len(body), 262_144)]
    messages, receive_calls = _run_asgi_post(
        app,
        [
            (b"host", b"127.0.0.1"),
            (b"content-type", b"multipart/form-data; boundary=stream-guard"),
            (b"content-length", str(body_limit).encode()),
            (b"cookie", f"{CSRF_COOKIE_NAME}={token}".encode()),
        ],
        chunks,
    )

    assert _response_status(messages) == 413
    headers = _response_headers(messages)
    assert headers[b"cache-control"] == b"no-store"
    assert headers[b"x-content-type-options"] == b"nosniff"
    assert receive_calls < len(chunks)
    assert sum(len(chunk) for chunk in chunks[receive_calls:]) > 0


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
