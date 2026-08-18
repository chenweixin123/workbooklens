from __future__ import annotations

import re
from pathlib import Path

from fastapi.testclient import TestClient

from workbooklens.demo.workflow import generate_demo_workbook
from workbooklens.web import create_app


def test_local_web_scan_apply_and_download_workflow(tmp_path: Path) -> None:
    workbook = tmp_path / "demo.xlsx"
    generate_demo_workbook(workbook)
    app = create_app(max_file_bytes=5 * 1024 * 1024)
    with TestClient(app) as client:
        assert client.get("/health").json() == {"status": "ok"}
        home = client.get("/")
        assert home.status_code == 200
        assert "Files remain in a temporary local directory" in home.text
        with workbook.open("rb") as handle:
            response = client.post(
                "/scan",
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
        patch_ids = re.findall(r'name="patch_id" value="([^"]+)"', response.text)
        assert len(patch_ids) == 5
        assert client.get(f"/sessions/{session_id}/report").status_code == 200
        assert client.get(f"/sessions/{session_id}/plan").status_code == 200
        applied = client.post(
            f"/sessions/{session_id}/apply",
            data={"patch_id": patch_ids},
        )
        assert applied.status_code == 200, applied.text
        assert "Validated copy created" in applied.text
        fixed = client.get(f"/sessions/{session_id}/fixed")
        assert fixed.status_code == 200
        assert fixed.content.startswith(b"PK")
        assert client.get(f"/sessions/{session_id}/diff").status_code == 200
        assert client.get(f"/sessions/{session_id}/apply-report").status_code == 200


def test_web_rejects_wrong_extension_and_oversized_upload() -> None:
    app = create_app(max_file_bytes=10)
    with TestClient(app) as client:
        wrong = client.post(
            "/scan",
            files={"workbook": ("notes.txt", b"hello", "text/plain")},
        )
        assert wrong.status_code == 400
        oversized = client.post(
            "/scan",
            files={"workbook": ("book.xlsx", b"x" * 11, "application/octet-stream")},
        )
        assert oversized.status_code == 413


def test_xlsm_web_scan_does_not_offer_repairs(tmp_path: Path) -> None:
    xlsx = tmp_path / "demo.xlsx"
    generate_demo_workbook(xlsx)
    app = create_app(max_file_bytes=5 * 1024 * 1024)
    with TestClient(app) as client, xlsx.open("rb") as handle:
        response = client.post(
            "/scan",
            files={
                "workbook": (
                    "demo.xlsm",
                    handle,
                    "application/vnd.ms-excel.sheet.macroEnabled.12",
                )
            },
        )
    assert response.status_code == 200
    assert 'name="patch_id"' not in response.text
    assert "No safe deterministic patches are available" in response.text


def test_serve_command_hardcodes_loopback_binding() -> None:
    cli_path = Path(__file__).parents[2] / "src" / "workbooklens" / "cli.py"
    source = cli_path.read_text(encoding="utf-8")
    assert 'host="127.0.0.1"' in source
