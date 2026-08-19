from __future__ import annotations

import json
from pathlib import Path

import yaml
from openpyxl import Workbook
from typer.testing import CliRunner

from workbooklens.cli import app

runner = CliRunner()


def _broken_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["A1"] = "=#REF!+1"
    workbook.save(path)
    workbook.close()


def test_scan_baseline_new_only_and_suppression_config(tmp_path: Path) -> None:
    workbook = tmp_path / "broken.xlsx"
    _broken_workbook(workbook)
    first_report = tmp_path / "first"
    first = runner.invoke(app, ["scan", str(workbook), "--out", str(first_report)])
    assert first.exit_code == 0, first.stdout

    second_report = tmp_path / "second"
    second = runner.invoke(
        app,
        [
            "scan",
            str(workbook),
            "--out",
            str(second_report),
            "--baseline",
            str(first_report / "findings.json"),
            "--new-only",
            "--fail-on",
            "error",
        ],
    )
    assert second.exit_code == 0, second.stdout
    payload = json.loads((second_report / "findings.json").read_text(encoding="utf-8"))
    assert payload["schema_version"] == 2
    assert payload["source_scope"].startswith("external/")
    assert payload["source_scope"].endswith("/broken.xlsx")
    assert payload["findings"] == []
    assert payload["summary"]["baseline_known"] >= 1
    second_html = (second_report / "report.html").read_text(encoding="utf-8")
    assert "No active findings remain after baseline and suppression policy" in second_html
    assert "No findings were raised by the enabled deterministic rules" not in second_html

    chained_report = tmp_path / "chained"
    chained = runner.invoke(
        app,
        [
            "scan",
            str(workbook),
            "--out",
            str(chained_report),
            "--baseline",
            str(second_report / "findings.json"),
            "--new-only",
            "--fail-on",
            "error",
        ],
    )
    assert chained.exit_code == 0, chained.stdout
    chained_payload = json.loads((chained_report / "findings.json").read_text(encoding="utf-8"))
    assert chained_payload["findings"] == []
    assert chained_payload["summary"]["baseline_known"] >= 1

    wrong_scope = runner.invoke(
        app,
        [
            "scan",
            str(workbook),
            "--out",
            str(tmp_path / "wrong-scope"),
            "--baseline",
            str(first_report / "findings.json"),
            "--source-scope",
            "archive/broken.xlsx",
            "--new-only",
        ],
    )
    assert wrong_scope.exit_code == 2
    assert "does not match" in wrong_scope.stdout

    config = tmp_path / "workbooklens.yml"
    config.write_text(
        yaml.safe_dump(
            {
                "version": 2,
                "suppressions": [
                    {
                        "id": "accepted-broken-reference",
                        "reason": "Synthetic regression fixture",
                        "rules": ["WL001_BROKEN_REFERENCE"],
                        "sheets": ["Data"],
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    suppressed_report = tmp_path / "suppressed"
    suppressed = runner.invoke(
        app,
        [
            "scan",
            str(workbook),
            "--out",
            str(suppressed_report),
            "--config",
            str(config),
            "--fail-on",
            "error",
        ],
    )
    assert suppressed.exit_code == 0, suppressed.stdout
    suppressed_payload = json.loads(
        (suppressed_report / "findings.json").read_text(encoding="utf-8")
    )
    assert suppressed_payload["summary"]["suppressed"] >= 1
    assert all(
        finding["rule_id"] != "WL001_BROKEN_REFERENCE" for finding in suppressed_payload["findings"]
    )


def test_new_only_requires_baseline(tmp_path: Path) -> None:
    workbook = tmp_path / "broken.xlsx"
    _broken_workbook(workbook)
    result = runner.invoke(
        app,
        ["scan", str(workbook), "--out", str(tmp_path / "report"), "--new-only"],
    )
    assert result.exit_code == 2
    assert "--new-only requires --baseline" in result.stdout
