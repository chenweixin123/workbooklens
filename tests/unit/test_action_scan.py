from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from scripts import action_scan


def test_revision_and_boolean_validation() -> None:
    assert action_scan._revision("abc123", "base") == "abc123"
    assert action_scan._boolean("TRUE", "new-only") is True
    assert action_scan._boolean("false", "new-only") is False
    with pytest.raises(RuntimeError, match="safe Git revision"):
        action_scan._revision("--output=/tmp/x", "base")
    with pytest.raises(RuntimeError, match="true or false"):
        action_scan._boolean("sometimes", "new-only")


def test_scan_command_forwards_config_baseline_and_new_only(tmp_path: Path) -> None:
    command = action_scan._command(
        tmp_path / "book.xlsm",
        tmp_path / "report",
        mode="scan",
        config=tmp_path / "workbooklens.yml",
        baseline=tmp_path / "baseline.json",
        new_only=True,
        fail_on="warning",
        source_scope="books/book.xlsm",
    )
    assert command[3:5] == ["scan", str(tmp_path / "book.xlsm")]
    assert "--config" in command
    assert "--baseline" in command
    assert "--new-only" in command
    assert command[command.index("--source-scope") + 1] == "books/book.xlsm"


def test_action_main_processes_xlsx_and_xlsm(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    (tmp_path / "books").mkdir()
    for name in ("a.xlsx", "b.xlsm", "ignored.txt"):
        (tmp_path / "books" / name).write_bytes(b"fixture")
    config = tmp_path / "workbooklens.yml"
    baseline = tmp_path / "baseline.json"
    config.write_text("version: 1\n", encoding="utf-8")
    baseline.write_text('{"finding_ids": []}\n', encoding="utf-8")
    github_output = tmp_path / "github-output.txt"
    monkeypatch.setenv("GITHUB_WORKSPACE", str(tmp_path))
    monkeypatch.setenv("GITHUB_OUTPUT", str(github_output))
    monkeypatch.setenv("WORKBOOKLENS_SCAN_PATH", "books")
    monkeypatch.setenv("WORKBOOKLENS_OUTPUT", "reports")
    monkeypatch.setenv("WORKBOOKLENS_CONFIG", "workbooklens.yml")
    monkeypatch.setenv("WORKBOOKLENS_BASELINE", "baseline.json")
    monkeypatch.setenv("WORKBOOKLENS_NEW_ONLY", "true")
    monkeypatch.setattr(
        action_scan,
        "_changed_files",
        lambda base, head: ["books/a.xlsx", "books/b.xlsm", "books/ignored.txt"],
    )
    commands: list[list[str]] = []

    def fake_run(command: list[str], check: bool) -> subprocess.CompletedProcess[str]:
        commands.append(command)
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(action_scan.subprocess, "run", fake_run)
    assert action_scan.main() == 0
    assert len(commands) == 2
    assert all("--baseline" in command and "--new-only" in command for command in commands)
    assert {command[command.index("--source-scope") + 1] for command in commands} == {
        "books/a.xlsx",
        "books/b.xlsm",
    }
    manifest = json.loads((tmp_path / "reports" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["schema_version"] == 2
    assert manifest["file_count"] == 2
    assert {item["workbook"] for item in manifest["results"]} == {
        "books/a.xlsx",
        "books/b.xlsm",
    }
    assert "file-count=2" in github_output.read_text(encoding="utf-8")


@pytest.mark.parametrize(
    ("scan_path", "expected_workbooks"),
    [
        (".", {"books/a.xlsx", "books/b.xlsm"}),
        ("books", {"books/a.xlsx", "books/b.xlsm"}),
        ("books/a.xlsx", {"books/a.xlsx"}),
    ],
)
def test_action_scan_path_accepts_default_directory_and_file(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    scan_path: str,
    expected_workbooks: set[str],
) -> None:
    books = tmp_path / "books"
    books.mkdir()
    for name in ("a.xlsx", "b.xlsm"):
        (books / name).write_bytes(b"fixture")
    monkeypatch.setenv("GITHUB_WORKSPACE", str(tmp_path))
    monkeypatch.setenv("WORKBOOKLENS_MODE", "scan")
    monkeypatch.setenv("WORKBOOKLENS_SCAN_PATH", scan_path)
    monkeypatch.setenv("WORKBOOKLENS_OUTPUT", "reports")
    for variable in (
        "WORKBOOKLENS_CONFIG",
        "WORKBOOKLENS_BASELINE",
        "WORKBOOKLENS_NEW_ONLY",
    ):
        monkeypatch.delenv(variable, raising=False)
    monkeypatch.setattr(
        action_scan,
        "_changed_files",
        lambda base, head: ["books/a.xlsx", "books/b.xlsm"],
    )
    commands: list[list[str]] = []

    def fake_run(command: list[str], check: bool) -> subprocess.CompletedProcess[str]:
        commands.append(command)
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(action_scan.subprocess, "run", fake_run)
    assert action_scan.main() == 0
    assert {
        Path(command[4]).relative_to(tmp_path).as_posix() for command in commands
    } == expected_workbooks


def test_action_missing_scan_path_exits_nonzero(tmp_path: Path) -> None:
    environment = os.environ.copy()
    for variable in tuple(environment):
        if variable.startswith("WORKBOOKLENS_") or variable == "GITHUB_OUTPUT":
            environment.pop(variable)
    environment["GITHUB_WORKSPACE"] = str(tmp_path)
    environment["WORKBOOKLENS_SCAN_PATH"] = "missing-workbooks"
    environment["WORKBOOKLENS_OUTPUT"] = "reports"

    completed = subprocess.run(  # noqa: S603
        [sys.executable, str(Path(action_scan.__file__).resolve())],
        cwd=tmp_path,
        env=environment,
        check=False,
        capture_output=True,
        text=True,
    )

    assert completed.returncode != 0
    assert "path does not exist: missing-workbooks" in completed.stderr
    assert not (tmp_path / "reports").exists()


def test_test_mode_requires_config_and_rejects_baseline(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("GITHUB_WORKSPACE", str(tmp_path))
    monkeypatch.setenv("WORKBOOKLENS_MODE", "test")
    with pytest.raises(RuntimeError, match="requires the config"):
        action_scan.main()
    config = tmp_path / "workbooklens.yml"
    baseline = tmp_path / "baseline.json"
    config.write_text("version: 1\n", encoding="utf-8")
    baseline.write_text("{}\n", encoding="utf-8")
    monkeypatch.setenv("WORKBOOKLENS_CONFIG", config.name)
    monkeypatch.setenv("WORKBOOKLENS_BASELINE", baseline.name)
    with pytest.raises(RuntimeError, match="only in scan mode"):
        action_scan.main()
