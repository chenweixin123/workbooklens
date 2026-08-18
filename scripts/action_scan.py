"""Scan changed .xlsx files for the WorkbookLens composite GitHub Action."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path


def _git(*arguments: str) -> list[str]:
    git = shutil.which("git")
    if git is None:
        raise RuntimeError("git is required for changed-workbook detection")
    workspace = Path(os.environ.get("GITHUB_WORKSPACE", ".")).resolve()
    result = subprocess.run(  # noqa: S603 - absolute executable; arguments are a fixed internal set
        [git, *arguments],
        check=True,
        capture_output=True,
        cwd=workspace,
        text=True,
        encoding="utf-8",
    )
    separator = "\0" if "-z" in arguments else "\n"
    return [item for item in result.stdout.split(separator) if item]


def _revision(value: str, label: str) -> str:
    if not value or value.startswith("-") or any(character in value for character in "\0\r\n"):
        raise RuntimeError(f"{label} is not a safe Git revision")
    return value


def _changed_files(base: str, head: str) -> list[str]:
    if base and set(base) != {"0"}:
        try:
            return _git(
                "diff",
                "--name-only",
                "-z",
                "--diff-filter=ACMR",
                _revision(base, "base-sha"),
                _revision(head, "head-sha"),
                "--",
            )
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(
                "Unable to diff the base commit. Check out enough Git history (fetch-depth: 0)."
            ) from exc
    return _git("ls-files", "-z")


def _write_output(name: str, value: str) -> None:
    output_file = os.environ.get("GITHUB_OUTPUT")
    if output_file:
        with Path(output_file).open("a", encoding="utf-8") as handle:
            handle.write(f"{name}={value}\n")


def main() -> int:
    workspace = Path(os.environ.get("GITHUB_WORKSPACE", ".")).resolve()
    scan_root = (workspace / os.environ.get("WORKBOOKLENS_SCAN_PATH", ".")).resolve()
    if scan_root != workspace and workspace not in scan_root.parents:
        raise RuntimeError("WorkbookLens scan path must stay inside GITHUB_WORKSPACE")
    output_root = (
        workspace / os.environ.get("WORKBOOKLENS_OUTPUT", ".workbooklens-reports")
    ).resolve()
    if output_root != workspace and workspace not in output_root.parents:
        raise RuntimeError("WorkbookLens output path must stay inside GITHUB_WORKSPACE")
    output_root.mkdir(parents=True, exist_ok=True)
    base = os.environ.get("WORKBOOKLENS_BASE_SHA", "")
    head = os.environ.get("WORKBOOKLENS_HEAD_SHA") or "HEAD"
    fail_on = os.environ.get("WORKBOOKLENS_FAIL_ON", "error")
    changed = _changed_files(base, head)
    selected: list[Path] = []
    for relative in changed:
        candidate = (workspace / relative).resolve()
        if candidate.suffix.lower() != ".xlsx" or not candidate.is_file():
            continue
        if candidate != scan_root and scan_root not in candidate.parents:
            continue
        selected.append(candidate)
    results: list[dict[str, object]] = []
    overall_exit = 0
    for workbook in sorted(selected):
        relative = workbook.relative_to(workspace)
        report_directory = output_root / relative.with_suffix("")
        command = [
            sys.executable,
            "-m",
            "workbooklens",
            "scan",
            str(workbook),
            "--out",
            str(report_directory),
            "--fail-on",
            fail_on,
        ]
        completed = subprocess.run(  # noqa: S603 - sys.executable and argument list, never a shell
            command, check=False
        )
        results.append(
            {
                "workbook": relative.as_posix(),
                "report": report_directory.relative_to(workspace).as_posix(),
                "exit_code": completed.returncode,
            }
        )
        if completed.returncode not in {0, 1}:
            overall_exit = completed.returncode
        elif completed.returncode == 1 and overall_exit == 0:
            overall_exit = 1
    manifest = {
        "base_sha": base or None,
        "head_sha": head,
        "file_count": len(selected),
        "results": results,
    }
    (output_root / "manifest.json").write_text(
        json.dumps(manifest, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    _write_output("report-path", str(output_root))
    _write_output("file-count", str(len(selected)))
    _write_output("exit-code", str(overall_exit))
    if not selected:
        print("WorkbookLens: no changed .xlsx files to scan")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
