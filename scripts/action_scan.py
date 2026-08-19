"""Run WorkbookLens against changed workbooks for the composite GitHub Action."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
MODES = {"scan", "test"}


def _workspace() -> Path:
    return Path(os.environ.get("GITHUB_WORKSPACE", ".")).resolve()


def _git(*arguments: str) -> list[str]:
    git = shutil.which("git")
    if git is None:
        raise RuntimeError("git is required for changed-workbook detection")
    result = subprocess.run(  # noqa: S603 - resolved executable and an argument list, never a shell
        [git, *arguments],
        check=True,
        capture_output=True,
        cwd=_workspace(),
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


def _boolean(value: str, label: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes"}:
        return True
    if normalized in {"0", "false", "no", ""}:
        return False
    raise RuntimeError(f"{label} must be true or false")


def _inside_workspace(
    value: str,
    label: str,
    *,
    required_file: bool = False,
    required_path: bool = False,
) -> Path:
    workspace = _workspace()
    resolved = (workspace / value).resolve()
    if resolved != workspace and workspace not in resolved.parents:
        raise RuntimeError(f"{label} must stay inside GITHUB_WORKSPACE")
    if required_path and not resolved.exists():
        raise RuntimeError(f"{label} does not exist: {value}")
    if required_file and not resolved.is_file():
        raise RuntimeError(f"{label} does not exist or is not a file: {value}")
    return resolved


def _write_output(name: str, value: str) -> None:
    output_file = os.environ.get("GITHUB_OUTPUT")
    if output_file:
        with Path(output_file).open("a", encoding="utf-8") as handle:
            handle.write(f"{name}={value}\n")


def _command(
    workbook: Path,
    report_directory: Path,
    *,
    mode: str,
    config: Path | None,
    baseline: Path | None,
    new_only: bool,
    fail_on: str,
    source_scope: str,
) -> list[str]:
    command = [sys.executable, "-m", "workbooklens", mode, str(workbook)]
    if mode == "test":
        if config is None:
            raise RuntimeError("test mode requires the config input")
        command.extend(
            ["--config", str(config), "--out", str(report_directory / "test-results.json")]
        )
        return command
    command.extend(
        [
            "--out",
            str(report_directory),
            "--fail-on",
            fail_on,
            "--source-scope",
            source_scope,
        ]
    )
    if config is not None:
        command.extend(["--config", str(config)])
    if baseline is not None:
        command.extend(["--baseline", str(baseline)])
    if new_only:
        command.append("--new-only")
    return command


def main() -> int:
    mode = os.environ.get("WORKBOOKLENS_MODE", "scan").strip().lower()
    if mode not in MODES:
        raise RuntimeError("mode must be scan or test")
    workspace = _workspace()
    scan_root = _inside_workspace(
        os.environ.get("WORKBOOKLENS_SCAN_PATH", "."),
        "path",
        required_path=True,
    )
    output_root = _inside_workspace(
        os.environ.get("WORKBOOKLENS_OUTPUT", ".workbooklens-reports"), "output"
    )
    output_root.mkdir(parents=True, exist_ok=True)

    config_value = os.environ.get("WORKBOOKLENS_CONFIG", "").strip()
    baseline_value = os.environ.get("WORKBOOKLENS_BASELINE", "").strip()
    config = _inside_workspace(config_value, "config", required_file=True) if config_value else None
    baseline = (
        _inside_workspace(baseline_value, "baseline", required_file=True)
        if baseline_value
        else None
    )
    new_only = _boolean(os.environ.get("WORKBOOKLENS_NEW_ONLY", "false"), "new-only")
    if mode == "test" and (baseline is not None or new_only):
        raise RuntimeError("baseline and new-only are supported only in scan mode")
    if mode == "test" and config is None:
        raise RuntimeError("test mode requires the config input")
    if mode == "scan" and new_only and baseline is None:
        raise RuntimeError("new-only requires the baseline input")

    base = os.environ.get("WORKBOOKLENS_BASE_SHA", "")
    head = os.environ.get("WORKBOOKLENS_HEAD_SHA") or "HEAD"
    fail_on = os.environ.get("WORKBOOKLENS_FAIL_ON", "error").strip().lower()
    changed = _changed_files(base, head)
    selected: list[Path] = []
    for relative in changed:
        candidate = (workspace / relative).resolve()
        if candidate.suffix.lower() not in SUPPORTED_SUFFIXES or not candidate.is_file():
            continue
        if candidate != scan_root and scan_root not in candidate.parents:
            continue
        selected.append(candidate)

    results: list[dict[str, object]] = []
    overall_exit = 0
    for workbook in sorted(set(selected)):
        relative = workbook.relative_to(workspace)
        report_directory = output_root / relative.parent / f"{relative.name}.report"
        command = _command(
            workbook,
            report_directory,
            mode=mode,
            config=config,
            baseline=baseline,
            new_only=new_only,
            fail_on=fail_on,
            source_scope=relative.as_posix(),
        )
        completed = subprocess.run(command, check=False)  # noqa: S603 - no shell
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
        "schema_version": 2,
        "mode": mode,
        "base_sha": base or None,
        "head_sha": head,
        "config": config.relative_to(workspace).as_posix() if config else None,
        "baseline": baseline.relative_to(workspace).as_posix() if baseline else None,
        "new_only": new_only,
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
        print("WorkbookLens: no changed .xlsx or .xlsm files to process")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
