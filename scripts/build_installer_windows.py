"""Build the official WorkbookLens Windows installer from the validated portable ZIP."""

from __future__ import annotations

import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Final

if __package__:
    from .check_installer_artifact import InstallerArtifactError, inspect_installer
    from .check_portable_artifact import (
        PortableArtifactError,
        extract_checked_artifact,
        inspect_artifact,
    )
else:
    from check_installer_artifact import InstallerArtifactError, inspect_installer
    from check_portable_artifact import (
        PortableArtifactError,
        extract_checked_artifact,
        inspect_artifact,
    )

VERSION_RE: Final = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._+-]*$")


class InstallerBuildError(RuntimeError):
    """The installer could not be built from the release portable artifact."""


def numeric_installer_version(version: str) -> str:
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)", version)
    if match is None:
        raise InstallerBuildError(
            f"version {version!r} cannot be represented as a Windows file version"
        )
    components = tuple(int(value) for value in match.groups())
    if any(value > 65535 for value in components):
        raise InstallerBuildError("Windows file-version components must be <= 65535")
    return ".".join((*map(str, components), "0"))


def resolve_iscc(explicit: Path | None = None) -> Path:
    if explicit is not None:
        candidate = explicit.resolve()
        if not candidate.is_file():
            raise InstallerBuildError(f"Inno Setup compiler does not exist: {candidate}")
        return candidate
    discovered = shutil.which("ISCC.exe") or shutil.which("iscc")
    if discovered:
        return Path(discovered).resolve()
    candidates: list[Path] = []
    for variable in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA"):
        base = os.environ.get(variable)
        if not base:
            continue
        root = Path(base)
        candidates.extend(
            [
                root / "Inno Setup 6" / "ISCC.exe",
                root / "Programs" / "Inno Setup 6" / "ISCC.exe",
            ]
        )
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    raise InstallerBuildError(
        "Inno Setup 6 compiler was not found; install it or pass --iscc explicitly"
    )


def _run(command: Sequence[os.PathLike[str] | str], *, cwd: Path) -> None:
    rendered = subprocess.list2cmdline([os.fspath(part) for part in command])
    print(f"+ {rendered}", flush=True)
    try:
        completed = subprocess.run(  # noqa: S603 - fixed local compiler command.
            [os.fspath(part) for part in command],
            cwd=cwd,
            check=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
    except subprocess.CalledProcessError as exc:
        raise InstallerBuildError(
            f"Inno Setup failed with exit code {exc.returncode}: {rendered}\n{exc.stdout or ''}"
        ) from exc
    except OSError as exc:
        raise InstallerBuildError(f"cannot run Inno Setup compiler: {exc}") from exc
    if completed.stdout:
        print(completed.stdout, end="" if completed.stdout.endswith("\n") else "\n")


def _write_sidecar(installer: Path) -> Path:
    digest = hashlib.sha256()
    with installer.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    sidecar = installer.with_suffix(installer.suffix + ".sha256")
    sidecar.write_text(f"{digest.hexdigest()}  {installer.name}\n", encoding="ascii", newline="\n")
    return sidecar


def _publish_validated_pair(
    staged_installer: Path,
    staged_sidecar: Path,
    installer: Path,
    sidecar: Path,
    *,
    overwrite: bool,
    validate: Callable[[Path], None],
) -> None:
    staged_pair = (staged_sidecar, staged_installer)
    final_pair = (sidecar, installer)
    for staged in staged_pair:
        if not staged.is_file() or staged.is_symlink():
            raise InstallerBuildError(f"staged installer output is missing or unsafe: {staged}")
    existing = [path for path in final_pair if path.exists() or path.is_symlink()]
    if existing and not overwrite:
        raise InstallerBuildError(
            f"installer output already exists; pass --overwrite to replace {installer.name}"
        )
    for path in existing:
        if path.is_symlink() or not path.is_file():
            raise InstallerBuildError(f"refusing to replace unsafe installer output: {path}")

    backup_dir = staged_installer.parent / ".previous-output"
    backup_dir.mkdir(exist_ok=False)
    backups: dict[Path, Path] = {}
    published: list[Path] = []
    try:
        for final in final_pair:
            if final.exists():
                backup = backup_dir / final.name
                os.replace(final, backup)
                backups[final] = backup
        for staged, final in zip(staged_pair, final_pair, strict=True):
            os.replace(staged, final)
            published.append(final)
        validate(installer)
    except Exception as exc:
        rollback_errors: list[str] = []
        for final in reversed(published):
            if final not in backups:
                try:
                    final.unlink(missing_ok=True)
                except OSError as rollback_exc:
                    rollback_errors.append(f"cannot remove {final}: {rollback_exc}")
        for final, backup in reversed(tuple(backups.items())):
            try:
                os.replace(backup, final)
            except OSError as rollback_exc:
                rollback_errors.append(f"cannot restore {final}: {rollback_exc}")
        detail = f"; rollback errors: {'; '.join(rollback_errors)}" if rollback_errors else ""
        raise InstallerBuildError(
            f"cannot publish validated installer outputs: {exc}{detail}"
        ) from exc


def build_installer(args: argparse.Namespace) -> tuple[Path, Path]:
    if os.name != "nt":
        raise InstallerBuildError("the Windows installer must be built on Windows")
    repository_root = args.repository_root.resolve()
    portable = args.portable_zip.resolve()
    report = inspect_artifact(
        portable,
        expected_version=args.expected_version,
        repository_root=repository_root,
    )
    if not VERSION_RE.fullmatch(report.version):
        raise InstallerBuildError(f"portable ZIP has an unsafe version: {report.version!r}")
    iscc = resolve_iscc(args.iscc)
    script = repository_root / "packaging" / "windows" / "WorkbookLens.iss"
    if not script.is_file():
        raise InstallerBuildError(f"Inno Setup definition is missing: {script}")
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    output_base = f"WorkbookLens-{report.version}-windows-x64-setup"
    installer = output_dir / f"{output_base}.exe"
    sidecar = installer.with_suffix(installer.suffix + ".sha256")
    if (
        installer.exists() or installer.is_symlink() or sidecar.exists() or sidecar.is_symlink()
    ) and not args.overwrite:
        raise InstallerBuildError(
            f"installer output already exists; pass --overwrite to replace {installer.name}"
        )

    work_parent = args.work_dir.resolve() if args.work_dir else None
    if work_parent is not None:
        work_parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(
        prefix=".workbooklens-installer-output-",
        dir=output_dir,
    ) as publish_raw:
        publish_dir = Path(publish_raw)
        with tempfile.TemporaryDirectory(prefix="workbooklens-installer-", dir=work_parent) as raw:
            scratch = Path(raw)
            portable_root = extract_checked_artifact(
                portable,
                scratch / "portable",
                expected_version=report.version,
                repository_root=repository_root,
            )
            command = [
                iscc,
                "/Qp",
                f"/DAppVersion={report.version}",
                f"/DNumericVersion={numeric_installer_version(report.version)}",
                f"/DPortableRoot={portable_root}",
                f"/DOutputDir={publish_dir}",
                f"/DOutputBaseFilename={output_base}",
                script,
            ]
            _run(command, cwd=repository_root)
        staged_installer = publish_dir / installer.name
        if not staged_installer.is_file():
            raise InstallerBuildError(
                f"Inno Setup did not produce the expected file: {staged_installer}"
            )
        staged_sidecar = _write_sidecar(staged_installer)
        inspect_installer(staged_installer, expected_version=report.version)
        _publish_validated_pair(
            staged_installer,
            staged_sidecar,
            installer,
            sidecar,
            overwrite=args.overwrite,
            validate=lambda path: inspect_installer(path, expected_version=report.version),
        )
    return installer, sidecar


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--portable-zip", type=Path, required=True)
    parser.add_argument("--expected-version")
    parser.add_argument("--iscc", type=Path)
    parser.add_argument(
        "--repository-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    parser.add_argument("--output-dir", type=Path, default=Path("dist"))
    parser.add_argument("--work-dir", type=Path)
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        installer, sidecar = build_installer(args)
    except (
        InstallerArtifactError,
        InstallerBuildError,
        PortableArtifactError,
        OSError,
    ) as exc:
        print(f"installer build failed: {exc}", file=sys.stderr)
        return 1
    print(f"Windows installer: {installer}")
    print(f"SHA256 sidecar: {sidecar}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
