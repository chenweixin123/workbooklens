from __future__ import annotations

import argparse
import email.policy
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from email.message import EmailMessage
from email.parser import BytesParser
from pathlib import Path, PurePosixPath
from typing import Any, Final

if __package__:
    from .check_portable_artifact import inspect_artifact
else:
    from check_portable_artifact import inspect_artifact


PYINSTALLER_VERSION: Final = "6.22.2"
REQUIRED_PYTHON: Final = (3, 12)
VERSION_RE: Final = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._+-]*$")
PROJECT_CANONICAL_NAME: Final = "workbooklens"
BUILD_MARKER: Final = ".workbooklens-portable-build"
PYTHON_INJECTION_VARIABLES: Final = {
    "__pyvenv_launcher__",
    "_old_virtual_path",
    "conda_default_env",
    "conda_prefix",
    "pyenv",
    "pythonexecutableroot",
    "pythonhome",
    "pythonpath",
    "virtual_env",
    "virtual_env_prompt",
}


class PortableBuildError(RuntimeError):
    pass


@dataclass(frozen=True)
class WheelMetadata:
    name: str
    version: str


@dataclass(frozen=True)
class PythonRuntime:
    executable: Path
    version: str
    version_info: tuple[int, int, int]
    machine: str
    pointer_bits: int
    base_prefix: Path


@dataclass(frozen=True)
class DistributionNotice:
    name: str
    version: str
    license_expression: str
    license_text: str
    copied_files: tuple[str, ...]


def _canonicalize_name(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).casefold()


def read_wheel_metadata(wheel: Path) -> WheelMetadata:
    if not wheel.is_file():
        raise PortableBuildError(f"wheel does not exist: {wheel}")
    try:
        with zipfile.ZipFile(wheel) as archive:
            metadata_names = [
                name
                for name in archive.namelist()
                if name.count("/") == 1 and name.endswith(".dist-info/METADATA")
            ]
            if len(metadata_names) != 1:
                raise PortableBuildError(
                    "wheel must contain exactly one top-level dist-info/METADATA file"
                )
            message = BytesParser(policy=email.policy.default).parsebytes(
                archive.read(metadata_names[0])
            )
    except (OSError, zipfile.BadZipFile, KeyError) as exc:
        raise PortableBuildError(f"cannot read wheel metadata: {exc}") from exc

    name = message.get("Name", "").strip()
    version = message.get("Version", "").strip()
    if _canonicalize_name(name) != PROJECT_CANONICAL_NAME:
        raise PortableBuildError(f"expected a WorkbookLens wheel, found {name!r}")
    if not VERSION_RE.fullmatch(version):
        raise PortableBuildError(f"wheel has an unsafe version string: {version!r}")
    return WheelMetadata(name=name, version=version)


def _display_command(command: Iterable[os.PathLike[str] | str]) -> str:
    return subprocess.list2cmdline([os.fspath(part) for part in command])


def sanitized_python_environment(
    source: dict[str, str] | None = None,
) -> dict[str, str]:
    environment = dict(os.environ if source is None else source)
    for key in tuple(environment):
        if key.casefold() in PYTHON_INJECTION_VARIABLES:
            environment.pop(key, None)
    environment["PYTHONNOUSERSITE"] = "1"
    return environment


def _run(
    command: list[os.PathLike[str] | str],
    *,
    cwd: Path | None = None,
    env: dict[str, str] | None = None,
) -> subprocess.CompletedProcess[str]:
    rendered = _display_command(command)
    print(f"+ {rendered}", flush=True)
    try:
        return subprocess.run(  # noqa: S603 - command is an explicit local build step.
            [os.fspath(part) for part in command],
            cwd=cwd,
            env=env,
            check=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
    except subprocess.CalledProcessError as exc:
        output = exc.stdout or ""
        raise PortableBuildError(
            f"command failed with exit code {exc.returncode}: {rendered}\n{output}"
        ) from exc
    except OSError as exc:
        raise PortableBuildError(f"cannot run command {rendered}: {exc}") from exc


def query_python(executable: Path) -> PythonRuntime:
    probe = (
        "import json, platform, struct, sys; "
        "print(json.dumps({'version': platform.python_version(), "
        "'version_info': list(sys.version_info[:3]), "
        "'machine': platform.machine(), "
        "'pointer_bits': struct.calcsize('P') * 8, "
        "'base_prefix': sys.base_prefix}))"
    )
    completed = _run(
        [executable, "-c", probe],
        env=sanitized_python_environment(),
    )
    try:
        payload: dict[str, Any] = json.loads(completed.stdout)
        version_info = tuple(int(part) for part in payload["version_info"])
        runtime = PythonRuntime(
            executable=executable.resolve(),
            version=str(payload["version"]),
            version_info=(version_info[0], version_info[1], version_info[2]),
            machine=str(payload["machine"]),
            pointer_bits=int(payload["pointer_bits"]),
            base_prefix=Path(str(payload["base_prefix"])).resolve(),
        )
    except (KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
        raise PortableBuildError("target Python returned invalid runtime metadata") from exc
    if runtime.version_info[:2] != REQUIRED_PYTHON:
        raise PortableBuildError(f"portable build requires Python 3.12, found {runtime.version}")
    if runtime.pointer_bits != 64 or runtime.machine.casefold() not in {
        "amd64",
        "x86_64",
    }:
        raise PortableBuildError("portable build requires an AMD64/x86_64 64-bit Python runtime")
    return runtime


def _venv_python(venv: Path) -> Path:
    return venv / "Scripts" / "python.exe"


def _query_site_packages(python: Path, *, env: dict[str, str]) -> Path:
    completed = _run(
        [
            python,
            "-c",
            "import sysconfig; print(sysconfig.get_paths()['purelib'])",
        ],
        env=env,
    )
    path = Path(completed.stdout.strip()).resolve()
    if not path.is_dir():
        raise PortableBuildError(f"isolated site-packages does not exist: {path}")
    return path


def _numeric_file_version(version: str) -> tuple[int, int, int, int]:
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)", version)
    if match is None:
        raise PortableBuildError(
            f"version {version!r} cannot be represented as a Windows file version"
        )
    parts = tuple(int(value) for value in match.groups())
    if any(value > 65535 for value in parts):
        raise PortableBuildError("Windows file-version components must be <= 65535")
    return (parts[0], parts[1], parts[2], 0)


def render_version_info(version: str) -> str:
    numeric = _numeric_file_version(version)
    return f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={numeric!r},
    prodvers={numeric!r},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', 'WorkbookLens'),
         StringStruct('FileDescription', 'WorkbookLens local spreadsheet auditor'),
         StringStruct('FileVersion', '{version}'),
         StringStruct('InternalName', 'WorkbookLens'),
         StringStruct('LegalCopyright', 'WorkbookLens contributors'),
         StringStruct('OriginalFilename', 'WorkbookLens.exe'),
         StringStruct('ProductName', 'WorkbookLens'),
         StringStruct('ProductVersion', '{version}')])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""


def _safe_license_component(value: str) -> str:
    cleaned = re.sub(r"[<>:\"/\\|?*\x00-\x1f]+", "-", value).strip(" .")
    cleaned = re.sub(r"\s+", "-", cleaned)
    if not cleaned:
        return "unknown"
    stem = cleaned.split(".", 1)[0].upper()
    if stem in {"AUX", "CON", "NUL", "PRN"} or re.fullmatch(r"(?:COM|LPT)[1-9]", stem):
        cleaned = f"package-{cleaned}"
    return cleaned[:120]


def _metadata_license_files(dist_info: Path, message: EmailMessage) -> list[Path]:
    declared: set[str] = set()
    for raw_value in message.get_all("License-File", []):
        value = str(raw_value).strip()
        relative = PurePosixPath(value)
        if (
            not value
            or "\\" in value
            or relative.is_absolute()
            or any(part in {"", ".", ".."} for part in relative.parts)
        ):
            raise PortableBuildError(
                f"{dist_info.name} declares an unsafe License-File path: {value!r}"
            )
        declared.add(relative.as_posix().casefold())

    result: list[Path] = []
    found_declared: set[str] = set()
    for candidate in dist_info.rglob("*"):
        if not candidate.is_file() or candidate.is_symlink():
            continue
        relative = candidate.relative_to(dist_info)
        folded_parts = tuple(part.casefold() for part in relative.parts)
        leaf = folded_parts[-1]
        relative_name = PurePosixPath(*relative.parts).as_posix().casefold()
        declared_name = (
            PurePosixPath(*relative.parts[1:]).as_posix().casefold()
            if folded_parts[0] == "licenses" and len(relative.parts) > 1
            else relative_name
        )
        is_declared = relative_name in declared or declared_name in declared
        if is_declared:
            found_declared.update({relative_name, declared_name} & declared)
        if (
            folded_parts[0] == "licenses"
            or leaf.startswith(("license", "licence", "copying", "notice"))
            or is_declared
        ):
            result.append(candidate)
    missing = sorted(declared - found_declared)
    if missing:
        raise PortableBuildError(
            f"{dist_info.name} is missing declared License-File entries: {', '.join(missing)}"
        )
    return sorted(result, key=lambda path: path.as_posix().casefold())


def _metadata_message(dist_info: Path) -> EmailMessage:
    metadata = dist_info / "METADATA"
    if not metadata.is_file():
        raise PortableBuildError(f"distribution metadata is missing: {metadata}")
    try:
        parsed = BytesParser(policy=email.policy.default).parsebytes(metadata.read_bytes())
    except (OSError, UnicodeError) as exc:
        raise PortableBuildError(f"cannot parse {metadata}: {exc}") from exc
    if not isinstance(parsed, EmailMessage):
        raise PortableBuildError(f"unexpected metadata message type: {metadata}")
    return parsed


def collect_distribution_licenses(
    site_packages: Path,
    licenses_dir: Path,
) -> list[DistributionNotice]:
    notices: list[DistributionNotice] = []
    used_directories: set[str] = set()
    for dist_info in sorted(
        site_packages.glob("*.dist-info"), key=lambda path: path.name.casefold()
    ):
        message = _metadata_message(dist_info)
        name = str(message.get("Name", dist_info.stem)).strip()
        version = str(message.get("Version", "unknown")).strip()
        license_expression = str(message.get("License-Expression", "")).strip()
        license_text = str(message.get("License", "")).strip()
        classifiers = [
            str(value).strip()
            for value in message.get_all("Classifier", [])
            if str(value).startswith("License ::")
        ]
        if not license_text and classifiers:
            license_text = "; ".join(classifiers)

        base_name = _safe_license_component(f"{name}-{version}")
        unique_name = base_name
        counter = 2
        while unique_name.casefold() in used_directories:
            unique_name = f"{base_name}-{counter}"
            counter += 1
        used_directories.add(unique_name.casefold())
        destination = licenses_dir / unique_name
        destination.mkdir(parents=True, exist_ok=False)

        copied: list[str] = []
        for source in _metadata_license_files(dist_info, message):
            relative = source.relative_to(dist_info)
            safe_relative = Path(*(_safe_license_component(part) for part in relative.parts))
            target = destination / safe_relative
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                raise PortableBuildError(f"license filename collision: {target}")
            shutil.copyfile(source, target)
            copied.append(target.relative_to(licenses_dir).as_posix())

        if not copied:
            fallback = destination / "LICENSE-METADATA.txt"
            fallback.write_text(
                "\n".join(
                    [
                        f"Name: {name}",
                        f"Version: {version}",
                        f"License-Expression: {license_expression or 'not declared'}",
                        f"License: {license_text or 'not declared'}",
                        "",
                        "No license file was present in this distribution's dist-info ",
                        "directory. Consult the upstream project before redistribution.",
                        "",
                    ]
                ),
                encoding="utf-8",
            )
            copied.append(fallback.relative_to(licenses_dir).as_posix())

        notices.append(
            DistributionNotice(
                name=name,
                version=version,
                license_expression=license_expression,
                license_text=license_text,
                copied_files=tuple(copied),
            )
        )
    if not notices:
        raise PortableBuildError("isolated environment has no dist-info license metadata")
    return notices


def locate_cpython_license(
    runtime: PythonRuntime,
    explicit: Path | None,
) -> Path:
    candidates = [
        explicit,
        runtime.base_prefix / "LICENSE.txt",
        runtime.base_prefix / "LICENSE",
        runtime.executable.parent / "LICENSE.txt",
        runtime.executable.parent / "LICENSE",
    ]
    for candidate in candidates:
        if candidate is not None and candidate.is_file():
            return candidate.resolve()
    raise PortableBuildError("cannot locate the CPython license; pass --cpython-license explicitly")


def render_third_party_notices(
    runtime: PythonRuntime,
    distributions: list[DistributionNotice],
) -> str:
    lines = [
        "WorkbookLens Portable - Third-Party Notices",
        "============================================",
        "",
        "This portable distribution includes CPython and third-party Python packages.",
        "The corresponding license files copied from the isolated build environment",
        "are stored under LICENSES.",
        "",
        f"CPython {runtime.version}",
        "License file: LICENSES/CPython-3.12-LICENSE.txt",
        "",
    ]
    for item in sorted(distributions, key=lambda value: value.name.casefold()):
        lines.extend(
            [
                f"{item.name} {item.version}",
                f"License-Expression: {item.license_expression or 'not declared'}",
                f"License metadata: {item.license_text or 'not declared'}",
                "License files:",
                *(f"  - LICENSES/{path}" for path in item.copied_files),
                "",
            ]
        )
    return "\n".join(lines)


def _copy_portable_root_files(
    repository_root: Path,
    staging_root: Path,
    version: str,
) -> None:
    shutil.copyfile(repository_root / "LICENSE", staging_root / "LICENSE")
    shutil.copyfile(
        repository_root / "workbooklens.example.yml",
        staging_root / "workbooklens.example.yml",
    )
    shutil.copyfile(
        repository_root / "packaging" / "windows" / "Start-WorkbookLens.cmd",
        staging_root / "Start-WorkbookLens.cmd",
    )
    readme_template = (repository_root / "packaging" / "windows" / "README-PORTABLE.txt").read_text(
        encoding="utf-8"
    )
    if readme_template.count("@VERSION@") != 1:
        raise PortableBuildError(
            "README-PORTABLE.txt must contain exactly one @VERSION@ placeholder"
        )
    (staging_root / "README-PORTABLE.txt").write_text(
        readme_template.replace("@VERSION@", version),
        encoding="utf-8",
        newline="\n",
    )


def create_deterministic_zip(source_root: Path, archive: Path) -> None:
    source_root = source_root.resolve()
    files = sorted(
        (path for path in source_root.rglob("*") if path.is_file()),
        key=lambda path: path.relative_to(source_root.parent).as_posix().casefold(),
    )
    if not files:
        raise PortableBuildError(f"portable staging directory is empty: {source_root}")
    for path in source_root.rglob("*"):
        if path.is_symlink():
            raise PortableBuildError(f"portable staging contains a symlink: {path}")

    archive.parent.mkdir(parents=True, exist_ok=True)
    temporary = archive.with_suffix(archive.suffix + ".tmp")
    if temporary.exists():
        temporary.unlink()
    try:
        with zipfile.ZipFile(
            temporary,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=9,
            allowZip64=True,
        ) as output:
            for source in files:
                relative = source.relative_to(source_root.parent).as_posix()
                info = zipfile.ZipInfo(relative, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 3
                executable = source.suffix.casefold() in {".cmd", ".exe"}
                mode = 0o755 if executable else 0o644
                info.external_attr = (stat.S_IFREG | mode) << 16
                with (
                    source.open("rb") as input_stream,
                    output.open(info, "w", force_zip64=True) as output_stream,
                ):
                    shutil.copyfileobj(input_stream, output_stream, length=1024 * 1024)
        os.replace(temporary, archive)
    finally:
        if temporary.exists():
            temporary.unlink()


def _write_sha256_sidecar(archive: Path) -> Path:
    digest = hashlib.sha256()
    with archive.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    sidecar = archive.with_suffix(archive.suffix + ".sha256")
    sidecar.write_text(
        f"{digest.hexdigest()}  {archive.name}\n",
        encoding="ascii",
        newline="\n",
    )
    return sidecar


def _safe_remove_owned_build(path: Path) -> None:
    marker = path / BUILD_MARKER
    if not marker.is_file() or path.parent == path:
        raise PortableBuildError(f"refusing to remove unowned build directory: {path}")
    shutil.rmtree(path)


def build_portable(args: argparse.Namespace) -> tuple[Path, Path]:
    if os.name != "nt":
        raise PortableBuildError("Windows portable artifacts must be built on Windows")

    repository_root = args.repository_root.resolve()
    wheel = args.wheel.resolve()
    constraints = args.constraints.resolve() if args.constraints else None
    if constraints is not None and not constraints.is_file():
        raise PortableBuildError(f"constraints file does not exist: {constraints}")
    metadata = read_wheel_metadata(wheel)
    if args.expected_version and metadata.version != args.expected_version:
        raise PortableBuildError(
            f"wheel version {metadata.version!r} does not match {args.expected_version!r}"
        )

    runtime = query_python(args.python.resolve())
    cpython_license = locate_cpython_license(runtime, args.cpython_license)
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    archive = output_dir / (f"WorkbookLens-{metadata.version}-windows-x64-portable.zip")
    sidecar = archive.with_suffix(archive.suffix + ".sha256")
    if not args.overwrite and (archive.exists() or sidecar.exists()):
        raise PortableBuildError(
            f"portable output already exists; pass --overwrite to replace {archive.name}"
        )

    scratch_parent = args.work_dir.resolve() if args.work_dir else None
    if scratch_parent is not None:
        scratch_parent.mkdir(parents=True, exist_ok=True)
    scratch = Path(tempfile.mkdtemp(prefix="workbooklens-portable-", dir=scratch_parent)).resolve()
    (scratch / BUILD_MARKER).write_text("owned by build_portable_windows.py\n")

    try:
        isolated_env = sanitized_python_environment()
        venv = scratch / "venv"
        _run([runtime.executable, "-m", "venv", venv], env=isolated_env)
        venv_python = _venv_python(venv)
        if not venv_python.is_file():
            raise PortableBuildError(f"isolated Python was not created: {venv_python}")
        install_command: list[os.PathLike[str] | str] = [
            venv_python,
            "-m",
            "pip",
            "install",
            "--disable-pip-version-check",
            "--no-input",
        ]
        if constraints is not None:
            install_command.extend(["--constraint", constraints])
        install_command.extend([f"pyinstaller=={PYINSTALLER_VERSION}", wheel])
        _run(install_command, env=isolated_env)
        verification = _run(
            [
                venv_python,
                "-c",
                (
                    "import json, PyInstaller, workbooklens; "
                    "print(json.dumps({'pyinstaller': PyInstaller.__version__, "
                    "'workbooklens': workbooklens.__version__}))"
                ),
            ],
            env=isolated_env,
        )
        installed = json.loads(verification.stdout)
        if installed != {
            "pyinstaller": PYINSTALLER_VERSION,
            "workbooklens": metadata.version,
        }:
            raise PortableBuildError(f"isolated environment version mismatch: {installed!r}")

        version_file = scratch / "version_info.txt"
        version_file.write_text(
            render_version_info(metadata.version), encoding="utf-8", newline="\n"
        )
        pyinstaller_env = isolated_env.copy()
        pyinstaller_env["WORKBOOKLENS_VERSION_FILE"] = str(version_file)
        pyinstaller_env["PYINSTALLER_CONFIG_DIR"] = str(scratch / "pyinstaller-config")
        raw_dist = scratch / "pyinstaller-dist"
        _run(
            [
                venv_python,
                "-m",
                "PyInstaller",
                "--noconfirm",
                "--clean",
                "--distpath",
                raw_dist,
                "--workpath",
                scratch / "pyinstaller-build",
                repository_root / "packaging" / "windows" / "WorkbookLens.spec",
            ],
            cwd=repository_root,
            env=pyinstaller_env,
        )

        pyinstaller_output = raw_dist / "WorkbookLens"
        if not (pyinstaller_output / "WorkbookLens.exe").is_file():
            raise PortableBuildError("PyInstaller did not produce WorkbookLens.exe")
        if not (pyinstaller_output / "_internal").is_dir():
            raise PortableBuildError("PyInstaller did not produce an _internal directory")

        root_name = f"WorkbookLens-{metadata.version}-windows-x64"
        staging_root = scratch / "staging" / root_name
        shutil.copytree(pyinstaller_output, staging_root)
        _copy_portable_root_files(repository_root, staging_root, metadata.version)

        licenses_dir = staging_root / "LICENSES"
        licenses_dir.mkdir(parents=True, exist_ok=False)
        shutil.copyfile(
            cpython_license,
            licenses_dir / "CPython-3.12-LICENSE.txt",
        )
        distributions = collect_distribution_licenses(
            _query_site_packages(venv_python, env=isolated_env),
            licenses_dir,
        )
        (staging_root / "THIRD-PARTY-NOTICES.txt").write_text(
            render_third_party_notices(runtime, distributions),
            encoding="utf-8",
            newline="\n",
        )

        if args.overwrite:
            archive.unlink(missing_ok=True)
            sidecar.unlink(missing_ok=True)
        create_deterministic_zip(staging_root, archive)
        sidecar = _write_sha256_sidecar(archive)
        inspect_artifact(
            archive,
            expected_version=metadata.version,
            repository_root=repository_root,
        )
        return archive, sidecar
    finally:
        if args.keep_work_dir:
            print(f"portable build workspace retained at {scratch}")
        else:
            _safe_remove_owned_build(scratch)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Build the WorkbookLens Windows x64 PyInstaller onedir portable ZIP "
            "from an already-built wheel."
        )
    )
    parser.add_argument("--wheel", type=Path, required=True)
    parser.add_argument("--python", type=Path, default=Path(sys.executable))
    parser.add_argument("--expected-version")
    parser.add_argument(
        "--repository-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    parser.add_argument("--output-dir", type=Path, default=Path("dist"))
    parser.add_argument("--work-dir", type=Path)
    parser.add_argument("--cpython-license", type=Path)
    parser.add_argument(
        "--constraints",
        type=Path,
        help="Optional pip constraints file for reproducible dependency resolution.",
    )
    parser.add_argument("--keep-work-dir", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        archive, sidecar = build_portable(args)
    except (PortableBuildError, OSError, zipfile.BadZipFile) as exc:
        print(f"portable build failed: {exc}", file=sys.stderr)
        return 1
    print(f"portable archive: {archive}")
    print(f"SHA256 sidecar: {sidecar}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
