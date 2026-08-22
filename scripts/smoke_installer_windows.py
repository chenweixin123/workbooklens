"""Install, exercise, and uninstall the exact WorkbookLens Windows setup executable."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import winreg
from contextlib import suppress
from pathlib import Path
from typing import Any, Final

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

UNINSTALL_KEY: Final = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
SHELL_FOLDERS_KEY: Final = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
APP_ID: Final = "{7B7534E0-8485-4F4F-8DE7-561869FF7C0C}"
OWNED_UNINSTALL_REGISTRY_KEY: Final = f"{APP_ID}_is1"
OWNERSHIP_MARKER_NAME: Final = ".workbooklens-install-owner"
OWNERSHIP_MARKER_CONTENT: Final = f"WorkbookLens|{APP_ID}|owner-schema=1"
UPGRADE_STALE_MEMBERS: Final = (
    "_internal/workbooklens-upgrade-stale.bin",
    "LICENSES/workbooklens-upgrade-stale.txt",
)


class InstallerSmokeError(RuntimeError):
    """The setup executable failed an end-to-end user installation check."""


def _run(
    command: list[str],
    *,
    timeout: float = 120.0,
    env: dict[str, str] | None = None,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    rendered = subprocess.list2cmdline(command)
    print(f"+ {rendered}", flush=True)
    try:
        return subprocess.run(  # noqa: S603 - explicit local release executable.
            command,
            check=check,
            timeout=timeout,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            env=env,
        )
    except subprocess.CalledProcessError as exc:
        raise InstallerSmokeError(
            f"command failed with exit code {exc.returncode}: {rendered}\n{exc.stdout or ''}"
        ) from exc
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise InstallerSmokeError(f"cannot complete command {rendered}: {exc}") from exc


def _shell_folder(name: str) -> Path:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SHELL_FOLDERS_KEY) as key:
        raw, _ = winreg.QueryValueEx(key, name)
    return Path(os.path.expandvars(str(raw))).resolve()


def _default_install_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise InstallerSmokeError("LOCALAPPDATA is unavailable")
    return (Path(local_app_data) / "Programs" / "WorkbookLens").resolve()


def _uninstall_entries() -> list[dict[str, str]]:
    entries: list[dict[str, str]] = []
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY)
    except FileNotFoundError:
        return entries
    with root:
        for index in range(winreg.QueryInfoKey(root)[0]):
            name = winreg.EnumKey(root, index)
            try:
                child = winreg.OpenKey(root, name)
            except OSError:
                continue
            with child:
                values: dict[str, str] = {"RegistryKey": name}
                for value_name in ("DisplayName", "DisplayVersion", "InstallLocation"):
                    try:
                        value, _ = winreg.QueryValueEx(child, value_name)
                    except FileNotFoundError:
                        continue
                    values[value_name] = str(value)
                if values.get("DisplayName") == "WorkbookLens":
                    entries.append(values)
    return entries


def _owned_uninstall_entry() -> dict[str, str] | None:
    path = f"{UNINSTALL_KEY}\\{OWNED_UNINSTALL_REGISTRY_KEY}"
    try:
        child = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
    except FileNotFoundError:
        return None
    with child:
        values: dict[str, str] = {"RegistryKey": OWNED_UNINSTALL_REGISTRY_KEY}
        for value_name in ("DisplayName", "DisplayVersion", "InstallLocation"):
            try:
                value, _ = winreg.QueryValueEx(child, value_name)
            except FileNotFoundError:
                continue
            values[value_name] = str(value)
    return values


def _shortcut_details(path: Path) -> dict[str, str]:
    command = (
        "$path=[Environment]::GetEnvironmentVariable('WORKBOOKLENS_SHORTCUT_PATH');"
        "$shortcut=(New-Object -ComObject WScript.Shell).CreateShortcut($path);"
        "[pscustomobject]@{TargetPath=$shortcut.TargetPath;Arguments=$shortcut.Arguments;"
        "WorkingDirectory=$shortcut.WorkingDirectory}|ConvertTo-Json -Compress"
    )
    environment = os.environ.copy()
    environment["WORKBOOKLENS_SHORTCUT_PATH"] = str(path)
    completed = _run(
        ["powershell.exe", "-NoProfile", "-Command", command],
        env=environment,
    )
    try:
        payload: dict[str, Any] = json.loads(completed.stdout)
    except (TypeError, json.JSONDecodeError) as exc:
        raise InstallerSmokeError(f"cannot inspect shortcut {path}: {exc}") from exc
    return {key: str(value) for key, value in payload.items()}


def _file_hashes(root: Path) -> dict[str, str]:
    import hashlib

    result: dict[str, str] = {}
    for path in sorted(root.rglob("*")):
        if path.is_file():
            result[path.relative_to(root).as_posix()] = hashlib.sha256(
                path.read_bytes()
            ).hexdigest()
    return result


def _installed_payload_hashes(root: Path) -> dict[str, str]:
    return {
        name: digest
        for name, digest in _file_hashes(root).items()
        if not name.casefold().startswith("unins000.")
        and name.casefold() != OWNERSHIP_MARKER_NAME.casefold()
    }


def _write_test_uninstall_entry(install_dir: Path, *, version: str) -> None:
    path = f"{UNINSTALL_KEY}\\{OWNED_UNINSTALL_REGISTRY_KEY}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "WorkbookLens")
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, version)
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(install_dir))


def _wait_removed(path: Path, *, timeout: float = 20.0) -> None:
    deadline = time.monotonic() + timeout
    while path.exists() and time.monotonic() < deadline:
        time.sleep(0.2)
    if path.exists():
        raise InstallerSmokeError(f"uninstaller did not remove {path}")


def _same_path(raw: str | None, expected: Path) -> bool:
    if not raw:
        return False
    try:
        return Path(raw).expanduser().resolve() == expected.resolve()
    except OSError:
        return False


def _is_reparse_path(path: Path) -> bool:
    is_junction = getattr(path, "is_junction", None)
    return path.is_symlink() or bool(is_junction is not None and is_junction())


def _create_directory_junction(link: Path, target: Path) -> None:
    if link.exists() or _is_reparse_path(link):
        raise InstallerSmokeError(f"junction path already exists: {link}")
    target.mkdir(parents=True, exist_ok=True)
    link.parent.mkdir(parents=True, exist_ok=True)
    _run(["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target)])
    if not _is_reparse_path(link):
        raise InstallerSmokeError(f"mklink did not create a reparse point: {link}")


def _remove_directory_junction(link: Path) -> None:
    if not _is_reparse_path(link):
        if link.exists():
            raise InstallerSmokeError(f"refusing to remove non-reparse path as junction: {link}")
        return
    try:
        link.rmdir()
    except OSError as exc:
        raise InstallerSmokeError(f"cannot remove test junction {link}: {exc}") from exc


def _assert_reparse_target_rejected(
    installer: Path,
    log: Path,
    *,
    install_dir: Path,
    link: Path,
    target: Path,
    group: Path,
    desktop_link: Path,
    reason: str,
) -> None:
    expected_hashes = _file_hashes(target)
    _create_directory_junction(link, target)
    try:
        _assert_setup_rejected(
            installer,
            log,
            install_dir=None,
            reason=reason,
        )
        if _file_hashes(target) != expected_hashes:
            raise InstallerSmokeError(
                f"installer modified the target of a rejected reparse point: {reason}"
            )
        if (
            (install_dir / OWNERSHIP_MARKER_NAME).exists()
            or (install_dir / "unins000.exe").exists()
            or _owned_uninstall_entry() is not None
            or group.exists()
            or desktop_link.exists()
        ):
            raise InstallerSmokeError(
                f"rejected reparse-point target left installer-owned state: {reason}"
            )
    finally:
        _remove_directory_junction(link)
    if _file_hashes(target) != expected_hashes:
        raise InstallerSmokeError(f"junction cleanup modified the protected target: {reason}")


def _remove_owned_shortcut(
    shortcut: Path,
    install_dir: Path,
    errors: list[str],
) -> None:
    if not shortcut.exists() and not shortcut.is_symlink():
        return
    if shortcut.is_symlink() or not shortcut.is_file():
        errors.append(f"refusing to remove unsafe shortcut path: {shortcut}")
        return
    try:
        details = _shortcut_details(shortcut)
        target = Path(details.get("TargetPath", "")).resolve()
        target.relative_to(install_dir.resolve())
    except (InstallerSmokeError, OSError, ValueError) as exc:
        errors.append(f"cannot prove shortcut belongs to test install {shortcut}: {exc}")
        return
    try:
        shortcut.unlink()
    except OSError as exc:
        errors.append(f"cannot remove test shortcut {shortcut}: {exc}")


def _remove_owned_registry_entry(install_dir: Path, errors: list[str]) -> None:
    entry = _owned_uninstall_entry()
    if entry is None:
        return
    if entry.get("DisplayName") != "WorkbookLens" or not _same_path(
        entry.get("InstallLocation"),
        install_dir,
    ):
        errors.append(
            "refusing to remove uninstall entry because its name or install path is not test-owned"
        )
        return
    path = f"{UNINSTALL_KEY}\\{OWNED_UNINSTALL_REGISTRY_KEY}"
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except OSError as exc:
        errors.append(f"cannot remove test uninstall entry: {exc}")


def _cleanup_owned_state(
    *,
    install_dir: Path,
    scratch: Path,
    group: Path,
    desktop_link: Path,
) -> list[str]:
    errors: list[str] = []
    uninstaller = install_dir / "unins000.exe"
    if uninstaller.is_file() and not uninstaller.is_symlink():
        try:
            completed = subprocess.run(  # noqa: S603 - exact test-owned uninstaller.
                [
                    str(uninstaller),
                    "/VERYSILENT",
                    "/SUPPRESSMSGBOXES",
                    "/NORESTART",
                ],
                timeout=120,
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            if completed.returncode != 0:
                errors.append(f"test uninstaller returned {completed.returncode}")
            try:
                _wait_removed(install_dir)
            except InstallerSmokeError as exc:
                errors.append(str(exc))
        except (OSError, subprocess.TimeoutExpired) as exc:
            errors.append(f"cannot run test uninstaller during cleanup: {exc}")

    app_link = group / "WorkbookLens.lnk"
    uninstall_link = group / "Uninstall WorkbookLens.lnk"
    for shortcut in (app_link, uninstall_link, desktop_link):
        _remove_owned_shortcut(shortcut, install_dir, errors)
    _remove_owned_registry_entry(install_dir, errors)

    if install_dir.exists() or install_dir.is_symlink():
        if install_dir.is_symlink():
            errors.append(f"refusing to remove symbolic-link install directory: {install_dir}")
        elif _same_path(str(install_dir), _default_install_dir()):
            try:
                install_dir.rmdir()
            except OSError as exc:
                errors.append(
                    "refusing to recursively remove a non-empty default install "
                    f"directory {install_dir}: {exc}"
                )
        else:
            try:
                resolved_install = install_dir.resolve()
                resolved_install.relative_to(scratch.resolve())
                shutil.rmtree(install_dir)
            except (OSError, ValueError) as exc:
                errors.append(f"cannot remove test-owned install directory {install_dir}: {exc}")
    if group.exists():
        try:
            group.rmdir()
        except OSError as exc:
            errors.append(f"cannot remove non-empty test Start-menu group {group}: {exc}")
    if (
        group.exists()
        or desktop_link.exists()
        or _owned_uninstall_entry() is not None
        or install_dir.exists()
    ):
        errors.append("test-owned installer state remains after cleanup")
    return errors


def _record_cleanup_errors(
    primary_error: BaseException | None,
    cleanup_errors: list[str],
) -> None:
    if not cleanup_errors:
        return
    detail = "installer smoke cleanup issues: " + "; ".join(cleanup_errors)
    if primary_error is not None:
        primary_error.add_note(detail)
        return
    raise InstallerSmokeError(detail)


def _setup_command(
    installer: Path,
    log: Path,
    *,
    install_dir: Path | None,
) -> list[str]:
    command = [
        str(installer.resolve()),
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/SP-",
        "/TASKS=desktopicon",
        f"/LOG={log}",
    ]
    if install_dir is not None:
        command.append(f"/DIR={install_dir}")
    return command


def _install_setup(
    installer: Path,
    log: Path,
    *,
    install_dir: Path | None,
) -> None:
    _run(
        _setup_command(installer, log, install_dir=install_dir),
        timeout=180.0,
    )


def _assert_setup_rejected(
    installer: Path,
    log: Path,
    *,
    install_dir: Path | None,
    reason: str,
) -> None:
    completed = _run(
        _setup_command(installer, log, install_dir=install_dir),
        timeout=180.0,
        check=False,
    )
    if completed.returncode == 0:
        raise InstallerSmokeError(f"installer accepted unsafe target: {reason}")


def _assert_payload_matches(
    install_dir: Path,
    expected_hashes: dict[str, str],
    *,
    phase: str,
) -> None:
    actual = _installed_payload_hashes(install_dir)
    if actual != expected_hashes:
        missing = sorted(set(expected_hashes) - set(actual))
        extra = sorted(set(actual) - set(expected_hashes))
        changed = sorted(
            name
            for name in set(actual) & set(expected_hashes)
            if actual[name] != expected_hashes[name]
        )
        raise InstallerSmokeError(
            f"{phase} payload differs from its portable ZIP "
            f"(missing={missing!r}, extra={extra!r}, changed={changed!r})"
        )


def smoke_installer(
    installer: Path,
    portable_zip: Path,
    *,
    expected_version: str | None,
    repository_root: Path,
    previous_installer: Path | None = None,
    previous_portable_zip: Path | None = None,
) -> None:
    if os.name != "nt":
        raise InstallerSmokeError("the installer smoke test requires Windows")
    if (previous_installer is None) != (previous_portable_zip is None):
        raise InstallerSmokeError(
            "--previous-installer and --previous-portable-zip must be provided together"
        )
    installer_report = inspect_installer(installer, expected_version=expected_version)
    inspect_artifact(
        portable_zip,
        expected_version=installer_report.version,
        repository_root=repository_root,
    )
    baseline_installer = installer
    baseline_portable = portable_zip
    baseline_version = installer_report.version
    if previous_installer is not None and previous_portable_zip is not None:
        previous_report = inspect_installer(previous_installer)
        inspect_artifact(
            previous_portable_zip,
            expected_version=previous_report.version,
            repository_root=None,
        )
        baseline_installer = previous_installer
        baseline_portable = previous_portable_zip
        baseline_version = previous_report.version
    programs = _shell_folder("Programs")
    desktop = _shell_folder("Desktop")
    group = programs / "WorkbookLens"
    desktop_link = desktop / "WorkbookLens.lnk"
    install_dir = _default_install_dir()
    if (
        group.exists()
        or desktop_link.exists()
        or _uninstall_entries()
        or install_dir.exists()
        or install_dir.is_symlink()
    ):
        raise InstallerSmokeError(
            "an existing WorkbookLens installation, shortcut, or default install "
            "directory was found; refusing to overwrite it"
        )

    with tempfile.TemporaryDirectory(prefix="workbooklens-installer-smoke-") as raw:
        scratch = Path(raw)
        unsafe_install_dir = scratch / "unregistered custom target"
        portable_root = extract_checked_artifact(
            portable_zip,
            scratch / "portable-current",
            expected_version=installer_report.version,
            repository_root=repository_root,
        )
        expected_hashes = _file_hashes(portable_root)
        if baseline_portable == portable_zip:
            baseline_root = portable_root
        else:
            baseline_root = extract_checked_artifact(
                baseline_portable,
                scratch / "portable-previous",
                expected_version=baseline_version,
                repository_root=None,
            )
        baseline_hashes = _file_hashes(baseline_root)
        primary_error: BaseException | None = None
        try:
            forged_target = scratch / "forged registry victim"
            forged_internal = forged_target / "_internal"
            forged_licenses = forged_target / "LICENSES"
            forged_internal.mkdir(parents=True)
            forged_licenses.mkdir()
            (forged_internal / "foreign.bin").write_bytes(
                b"foreign internal data that must survive\n"
            )
            (forged_licenses / "foreign.txt").write_bytes(
                b"foreign license data that must survive\n"
            )
            forged_hashes = _file_hashes(forged_target)
            _write_test_uninstall_entry(forged_target, version="0.0.0-forged")
            forged_error: BaseException | None = None
            try:
                _assert_setup_rejected(
                    installer,
                    scratch / "reject-forged-registry-target.log",
                    install_dir=forged_target,
                    reason="HKCU InstallLocation pointed at an arbitrary non-default directory",
                )
                if _file_hashes(forged_target) != forged_hashes:
                    raise InstallerSmokeError(
                        "installer modified files referenced only by a forged uninstall entry"
                    )
                forged_entry = _owned_uninstall_entry()
                if forged_entry is None or not _same_path(
                    forged_entry.get("InstallLocation"),
                    forged_target,
                ):
                    raise InstallerSmokeError(
                        "rejected forged uninstall entry was unexpectedly replaced"
                    )
                if install_dir.exists() or group.exists() or desktop_link.exists():
                    raise InstallerSmokeError(
                        "rejected forged uninstall entry left installer-owned state"
                    )
            except BaseException as exc:
                forged_error = exc
                raise
            finally:
                forged_cleanup_errors: list[str] = []
                _remove_owned_registry_entry(forged_target, forged_cleanup_errors)
                _record_cleanup_errors(forged_error, forged_cleanup_errors)

            unsafe_error: BaseException | None = None
            try:
                _assert_setup_rejected(
                    installer,
                    scratch / "reject-custom-dir.log",
                    install_dir=unsafe_install_dir,
                    reason="fresh install used a non-default /DIR target",
                )
                if (
                    unsafe_install_dir.exists()
                    or unsafe_install_dir.is_symlink()
                    or _owned_uninstall_entry() is not None
                    or group.exists()
                    or desktop_link.exists()
                ):
                    raise InstallerSmokeError("rejected custom target left installer-owned state")
            except BaseException as exc:
                unsafe_error = exc
                raise
            finally:
                unsafe_cleanup_errors = _cleanup_owned_state(
                    install_dir=unsafe_install_dir,
                    scratch=scratch,
                    group=group,
                    desktop_link=desktop_link,
                )
                _record_cleanup_errors(unsafe_error, unsafe_cleanup_errors)

            root_junction_target = scratch / "root junction target"
            root_junction_target.mkdir()
            (root_junction_target / "protected.bin").write_bytes(
                b"root junction target must survive\n"
            )
            _assert_reparse_target_rejected(
                installer,
                scratch / "reject-root-reparse.log",
                install_dir=install_dir,
                link=install_dir,
                target=root_junction_target,
                group=group,
                desktop_link=desktop_link,
                reason="the fixed application directory was a junction",
            )
            if install_dir.exists() or _is_reparse_path(install_dir):
                raise InstallerSmokeError("root reparse-point test left the install path")

            internal_junction_target = scratch / "internal junction target"
            internal_junction_target.mkdir()
            (internal_junction_target / "protected.bin").write_bytes(
                b"nested internal junction target must survive\n"
            )
            internal_level = install_dir / "_internal" / "nested"
            internal_level.mkdir(parents=True)
            try:
                _assert_reparse_target_rejected(
                    installer,
                    scratch / "reject-nested-internal-reparse.log",
                    install_dir=install_dir,
                    link=internal_level / "redirect",
                    target=internal_junction_target,
                    group=group,
                    desktop_link=desktop_link,
                    reason="a nested _internal directory was a junction",
                )
            finally:
                with suppress(OSError):
                    internal_level.rmdir()
                with suppress(OSError):
                    internal_level.parent.rmdir()
                with suppress(OSError):
                    install_dir.rmdir()

            licenses_junction_target = scratch / "licenses junction target"
            licenses_junction_target.mkdir()
            (licenses_junction_target / "protected.txt").write_bytes(
                b"licenses junction target must survive\n"
            )
            install_dir.mkdir()
            try:
                _assert_reparse_target_rejected(
                    installer,
                    scratch / "reject-licenses-reparse.log",
                    install_dir=install_dir,
                    link=install_dir / "LICENSES",
                    target=licenses_junction_target,
                    group=group,
                    desktop_link=desktop_link,
                    reason="the LICENSES directory was a junction",
                )
            finally:
                with suppress(OSError):
                    install_dir.rmdir()

            foreign_payload = b"foreign data that WorkbookLens must not delete\n"
            foreign_file = install_dir / "foreign-user-data.txt"
            install_dir.mkdir(parents=True, exist_ok=False)
            foreign_file.write_bytes(foreign_payload)
            try:
                _assert_setup_rejected(
                    installer,
                    scratch / "reject-nonempty-default.log",
                    install_dir=None,
                    reason="fresh default install directory contained foreign data",
                )
                if (
                    not foreign_file.is_file()
                    or foreign_file.is_symlink()
                    or foreign_file.read_bytes() != foreign_payload
                ):
                    raise InstallerSmokeError(
                        "installer modified foreign data while rejecting a fresh install"
                    )
                unexpected = [path.name for path in install_dir.iterdir() if path != foreign_file]
                if unexpected or _owned_uninstall_entry() is not None:
                    raise InstallerSmokeError(
                        "rejected non-empty default target left installer-owned state "
                        f"(unexpected={unexpected!r})"
                    )
            finally:
                if (
                    foreign_file.is_file()
                    and not foreign_file.is_symlink()
                    and foreign_file.read_bytes() == foreign_payload
                ):
                    foreign_file.unlink()
                if install_dir.is_dir() and not install_dir.is_symlink():
                    with suppress(OSError):
                        install_dir.rmdir()

            _install_setup(
                baseline_installer,
                scratch / "install-previous.log",
                install_dir=None,
            )
            _assert_payload_matches(
                install_dir,
                baseline_hashes,
                phase="previous installation",
            )
            marker = install_dir / OWNERSHIP_MARKER_NAME
            if (
                not marker.is_file()
                or marker.is_symlink()
                or marker.read_text(encoding="ascii") != OWNERSHIP_MARKER_CONTENT
            ):
                raise InstallerSmokeError(
                    "initial installation did not create the expected ownership marker"
                )
            registered_entry = _owned_uninstall_entry()
            if registered_entry is None or not _same_path(
                registered_entry.get("InstallLocation"),
                install_dir,
            ):
                raise InstallerSmokeError(
                    "initial installation did not register the expected install directory"
                )

            redirect_error: BaseException | None = None
            try:
                _assert_setup_rejected(
                    installer,
                    scratch / "reject-registered-redirect.log",
                    install_dir=unsafe_install_dir,
                    reason="registered installation was redirected to a different /DIR",
                )
                if unsafe_install_dir.exists() or unsafe_install_dir.is_symlink():
                    raise InstallerSmokeError(
                        "rejected registered-install redirect created its target directory"
                    )
                registered_entry = _owned_uninstall_entry()
                if registered_entry is None or not _same_path(
                    registered_entry.get("InstallLocation"),
                    install_dir,
                ):
                    raise InstallerSmokeError(
                        "rejected redirect changed the registered install directory"
                    )
                _assert_payload_matches(
                    install_dir,
                    baseline_hashes,
                    phase="post-redirect-rejection installation",
                )
            except BaseException as exc:
                redirect_error = exc
                raise
            finally:
                if redirect_error is not None:
                    redirect_cleanup_errors = _cleanup_owned_state(
                        install_dir=unsafe_install_dir,
                        scratch=scratch,
                        group=group,
                        desktop_link=desktop_link,
                    )
                    _record_cleanup_errors(
                        redirect_error,
                        redirect_cleanup_errors,
                    )

            for relative in UPGRADE_STALE_MEMBERS:
                stale = install_dir / Path(relative)
                stale.parent.mkdir(parents=True, exist_ok=True)
                stale.write_bytes(b"stale payload from previous WorkbookLens version\n")

            _install_setup(
                installer,
                scratch / "install-current.log",
                install_dir=install_dir,
            )
            executable = install_dir / "WorkbookLens.exe"
            uninstaller = install_dir / "unins000.exe"
            if not executable.is_file() or not uninstaller.is_file():
                raise InstallerSmokeError(
                    "installer did not create the application and uninstaller"
                )
            _assert_payload_matches(
                install_dir,
                expected_hashes,
                phase="upgraded installation",
            )
            if marker.read_text(encoding="ascii") != OWNERSHIP_MARKER_CONTENT:
                raise InstallerSmokeError("upgrade changed the ownership marker unexpectedly")

            app_link = group / "WorkbookLens.lnk"
            uninstall_link = group / "Uninstall WorkbookLens.lnk"
            for link in (app_link, uninstall_link, desktop_link):
                if not link.is_file():
                    raise InstallerSmokeError(f"installer did not create shortcut: {link}")
            for link in (app_link, desktop_link):
                details = _shortcut_details(link)
                if Path(details.get("TargetPath", "")).resolve() != executable.resolve():
                    raise InstallerSmokeError(f"shortcut target is incorrect: {link}")
                if details.get("Arguments") != "serve --open-browser --fallback-port":
                    raise InstallerSmokeError(f"shortcut launch arguments are incorrect: {link}")

            entries = _uninstall_entries()
            if len(entries) != 1:
                raise InstallerSmokeError(f"expected one Apps & Features entry, found {entries!r}")
            entry = entries[0]
            if entry.get("DisplayVersion") != installer_report.version:
                raise InstallerSmokeError("Apps & Features displays the wrong version")
            if Path(entry.get("InstallLocation", "")).resolve() != install_dir.resolve():
                raise InstallerSmokeError("Apps & Features records the wrong install directory")

            environment = os.environ.copy()
            for key in tuple(environment):
                if key.casefold() in {"pythonhome", "pythonpath", "virtual_env"}:
                    environment.pop(key, None)
            completed = subprocess.run(  # noqa: S603 - exact installed release executable.
                [str(executable), "--version"],
                check=True,
                timeout=30,
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                env=environment,
            )
            if installer_report.version not in completed.stdout:
                raise InstallerSmokeError("installed executable reports the wrong version")
            _assert_payload_matches(
                install_dir,
                expected_hashes,
                phase="post-launch installation",
            )
            if _file_hashes(portable_root) != expected_hashes:
                raise InstallerSmokeError("installer smoke test modified the portable source files")

            _run(
                [str(uninstaller), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"],
                timeout=120.0,
            )
            _wait_removed(install_dir)
            if group.exists() or desktop_link.exists() or _uninstall_entries():
                raise InstallerSmokeError("uninstall left shortcuts or an Apps & Features entry")
        except BaseException as exc:
            primary_error = exc
            raise
        finally:
            cleanup_errors = _cleanup_owned_state(
                install_dir=install_dir,
                scratch=scratch,
                group=group,
                desktop_link=desktop_link,
            )
            _record_cleanup_errors(primary_error, cleanup_errors)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("installer", type=Path)
    parser.add_argument("--portable-zip", type=Path, required=True)
    parser.add_argument("--expected-version")
    parser.add_argument("--previous-installer", type=Path)
    parser.add_argument("--previous-portable-zip", type=Path)
    parser.add_argument(
        "--repository-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    args = parser.parse_args()
    try:
        smoke_installer(
            args.installer.resolve(),
            args.portable_zip.resolve(),
            expected_version=args.expected_version,
            repository_root=args.repository_root.resolve(),
            previous_installer=(
                args.previous_installer.resolve() if args.previous_installer else None
            ),
            previous_portable_zip=(
                args.previous_portable_zip.resolve() if args.previous_portable_zip else None
            ),
        )
    except (
        InstallerArtifactError,
        InstallerSmokeError,
        PortableArtifactError,
        OSError,
    ) as exc:
        print(f"installer smoke test failed: {exc}", file=sys.stderr)
        return 1
    print(
        "installer smoke test passed: forged registry, custom paths, and root/nested "
        "reparse points rejected, ownership marker, install, stale-payload upgrade "
        "cleanup, shortcuts, local executable, and uninstall"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
