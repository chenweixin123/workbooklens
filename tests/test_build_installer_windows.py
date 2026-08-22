from __future__ import annotations

from pathlib import Path

import pytest
from scripts.build_installer_windows import (
    InstallerBuildError,
    _publish_validated_pair,
    numeric_installer_version,
    resolve_iscc,
)


def test_numeric_installer_version() -> None:
    assert numeric_installer_version("2.2.1") == "2.2.1.0"
    assert numeric_installer_version("2.2.1rc1") == "2.2.1.0"


@pytest.mark.parametrize("version", ("dev", "1.2.70000"))
def test_rejects_unrepresentable_installer_version(version: str) -> None:
    with pytest.raises(InstallerBuildError):
        numeric_installer_version(version)


def test_explicit_inno_compiler_must_exist(tmp_path: Path) -> None:
    compiler = tmp_path / "ISCC.exe"

    with pytest.raises(InstallerBuildError, match="does not exist"):
        resolve_iscc(compiler)

    compiler.write_bytes(b"MZ")
    assert resolve_iscc(compiler) == compiler.resolve()


def test_inno_definition_creates_named_shortcuts_and_uninstaller() -> None:
    definition = (
        Path(__file__).parents[1] / "packaging" / "windows" / "WorkbookLens.iss"
    ).read_text(encoding="utf-8")

    assert "PrivilegesRequired=lowest" in definition
    assert "DefaultDirName={localappdata}\\Programs\\{#AppName}" in definition
    assert 'Name: "{group}\\{#AppName}"' in definition
    assert 'Name: "{autodesktop}\\{#AppName}"' in definition
    assert "Tasks: desktopicon" in definition
    assert 'Name: "{group}\\Uninstall {#AppName}"' in definition
    assert 'Filename: "{uninstallexe}"' in definition
    assert 'Parameters: "serve --open-browser --fallback-port"' in definition
    assert "UninstallDisplayName={#AppName}" in definition
    assert "AppId={{7B7534E0-8485-4F4F-8DE7-561869FF7C0C}" in definition
    assert 'Type: filesandordirs; Name: "{app}\\_internal"' in definition
    assert 'Type: filesandordirs; Name: "{app}\\LICENSES"' in definition
    assert 'Type: filesandordirs; Name: "{app}"' not in definition


def test_publishes_validated_installer_pair(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    output = tmp_path / "output"
    staging.mkdir()
    output.mkdir()
    staged_installer = staging / "WorkbookLens-2.2.1-windows-x64-setup.exe"
    staged_sidecar = staging / f"{staged_installer.name}.sha256"
    installer = output / staged_installer.name
    sidecar = output / staged_sidecar.name
    staged_installer.write_bytes(b"new installer")
    staged_sidecar.write_bytes(b"new sidecar")

    _publish_validated_pair(
        staged_installer,
        staged_sidecar,
        installer,
        sidecar,
        overwrite=False,
        validate=lambda path: path.read_bytes(),
    )

    assert installer.read_bytes() == b"new installer"
    assert sidecar.read_bytes() == b"new sidecar"


def test_publish_failure_removes_new_outputs_and_preserves_error(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    output = tmp_path / "output"
    staging.mkdir()
    output.mkdir()
    staged_installer = staging / "WorkbookLens-2.2.1-windows-x64-setup.exe"
    staged_sidecar = staging / f"{staged_installer.name}.sha256"
    installer = output / staged_installer.name
    sidecar = output / staged_sidecar.name
    staged_installer.write_bytes(b"new installer")
    staged_sidecar.write_bytes(b"new sidecar")

    def reject(_path: Path) -> None:
        raise ValueError("post-publish validation failed")

    with pytest.raises(InstallerBuildError, match="post-publish validation failed"):
        _publish_validated_pair(
            staged_installer,
            staged_sidecar,
            installer,
            sidecar,
            overwrite=False,
            validate=reject,
        )

    assert not installer.exists()
    assert not sidecar.exists()


def test_overwrite_failure_restores_previous_pair(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    output = tmp_path / "output"
    staging.mkdir()
    output.mkdir()
    staged_installer = staging / "WorkbookLens-2.2.1-windows-x64-setup.exe"
    staged_sidecar = staging / f"{staged_installer.name}.sha256"
    installer = output / staged_installer.name
    sidecar = output / staged_sidecar.name
    staged_installer.write_bytes(b"new installer")
    staged_sidecar.write_bytes(b"new sidecar")
    installer.write_bytes(b"old installer")
    sidecar.write_bytes(b"old sidecar")

    def reject(_path: Path) -> None:
        raise ValueError("post-publish validation failed")

    with pytest.raises(InstallerBuildError, match="post-publish validation failed"):
        _publish_validated_pair(
            staged_installer,
            staged_sidecar,
            installer,
            sidecar,
            overwrite=True,
            validate=reject,
        )

    assert installer.read_bytes() == b"old installer"
    assert sidecar.read_bytes() == b"old sidecar"
