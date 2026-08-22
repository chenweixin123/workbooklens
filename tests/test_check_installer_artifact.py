from __future__ import annotations

import hashlib
import struct
from pathlib import Path

import pytest
from scripts.check_installer_artifact import (
    InstallerArtifactError,
    InstallerLimits,
    inspect_installer,
)

VERSION = "2.2.1"


def _fake_installer(*, machine: int = 0x14C) -> bytes:
    payload = bytearray(4096)
    payload[:2] = b"MZ"
    struct.pack_into("<I", payload, 0x3C, 0x80)
    payload[0x80:0x84] = b"PE\0\0"
    struct.pack_into("<H", payload, 0x84, machine)
    struct.pack_into("<H", payload, 0x98, 0x10B if machine == 0x14C else 0x20B)
    return bytes(payload)


def _write_installer(tmp_path: Path, *, payload: bytes | None = None) -> Path:
    installer = tmp_path / f"WorkbookLens-{VERSION}-windows-x64-setup.exe"
    installer.write_bytes(_fake_installer() if payload is None else payload)
    digest = hashlib.sha256(installer.read_bytes()).hexdigest()
    installer.with_suffix(".exe.sha256").write_text(
        f"{digest}  {installer.name}\n",
        encoding="ascii",
        newline="\n",
    )
    return installer


def test_inspects_valid_inno_bootstrap_shape(tmp_path: Path) -> None:
    installer = _write_installer(tmp_path)

    report = inspect_installer(
        installer,
        expected_version=VERSION,
        limits=InstallerLimits(min_size=512),
        verify_windows_version_resource=False,
    )

    assert report.version == VERSION
    assert report.machine == "x86 bootstrap"


def test_accepts_x64_setup_bootstrap(tmp_path: Path) -> None:
    installer = _write_installer(tmp_path, payload=_fake_installer(machine=0x8664))

    report = inspect_installer(
        installer,
        limits=InstallerLimits(min_size=512),
        verify_windows_version_resource=False,
    )

    assert report.machine == "x64"


def test_rejects_bad_installer_checksum(tmp_path: Path) -> None:
    installer = _write_installer(tmp_path)
    installer.with_suffix(".exe.sha256").write_text(
        f"{'0' * 64}  {installer.name}\n",
        encoding="ascii",
        newline="\n",
    )

    with pytest.raises(InstallerArtifactError, match="does not match"):
        inspect_installer(
            installer,
            limits=InstallerLimits(min_size=512),
            verify_windows_version_resource=False,
        )


def test_rejects_non_pe_installer(tmp_path: Path) -> None:
    installer = _write_installer(tmp_path, payload=b"not a setup executable" * 100)

    with pytest.raises(InstallerArtifactError, match="Windows PE"):
        inspect_installer(
            installer,
            limits=InstallerLimits(min_size=512),
            verify_windows_version_resource=False,
        )


def test_rejects_unexpected_installer_name(tmp_path: Path) -> None:
    installer = _write_installer(tmp_path).rename(tmp_path / "setup.exe")

    with pytest.raises(InstallerArtifactError, match="filename"):
        inspect_installer(
            installer,
            limits=InstallerLimits(min_size=512),
            verify_windows_version_resource=False,
        )


def test_rejects_installer_symlink_before_resolving(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    installer = _write_installer(tmp_path)
    original = Path.is_symlink

    def fake_is_symlink(path: Path) -> bool:
        return path == installer.absolute() or original(path)

    monkeypatch.setattr(Path, "is_symlink", fake_is_symlink)
    with pytest.raises(InstallerArtifactError, match="symbolic link"):
        inspect_installer(
            installer,
            limits=InstallerLimits(min_size=512),
            verify_windows_version_resource=False,
        )
