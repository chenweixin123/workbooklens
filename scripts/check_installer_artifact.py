"""Validate the official WorkbookLens Windows installer and checksum sidecar."""

from __future__ import annotations

import argparse
import ctypes
import hashlib
import os
import re
import struct
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Final, NoReturn

INSTALLER_RE: Final = re.compile(
    r"^WorkbookLens-(?P<version>[A-Za-z0-9][A-Za-z0-9._+-]*)-windows-x64-setup\.exe$"
)
ALLOWED_PE_MACHINES: Final = {0x014C: "x86 bootstrap", 0x8664: "x64"}


class InstallerArtifactError(RuntimeError):
    """The setup executable or its sidecar violates the release contract."""


@dataclass(frozen=True)
class InstallerLimits:
    min_size: int = 1 * 1024 * 1024
    max_size: int = 256 * 1024 * 1024


DEFAULT_LIMITS: Final = InstallerLimits()


@dataclass(frozen=True)
class InstallerReport:
    installer: Path
    version: str
    size: int
    sha256: str
    machine: str


def _fail(message: str) -> NoReturn:
    raise InstallerArtifactError(message)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _check_sidecar(installer: Path, digest: str) -> None:
    sidecar = installer.with_suffix(installer.suffix + ".sha256")
    if not sidecar.is_file() or sidecar.is_symlink():
        _fail(f"installer SHA256 sidecar is missing or unsafe: {sidecar}")
    try:
        raw = sidecar.read_bytes()
        text = raw.decode("ascii")
    except (OSError, UnicodeDecodeError) as exc:
        _fail(f"cannot read installer SHA256 sidecar: {exc}")
    expected = f"{digest}  {installer.name}\n"
    if text != expected:
        _fail("installer SHA256 sidecar does not match the setup executable")


def _read_pe_machine(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            dos = stream.read(64)
            if len(dos) != 64 or dos[:2] != b"MZ":
                _fail("installer is not a Windows PE executable")
            pe_offset = struct.unpack_from("<I", dos, 0x3C)[0]
            if pe_offset < 64 or pe_offset > 16 * 1024 * 1024:
                _fail("installer has an invalid PE header offset")
            stream.seek(pe_offset)
            header = stream.read(26)
    except OSError as exc:
        _fail(f"cannot read installer PE headers: {exc}")
    if len(header) != 26 or header[:4] != b"PE\0\0":
        _fail("installer has an invalid PE signature")
    machine_code = struct.unpack_from("<H", header, 4)[0]
    optional_magic = struct.unpack_from("<H", header, 24)[0]
    if machine_code not in ALLOWED_PE_MACHINES or optional_magic not in {0x10B, 0x20B}:
        _fail("installer does not contain a supported Windows setup bootstrap")
    return ALLOWED_PE_MACHINES[machine_code]


def _windows_version_strings(path: Path) -> dict[str, str]:
    version = ctypes.WinDLL("version", use_last_error=True)
    handle = ctypes.c_uint32()
    size = version.GetFileVersionInfoSizeW(str(path), ctypes.byref(handle))
    if not size:
        _fail("installer has no readable Windows version resource")
    buffer = ctypes.create_string_buffer(size)
    if not version.GetFileVersionInfoW(str(path), 0, size, buffer):
        _fail("cannot read the installer Windows version resource")

    value_pointer = ctypes.c_void_p()
    value_length = ctypes.c_uint()
    translations: list[tuple[int, int]] = []
    if version.VerQueryValueW(
        buffer,
        r"\VarFileInfo\Translation",
        ctypes.byref(value_pointer),
        ctypes.byref(value_length),
    ):
        values = ctypes.cast(value_pointer, ctypes.POINTER(ctypes.c_ushort))
        translations = [
            (values[index], values[index + 1]) for index in range(0, value_length.value // 2 - 1, 2)
        ]
    if not translations:
        translations = [(0x0409, 0x04B0)]

    result: dict[str, str] = {}
    for key in ("ProductName", "ProductVersion"):
        for language, code_page in translations:
            query = rf"\StringFileInfo\{language:04x}{code_page:04x}\{key}"
            if not version.VerQueryValueW(
                buffer,
                query,
                ctypes.byref(value_pointer),
                ctypes.byref(value_length),
            ):
                continue
            result[key] = ctypes.wstring_at(value_pointer, value_length.value).rstrip("\0 ")
            break
    return result


def inspect_installer(
    installer: Path,
    *,
    expected_version: str | None = None,
    limits: InstallerLimits = DEFAULT_LIMITS,
    verify_windows_version_resource: bool = True,
) -> InstallerReport:
    source = installer.expanduser().absolute()
    if source.is_symlink():
        _fail(f"installer path must not be a symbolic link: {source}")
    try:
        installer = source.resolve(strict=True)
    except OSError as exc:
        _fail(f"cannot resolve installer path: {exc}")
    if not installer.is_file():
        _fail(f"installer does not exist or is not a regular file: {installer}")
    match = INSTALLER_RE.fullmatch(installer.name)
    if match is None:
        _fail(f"unexpected installer filename: {installer.name!r}")
    version = match.group("version")
    if expected_version is not None and version != expected_version:
        _fail(f"installer version {version!r} does not match {expected_version!r}")
    size = installer.stat().st_size
    if not limits.min_size <= size <= limits.max_size:
        _fail(
            f"installer size {size} is outside the allowed range "
            f"{limits.min_size}..{limits.max_size}"
        )
    machine = _read_pe_machine(installer)
    digest = _sha256(installer)
    _check_sidecar(installer, digest)
    if verify_windows_version_resource and os.name == "nt":
        strings = _windows_version_strings(installer)
        if strings.get("ProductName") != "WorkbookLens":
            _fail(f"unexpected installer ProductName: {strings.get('ProductName')!r}")
        if strings.get("ProductVersion") != version:
            _fail(f"unexpected installer ProductVersion: {strings.get('ProductVersion')!r}")
    return InstallerReport(
        installer=installer,
        version=version,
        size=size,
        sha256=digest,
        machine=machine,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("installer", type=Path)
    parser.add_argument("--expected-version")
    args = parser.parse_args()
    try:
        report = inspect_installer(args.installer, expected_version=args.expected_version)
    except (InstallerArtifactError, OSError) as exc:
        print(f"installer artifact check failed: {exc}", file=sys.stderr)
        return 1
    print(
        f"installer artifact OK: {report.installer} "
        f"({report.size} bytes, {report.machine}, sha256={report.sha256})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
