from __future__ import annotations

import hashlib
import stat
import struct
import zipfile
from pathlib import Path

import pytest
from scripts.check_portable_artifact import (
    LEGACY_V2_2_1_PROFILE,
    ArtifactLimits,
    PortableArtifactError,
    _normalized_member_parts,
    _validate_member_type,
    extract_checked_artifact,
    inspect_artifact,
)

VERSION = "2.3.0"
ROOT = f"WorkbookLens-{VERSION}-windows-x64"
LEGACY_VERSION = "2.2.1"
LEGACY_ROOT = f"WorkbookLens-{LEGACY_VERSION}-windows-x64"


def _fake_pe(
    *,
    machine: int = 0x8664,
    optional_magic: int = 0x20B,
    version: str = VERSION,
    subsystem: int = 3,
) -> bytes:
    data = bytearray(512)
    data[:2] = b"MZ"
    struct.pack_into("<I", data, 0x3C, 0x80)
    data[0x80:0x84] = b"PE\0\0"
    struct.pack_into("<H", data, 0x84, machine)
    struct.pack_into("<H", data, 0x98, optional_magic)
    struct.pack_into("<H", data, 0xDC, subsystem)
    data.extend(version.encode("utf-16le"))
    return bytes(data)


def _fake_managed_anycpu_pe(*, version: str = "runtime") -> bytes:
    data = bytearray(1024)
    data[:2] = b"MZ"
    struct.pack_into("<I", data, 0x3C, 0x80)
    data[0x80:0x84] = b"PE\0\0"
    struct.pack_into("<H", data, 0x84, 0x14C)
    struct.pack_into("<H", data, 0x86, 1)
    struct.pack_into("<H", data, 0x94, 0xE0)
    struct.pack_into("<H", data, 0x98, 0x10B)
    struct.pack_into("<II", data, 0x168, 0x2000, 0x48)
    struct.pack_into("<IIII", data, 0x180, 0x200, 0x2000, 0x200, 0x200)
    struct.pack_into("<I", data, 0x210, 0x1)
    data.extend(version.encode("utf-16le"))
    return bytes(data)


def _repository_templates(root: Path) -> Path:
    (root / "packaging" / "windows").mkdir(parents=True)
    (root / "LICENSE").write_bytes(b"project license\n")
    (root / "workbooklens.example.yml").write_bytes(b"rules: []\n")
    (root / "packaging" / "windows" / "Start-WorkbookLens.cmd").write_bytes(b"@echo off\r\n")
    (root / "packaging" / "windows" / "README-PORTABLE.txt").write_text(
        "WorkbookLens @VERSION@ - Windows x64 Portable\n",
        encoding="utf-8",
    )
    return root


def _base_members(repository: Path) -> dict[str, bytes]:
    return {
        f"{ROOT}/WorkbookLens.exe": _fake_pe(subsystem=2),
        f"{ROOT}/WorkbookLensCLI.exe": _fake_pe(),
        f"{ROOT}/Start-WorkbookLens.cmd": (
            repository / "packaging" / "windows" / "Start-WorkbookLens.cmd"
        ).read_bytes(),
        f"{ROOT}/README-PORTABLE.txt": (
            f"WorkbookLens {VERSION} - Windows x64 Portable\n"
        ).encode(),
        f"{ROOT}/LICENSE": (repository / "LICENSE").read_bytes(),
        f"{ROOT}/THIRD-PARTY-NOTICES.txt": (
            b"CPython, PyInstaller, pywebview, and Lucide\n"
            b"LICENSES/sample-1.0/LICENSE.txt\n"
            b"LICENSES/Lucide-1.8.0/LICENSE.txt\n"
        ),
        f"{ROOT}/workbooklens.example.yml": (repository / "workbooklens.example.yml").read_bytes(),
        f"{ROOT}/LICENSES/CPython-3.12-LICENSE.txt": b"PSF license\n",
        f"{ROOT}/LICENSES/sample-1.0/LICENSE.txt": b"sample license\n",
        f"{ROOT}/LICENSES/Lucide-1.8.0/LICENSE.txt": b"Lucide ISC and MIT\n",
        f"{ROOT}/_internal/base_library.zip": b"runtime library\n",
        f"{ROOT}/_internal/python312.dll": _fake_pe(version="runtime"),
        f"{ROOT}/_internal/webview/js/api.js": b"webview api\n",
        f"{ROOT}/_internal/webview/lib/Microsoft.Web.WebView2.Core.dll": (
            _fake_managed_anycpu_pe()
        ),
        f"{ROOT}/_internal/webview/lib/runtimes/win-x64/native/WebView2Loader.dll": (
            _fake_pe(version="runtime")
        ),
        f"{ROOT}/_internal/webview/lib/runtimes/win-arm64/native/WebView2Loader.dll": (
            _fake_pe(machine=0xAA64, version="runtime")
        ),
        f"{ROOT}/_internal/webview/lib/runtimes/win-x86/native/WebView2Loader.dll": (
            _fake_pe(machine=0x14C, optional_magic=0x10B, version="runtime")
        ),
        f"{ROOT}/_internal/workbooklens/diff/templates/diff.html.j2": b"diff\n",
        f"{ROOT}/_internal/workbooklens/reports/templates/scan.html.j2": b"scan\n",
    }


def _legacy_members(repository: Path) -> dict[str, bytes]:
    return {
        f"{LEGACY_ROOT}/WorkbookLens.exe": _fake_pe(
            version=LEGACY_VERSION,
            subsystem=3,
        ),
        f"{LEGACY_ROOT}/Start-WorkbookLens.cmd": (
            repository / "packaging" / "windows" / "Start-WorkbookLens.cmd"
        ).read_bytes(),
        f"{LEGACY_ROOT}/README-PORTABLE.txt": (
            f"WorkbookLens {LEGACY_VERSION} - Windows x64 Portable\n"
        ).encode(),
        f"{LEGACY_ROOT}/LICENSE": (repository / "LICENSE").read_bytes(),
        f"{LEGACY_ROOT}/THIRD-PARTY-NOTICES.txt": (
            b"CPython and PyInstaller\nLICENSES/sample-1.0/LICENSE.txt\n"
        ),
        f"{LEGACY_ROOT}/workbooklens.example.yml": (
            repository / "workbooklens.example.yml"
        ).read_bytes(),
        f"{LEGACY_ROOT}/LICENSES/CPython-3.12-LICENSE.txt": b"PSF license\n",
        f"{LEGACY_ROOT}/LICENSES/sample-1.0/LICENSE.txt": b"sample license\n",
        f"{LEGACY_ROOT}/_internal/base_library.zip": b"runtime library\n",
        f"{LEGACY_ROOT}/_internal/python312.dll": _fake_pe(version="runtime"),
        f"{LEGACY_ROOT}/_internal/workbooklens/diff/templates/diff.html.j2": b"diff\n",
        f"{LEGACY_ROOT}/_internal/workbooklens/reports/templates/scan.html.j2": b"scan\n",
    }


def _write_artifact(
    tmp_path: Path,
    repository: Path,
    *,
    members: dict[str, bytes] | None = None,
    special: list[tuple[zipfile.ZipInfo, bytes]] | None = None,
    compression: int = zipfile.ZIP_DEFLATED,
) -> Path:
    archive = tmp_path / f"WorkbookLens-{VERSION}-windows-x64-portable.zip"
    payloads = _base_members(repository) if members is None else members
    with zipfile.ZipFile(archive, "w", compression=compression) as output:
        for name, data in payloads.items():
            output.writestr(name, data)
        for info, data in special or []:
            output.writestr(info, data)
    _refresh_checksum(archive)
    return archive


def _write_legacy_artifact(
    tmp_path: Path,
    repository: Path,
    *,
    members: dict[str, bytes] | None = None,
) -> Path:
    archive = tmp_path / f"WorkbookLens-{LEGACY_VERSION}-windows-x64-portable.zip"
    payloads = _legacy_members(repository) if members is None else members
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as output:
        for name, data in payloads.items():
            output.writestr(name, data)
    _refresh_checksum(archive)
    return archive


def _refresh_checksum(archive: Path) -> None:
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    archive.with_suffix(".zip.sha256").write_text(
        f"{digest}  {archive.name}\n",
        encoding="ascii",
    )


def test_inspect_and_extract_valid_artifact(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_artifact(tmp_path, repository)

    report = inspect_artifact(archive, expected_version=VERSION, repository_root=repository)
    extracted = extract_checked_artifact(
        archive,
        tmp_path / "extracted",
        expected_version=VERSION,
        repository_root=repository,
    )

    assert report.version == VERSION
    assert (extracted / "WorkbookLens.exe").read_bytes() == _fake_pe(subsystem=2)
    assert (extracted / "WorkbookLensCLI.exe").read_bytes() == _fake_pe()


def test_inspect_and_extract_valid_v2_2_1_legacy_artifact(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_legacy_artifact(tmp_path, repository)

    report = inspect_artifact(
        archive,
        expected_version=LEGACY_VERSION,
        profile=LEGACY_V2_2_1_PROFILE,
    )
    extracted = extract_checked_artifact(
        archive,
        tmp_path / "legacy-extracted",
        expected_version=LEGACY_VERSION,
        profile=LEGACY_V2_2_1_PROFILE,
    )

    assert report.version == LEGACY_VERSION
    assert (extracted / "WorkbookLens.exe").read_bytes() == _fake_pe(
        version=LEGACY_VERSION,
        subsystem=3,
    )
    assert not (extracted / "WorkbookLensCLI.exe").exists()


def test_current_profile_rejects_v2_2_1_legacy_contract(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_legacy_artifact(tmp_path, repository)

    with pytest.raises(PortableArtifactError, match="WorkbookLensCLI"):
        inspect_artifact(archive, expected_version=LEGACY_VERSION)


def test_legacy_profile_rejects_extra_cli_and_wrong_subsystem(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    with_cli = _legacy_members(repository)
    with_cli[f"{LEGACY_ROOT}/WorkbookLensCLI.exe"] = _fake_pe(
        version=LEGACY_VERSION,
        subsystem=3,
    )
    cli_archive = _write_legacy_artifact(tmp_path, repository, members=with_cli)

    with pytest.raises(PortableArtifactError, match="unexpected file at archive root"):
        inspect_artifact(
            cli_archive,
            expected_version=LEGACY_VERSION,
            profile=LEGACY_V2_2_1_PROFILE,
        )

    wrong_subsystem_dir = tmp_path / "wrong-subsystem"
    wrong_subsystem_dir.mkdir()
    wrong_subsystem = _legacy_members(repository)
    wrong_subsystem[f"{LEGACY_ROOT}/WorkbookLens.exe"] = _fake_pe(
        version=LEGACY_VERSION,
        subsystem=2,
    )
    subsystem_archive = _write_legacy_artifact(
        wrong_subsystem_dir,
        repository,
        members=wrong_subsystem,
    )

    with pytest.raises(PortableArtifactError, match="subsystem"):
        inspect_artifact(
            subsystem_archive,
            expected_version=LEGACY_VERSION,
            profile=LEGACY_V2_2_1_PROFILE,
        )


def test_legacy_profile_is_bound_to_v2_2_1_and_rejects_new_notices(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    current_archive = _write_artifact(tmp_path, repository)
    legacy_archive = _write_legacy_artifact(tmp_path, repository)

    with pytest.raises(PortableArtifactError, match="requires explicit expected_version"):
        inspect_artifact(
            legacy_archive,
            profile=LEGACY_V2_2_1_PROFILE,
        )

    with pytest.raises(PortableArtifactError, match="archive version"):
        inspect_artifact(
            current_archive,
            expected_version=LEGACY_VERSION,
            profile=LEGACY_V2_2_1_PROFILE,
        )

    legacy_dir = tmp_path / "new-notices"
    legacy_dir.mkdir()
    members = _legacy_members(repository)
    members[f"{LEGACY_ROOT}/THIRD-PARTY-NOTICES.txt"] += b"pywebview\n"
    new_notice_archive = _write_legacy_artifact(legacy_dir, repository, members=members)

    with pytest.raises(PortableArtifactError, match="unexpectedly mentions pywebview"):
        inspect_artifact(
            new_notice_archive,
            expected_version=LEGACY_VERSION,
            profile=LEGACY_V2_2_1_PROFILE,
        )


def test_legacy_profile_rejects_webview_runtime_and_new_license_directory(
    tmp_path: Path,
) -> None:
    repository = _repository_templates(tmp_path / "repo")
    runtime_members = _legacy_members(repository)
    runtime_members[f"{LEGACY_ROOT}/_internal/webview/js/api.js"] = b"unexpected webview"
    runtime_archive = _write_legacy_artifact(tmp_path, repository, members=runtime_members)

    with pytest.raises(PortableArtifactError, match="forbids runtime path"):
        inspect_artifact(
            runtime_archive,
            expected_version=LEGACY_VERSION,
            profile=LEGACY_V2_2_1_PROFILE,
        )

    license_dir = tmp_path / "new-license"
    license_dir.mkdir()
    license_members = _legacy_members(repository)
    license_members[f"{LEGACY_ROOT}/LICENSES/Lucide-1.8.0/LICENSE.txt"] = b"Lucide license"
    license_archive = _write_legacy_artifact(
        license_dir,
        repository,
        members=license_members,
    )

    with pytest.raises(PortableArtifactError, match="forbids license directory"):
        inspect_artifact(
            license_archive,
            expected_version=LEGACY_VERSION,
            profile=LEGACY_V2_2_1_PROFILE,
        )


def test_allows_windows_api_debug_runtime_dll(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/_internal/api-ms-win-core-debug-l1-1-0.dll"] = _fake_pe(version="runtime")
    archive = _write_artifact(tmp_path, repository, members=members)

    inspect_artifact(archive, expected_version=VERSION, repository_root=repository)


def test_allows_pythonnet_system_debug_runtime_dll(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/_internal/pythonnet/runtime/System.Diagnostics.Debug.dll"] = (
        _fake_managed_anycpu_pe()
    )
    archive = _write_artifact(tmp_path, repository, members=members)

    inspect_artifact(archive, expected_version=VERSION, repository_root=repository)


def test_rejects_other_sensitive_internal_filename(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/_internal/debug-secrets.txt"] = b"not for release"
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="sensitive filename token"):
        inspect_artifact(archive, expected_version=VERSION, repository_root=repository)


@pytest.mark.parametrize(
    "unsafe_name",
    [
        f"{ROOT}/../escape.txt",
        "/absolute.txt",
        "C:/absolute.txt",
        f"{ROOT}/_internal/file.txt:secret",
        f"{ROOT}/_internal/CON.txt",
        f"{ROOT}/_internal/.env",
        f"{ROOT}/_internal/COM¹.txt",
        f"{ROOT}/_internal/bad?.txt",
    ],
)
def test_rejects_unsafe_member_names(tmp_path: Path, unsafe_name: str) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[unsafe_name] = b"bad"
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_ambiguous_or_normalized_zipinfo_name() -> None:
    info = zipfile.ZipInfo(f"{ROOT}/_internal/runtime.dll")
    info.filename = f"{ROOT}\\_internal\\runtime.dll"
    info.orig_filename = info.filename
    with pytest.raises(PortableArtifactError, match="backslash"):
        _normalized_member_parts(info)

    info = zipfile.ZipInfo(f"{ROOT}/_internal/runtime.dll")
    info.orig_filename = f"{ROOT}/_internal/runtime.dll\x00hidden"
    with pytest.raises(PortableArtifactError, match="normalized or truncated"):
        _normalized_member_parts(info)


def test_rejects_case_colliding_member(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/license"] = b"collision"
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="colliding"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_symlink_member(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    info = zipfile.ZipInfo(f"{ROOT}/_internal/link")
    info.create_system = 3
    info.external_attr = (stat.S_IFLNK | 0o777) << 16
    archive = _write_artifact(tmp_path, repository, special=[(info, b"../../outside")])

    with pytest.raises(PortableArtifactError, match="symbolic link"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_encrypted_member_flag() -> None:
    info = zipfile.ZipInfo(f"{ROOT}/_internal/encrypted")
    info.flag_bits |= 0x1

    with pytest.raises(PortableArtifactError, match="encrypted"):
        _validate_member_type(info)


def test_rejects_excessive_compression_ratio(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_artifact(tmp_path, repository)

    with pytest.raises(PortableArtifactError, match="compression ratio"):
        inspect_artifact(
            archive,
            expected_version=VERSION,
            limits=ArtifactLimits(max_compression_ratio=1.0),
        )


def test_rejects_non_x64_pe(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/WorkbookLens.exe"] = _fake_pe(machine=0x14C, subsystem=2)
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="x64"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_console_subsystem_for_desktop_executable(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/WorkbookLens.exe"] = _fake_pe(subsystem=3)
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="subsystem"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_windowed_subsystem_for_cli_executable(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/WorkbookLensCLI.exe"] = _fake_pe(subsystem=2)
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="subsystem"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_bad_checksum(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_artifact(tmp_path, repository)
    archive.with_suffix(".zip.sha256").write_text(f"{'0' * 64}  {archive.name}\n", encoding="ascii")

    with pytest.raises(PortableArtifactError, match="does not match"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_file_directory_ancestor_conflict(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/_internal/conflict"] = b"file"
    members[f"{ROOT}/_internal/conflict/child.txt"] = b"child"
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="ancestor conflict"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_executable_under_licenses(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/LICENSES/sample-1.0/payload.exe"] = b"MZ"
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="forbidden under LICENSES"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_workbook_under_internal(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/_internal/payload.xlsx"] = b"workbook"
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="under _internal"):
        inspect_artifact(archive, expected_version=VERSION)


@pytest.mark.parametrize(
    "missing_name",
    [
        f"{ROOT}/_internal/base_library.zip",
        f"{ROOT}/_internal/python312.dll",
        f"{ROOT}/_internal/webview/js/api.js",
        f"{ROOT}/_internal/webview/lib/Microsoft.Web.WebView2.Core.dll",
        f"{ROOT}/_internal/webview/lib/runtimes/win-arm64/native/WebView2Loader.dll",
        f"{ROOT}/_internal/webview/lib/runtimes/win-x64/native/WebView2Loader.dll",
        f"{ROOT}/_internal/webview/lib/runtimes/win-x86/native/WebView2Loader.dll",
        f"{ROOT}/_internal/workbooklens/diff/templates/diff.html.j2",
        f"{ROOT}/_internal/workbooklens/reports/templates/scan.html.j2",
    ],
)
def test_rejects_missing_required_runtime_file(
    tmp_path: Path,
    missing_name: str,
) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members.pop(missing_name)
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="required runtime"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_x86_runtime_dll(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/_internal/python312.dll"] = _fake_pe(
        machine=0x14C,
        version="runtime",
    )
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="x64"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_wrong_architecture_for_pywebview_loader(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    members = _base_members(repository)
    members[f"{ROOT}/_internal/webview/lib/runtimes/win-x86/native/WebView2Loader.dll"] = _fake_pe(
        version="runtime"
    )
    archive = _write_artifact(tmp_path, repository, members=members)

    with pytest.raises(PortableArtifactError, match="required PE architecture"):
        inspect_artifact(archive, expected_version=VERSION)


def test_allows_managed_anycpu_runtime_dll(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_artifact(tmp_path, repository)

    inspect_artifact(archive, expected_version=VERSION)


def test_rejects_zip_preamble_and_trailing_bytes(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    preamble = _write_artifact(tmp_path, repository)
    preamble.write_bytes(b"preamble" + preamble.read_bytes())
    _refresh_checksum(preamble)
    with pytest.raises(PortableArtifactError, match="preamble"):
        inspect_artifact(preamble, expected_version=VERSION)

    trailing_dir = tmp_path / "trailing"
    trailing_dir.mkdir()
    trailing = _write_artifact(trailing_dir, repository)
    trailing.write_bytes(trailing.read_bytes() + b"trailing")
    _refresh_checksum(trailing)
    with pytest.raises(PortableArtifactError, match="exactly at EOF"):
        inspect_artifact(trailing, expected_version=VERSION)


def test_rejects_hidden_bytes_before_central_directory(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_artifact(tmp_path, repository)
    payload = bytearray(archive.read_bytes())
    eocd_offset = len(payload) - 22
    directory_offset = struct.unpack_from("<L", payload, eocd_offset + 16)[0]
    hidden = b"EMBEDDED-SECRET"
    mutated = payload[:directory_offset] + hidden + payload[directory_offset:]
    struct.pack_into(
        "<L",
        mutated,
        eocd_offset + len(hidden) + 16,
        directory_offset + len(hidden),
    )
    archive.write_bytes(mutated)
    _refresh_checksum(archive)

    with pytest.raises(PortableArtifactError, match="hidden data between local ZIP records"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_bad_member_crc(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_artifact(
        tmp_path,
        repository,
        compression=zipfile.ZIP_STORED,
    )
    member_name = f"{ROOT}/LICENSES/sample-1.0/LICENSE.txt"
    with zipfile.ZipFile(archive) as source:
        info = source.getinfo(member_name)
    data = bytearray(archive.read_bytes())
    name_length, extra_length = struct.unpack_from(
        "<HH",
        data,
        info.header_offset + 26,
    )
    payload_offset = info.header_offset + 30 + name_length + extra_length
    data[payload_offset] ^= 0x01
    archive.write_bytes(data)
    _refresh_checksum(archive)

    with pytest.raises(PortableArtifactError, match=r"cannot read|CRC|payload metadata"):
        inspect_artifact(archive, expected_version=VERSION)


def test_rejects_nonempty_extraction_destination(tmp_path: Path) -> None:
    repository = _repository_templates(tmp_path / "repo")
    archive = _write_artifact(tmp_path, repository)
    destination = tmp_path / "occupied"
    destination.mkdir()
    marker = destination / "keep.txt"
    marker.write_text("keep", encoding="utf-8")

    with pytest.raises(PortableArtifactError, match="must be empty"):
        extract_checked_artifact(
            archive,
            destination,
            expected_version=VERSION,
            repository_root=repository,
        )
    assert marker.read_text(encoding="utf-8") == "keep"
