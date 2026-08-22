from __future__ import annotations

import base64
import gzip
import hashlib
import io
import stat
import struct
import tarfile
import zipfile
from pathlib import Path

import pytest
from scripts.check_release_artifacts import (
    ArtifactError,
    _check_tar_members,
    _check_wheel,
    _check_zip_members,
    check_distribution,
)

VERSION = "2.2.1"


def _write_valid_distributions(
    directory: Path,
    *,
    excluded_sdist_member: str | None = None,
) -> None:
    wheel = directory / f"workbooklens-{VERSION}-py3-none-any.whl"
    dist_info = f"workbooklens-{VERSION}.dist-info"
    wheel_members = {
        "workbooklens/__init__.py": b"safe",
        "workbooklens/console.py": b"safe",
        "workbooklens/diff/templates/diff.html.j2": b"safe",
        "workbooklens/reports/templates/scan.html.j2": b"safe",
        "workbooklens/web/launcher.py": b"safe",
        f"{dist_info}/METADATA": (
            f"Metadata-Version: 2.4\nName: workbooklens\nVersion: {VERSION}\n"
        ).encode(),
        f"{dist_info}/WHEEL": (
            b"Wheel-Version: 1.0\nGenerator: tests\nRoot-Is-Purelib: true\nTag: py3-none-any\n"
        ),
        f"{dist_info}/licenses/LICENSE": b"safe",
    }
    record_lines = []
    for name, payload in wheel_members.items():
        digest = base64.urlsafe_b64encode(hashlib.sha256(payload).digest()).rstrip(b"=")
        record_lines.append(f"{name},sha256={digest.decode()},{len(payload)}")
    record_name = f"{dist_info}/RECORD"
    record_lines.append(f"{record_name},,")
    wheel_members[record_name] = ("\n".join(record_lines) + "\n").encode()
    with zipfile.ZipFile(wheel, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in wheel_members.items():
            archive.writestr(name, payload)

    sdist = directory / f"workbooklens-{VERSION}.tar.gz"
    root = f"workbooklens-{VERSION}"
    required = (
        "LICENSE",
        "PKG-INFO",
        "README.md",
        "packaging/windows/README-PORTABLE.txt",
        "packaging/windows/Start-WorkbookLens.cmd",
        "packaging/windows/WorkbookLens.spec",
        "packaging/windows/entry.py",
        "pyproject.toml",
        "scripts/action_scan.py",
        "scripts/build_portable_windows.py",
        "scripts/check_portable_artifact.py",
        "scripts/check_release_artifacts.py",
        "scripts/smoke_portable.py",
        "src/workbooklens/__init__.py",
        "src/workbooklens/console.py",
        "src/workbooklens/web/launcher.py",
    )
    with tarfile.open(sdist, "w:gz") as archive:
        for name in required:
            if name == excluded_sdist_member:
                continue
            payload = (
                f"Metadata-Version: 2.4\nName: workbooklens\nVersion: {VERSION}\n".encode()
                if name == "PKG-INFO"
                else b"safe"
            )
            info = tarfile.TarInfo(f"{root}/{name}")
            info.size = len(payload)
            archive.addfile(info, io.BytesIO(payload))


def test_release_directory_rejects_unexpected_extra_file(tmp_path: Path) -> None:
    _write_valid_distributions(tmp_path)
    (tmp_path / "debug-secrets.txt").write_text("not for release", encoding="utf-8")

    with pytest.raises(ArtifactError, match="unexpected: debug-secrets"):
        check_distribution(tmp_path, VERSION)


def test_release_directory_accepts_exact_expected_files(tmp_path: Path) -> None:
    _write_valid_distributions(tmp_path)

    check_distribution(tmp_path, VERSION)


@pytest.mark.parametrize("content", (b"*", b"*\n", b"*\r\n"))
def test_release_directory_accepts_uv_generated_ignore(
    tmp_path: Path,
    content: bytes,
) -> None:
    _write_valid_distributions(tmp_path)
    (tmp_path / ".gitignore").write_bytes(content)

    check_distribution(tmp_path, VERSION)


def test_release_directory_rejects_modified_uv_generated_ignore(tmp_path: Path) -> None:
    _write_valid_distributions(tmp_path)
    (tmp_path / ".gitignore").write_bytes(b"*\nsecrets.env\n")

    with pytest.raises(ArtifactError, match="unexpected content"):
        check_distribution(tmp_path, VERSION)


def test_release_directory_rejects_uv_generated_ignore_directory(tmp_path: Path) -> None:
    _write_valid_distributions(tmp_path)
    (tmp_path / ".gitignore").mkdir()

    with pytest.raises(ArtifactError, match="linked, external, or non-file"):
        check_distribution(tmp_path, VERSION)


def test_release_directory_rejects_linked_uv_generated_ignore(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _write_valid_distributions(tmp_path)
    (tmp_path / ".gitignore").write_bytes(b"*\n")
    original = Path.is_symlink

    def fake_is_symlink(path: Path) -> bool:
        return path.name == ".gitignore" or original(path)

    monkeypatch.setattr(Path, "is_symlink", fake_is_symlink)
    with pytest.raises(ArtifactError, match="linked, external"):
        check_distribution(tmp_path, VERSION)


@pytest.mark.parametrize(
    "missing",
    (
        "packaging/windows/entry.py",
        "packaging/windows/Start-WorkbookLens.cmd",
        "packaging/windows/README-PORTABLE.txt",
        "src/workbooklens/console.py",
        "src/workbooklens/web/launcher.py",
    ),
)
def test_sdist_requires_portable_build_inputs(tmp_path: Path, missing: str) -> None:
    _write_valid_distributions(tmp_path, excluded_sdist_member=missing)

    with pytest.raises(ArtifactError) as error:
        check_distribution(tmp_path, VERSION)
    assert missing in str(error.value)


def test_release_directory_rejects_linked_expected_file(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _write_valid_distributions(tmp_path)
    wheel_name = f"workbooklens-{VERSION}-py3-none-any.whl"
    original = Path.is_symlink

    def fake_is_symlink(path: Path) -> bool:
        return path.name == wheel_name or original(path)

    monkeypatch.setattr(Path, "is_symlink", fake_is_symlink)
    with pytest.raises(ArtifactError, match="linked, external"):
        check_distribution(tmp_path, VERSION)


def test_zip_member_case_collision_is_rejected() -> None:
    first = zipfile.ZipInfo("workbooklens/Rules.py")
    second = zipfile.ZipInfo("workbooklens/rules.py")

    with pytest.raises(ArtifactError, match="case-colliding"):
        _check_zip_members([first, second])


@pytest.mark.parametrize(
    "name",
    (
        "workbooklens//module.py",
        "workbooklens/./module.py",
        "workbooklens/../module.py",
    ),
)
def test_raw_non_portable_path_components_are_rejected(name: str) -> None:
    with pytest.raises(ArtifactError, match="non-portable"):
        _check_zip_members([zipfile.ZipInfo(name)])


@pytest.mark.parametrize(
    "original",
    (
        "workbooklens\\module.py",
        "workbooklens/module.py\x00hidden",
    ),
)
def test_original_zip_name_is_checked_before_normalization(original: str) -> None:
    info = zipfile.ZipInfo("workbooklens/module.py")
    info.orig_filename = original

    with pytest.raises(ArtifactError, match="non-portable"):
        _check_zip_members([info])


@pytest.mark.parametrize(
    "name",
    (
        "workbooklens/bad?.py",
        "workbooklens/bad|name.py",
        "workbooklens/control\x1f.py",
        "workbooklens/COM¹.txt",
    ),
)
def test_windows_invalid_member_names_are_rejected(name: str) -> None:
    with pytest.raises(ArtifactError, match="Windows-unsafe"):
        _check_zip_members([zipfile.ZipInfo(name)])


def test_zip_symlink_member_is_rejected() -> None:
    info = zipfile.ZipInfo("workbooklens/link.py")
    info.create_system = 3
    info.external_attr = (stat.S_IFLNK | 0o777) << 16

    with pytest.raises(ArtifactError, match="symbolic-link"):
        _check_zip_members([info])


def test_zip_encrypted_member_is_rejected() -> None:
    info = zipfile.ZipInfo("workbooklens/secret.py")
    info.flag_bits = 0x1

    with pytest.raises(ArtifactError, match="encrypted"):
        _check_zip_members([info])


def test_zip_file_directory_ancestor_conflict_is_rejected() -> None:
    file_info = zipfile.ZipInfo("workbooklens/conflict")
    child_info = zipfile.ZipInfo("workbooklens/conflict/child.py")

    with pytest.raises(ArtifactError, match="also used as a directory"):
        _check_zip_members([file_info, child_info])


def test_tar_file_directory_ancestor_conflict_is_rejected() -> None:
    file_info = tarfile.TarInfo("workbooklens-2.2.1/conflict")
    child_info = tarfile.TarInfo("workbooklens-2.2.1/conflict/child.py")

    with pytest.raises(ArtifactError, match="also used as a directory"):
        _check_tar_members([file_info, child_info])


@pytest.mark.parametrize("position", ("before", "after"))
def test_wheel_rejects_hidden_bytes_outside_zip_records(
    tmp_path: Path,
    position: str,
) -> None:
    _write_valid_distributions(tmp_path)
    wheel = tmp_path / f"workbooklens-{VERSION}-py3-none-any.whl"
    payload = wheel.read_bytes()
    hidden = b"EMBEDDED-SECRET"
    wheel.write_bytes(hidden + payload if position == "before" else payload + hidden)

    with pytest.raises(ArtifactError, match=r"preamble|trailing"):
        _check_wheel(wheel, VERSION)


def test_wheel_rejects_hidden_bytes_before_central_directory(tmp_path: Path) -> None:
    _write_valid_distributions(tmp_path)
    wheel = tmp_path / f"workbooklens-{VERSION}-py3-none-any.whl"
    payload = bytearray(wheel.read_bytes())
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
    wheel.write_bytes(mutated)

    with pytest.raises(ArtifactError, match="hidden data between ZIP member records"):
        _check_wheel(wheel, VERSION)


@pytest.mark.parametrize("position", ("before", "after", "concatenated"))
def test_sdist_rejects_hidden_bytes_outside_tar_stream(
    tmp_path: Path,
    position: str,
) -> None:
    _write_valid_distributions(tmp_path)
    sdist = tmp_path / f"workbooklens-{VERSION}.tar.gz"
    payload = sdist.read_bytes()
    hidden = b"EMBEDDED-SECRET"
    if position == "before":
        mutated = hidden + payload
    elif position == "after":
        mutated = payload + hidden
    else:
        mutated = payload + gzip.compress(hidden)
    sdist.write_bytes(mutated)

    with pytest.raises(ArtifactError, match=r"gzip|hidden data"):
        check_distribution(tmp_path, VERSION)


def test_sdist_rejects_hidden_bytes_in_member_padding(tmp_path: Path) -> None:
    _write_valid_distributions(tmp_path)
    sdist = tmp_path / f"workbooklens-{VERSION}.tar.gz"
    payload = bytearray(gzip.decompress(sdist.read_bytes()))
    with tarfile.open(fileobj=io.BytesIO(payload), mode="r:") as archive:
        member = next(item for item in archive.getmembers() if item.isfile() and item.size % 512)
    payload[member.offset_data + member.size] = ord("X")
    sdist.write_bytes(gzip.compress(bytes(payload), mtime=0))

    with pytest.raises(ArtifactError, match="tar padding"):
        check_distribution(tmp_path, VERSION)
