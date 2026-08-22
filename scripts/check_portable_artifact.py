from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import struct
import unicodedata
import zipfile
import zlib
from contextlib import suppress
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Final, NoReturn

ARCHIVE_RE: Final = re.compile(
    r"^WorkbookLens-(?P<version>[A-Za-z0-9][A-Za-z0-9._+-]*)-"
    r"windows-x64-portable\.zip$"
)
WINDOWS_DRIVE_RE: Final = re.compile(r"^[A-Za-z]:")
WINDOWS_DEVICES: Final = {
    "AUX",
    "CLOCK$",
    "CON",
    "CONIN$",
    "CONOUT$",
    "NUL",
    "PRN",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
    "COM¹",
    "COM²",
    "COM³",
    "LPT¹",
    "LPT²",
    "LPT³",
}
WINDOWS_INVALID_CHARACTERS: Final = frozenset('<>"|?*')
ALLOWED_ROOT_FILES: Final = {
    "LICENSE",
    "README-PORTABLE.txt",
    "Start-WorkbookLens.cmd",
    "THIRD-PARTY-NOTICES.txt",
    "WorkbookLens.exe",
    "workbooklens.example.yml",
}
ALLOWED_ROOT_DIRECTORIES: Final = {"LICENSES", "_internal"}
SENSITIVE_DIRECTORY_NAMES: Final = {
    ".aws",
    ".azure",
    ".git",
    ".gnupg",
    ".hg",
    ".ssh",
    ".svn",
}
SENSITIVE_FILE_NAMES: Final = {
    ".env",
    "credentials",
    "credentials.json",
    "id_dsa",
    "id_ecdsa",
    "id_ed25519",
    "id_rsa",
    "pyvenv.cfg",
    "secrets.json",
    "token.json",
}
SENSITIVE_SUFFIXES: Final = {".key", ".kdbx", ".p12", ".pfx"}
ALLOWED_COMPRESSION: Final = {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED}
LICENSE_FORBIDDEN_SUFFIXES: Final = {
    ".bat",
    ".cmd",
    ".dll",
    ".exe",
    ".lnk",
    ".msi",
    ".ps1",
    ".pyd",
    ".reg",
    ".url",
    ".vbs",
}
INTERNAL_FORBIDDEN_DIRECTORIES: Final = {
    ".pytest_cache",
    ".venv",
    "__pycache__",
    "build",
    "dist",
    "tests",
}
INTERNAL_FORBIDDEN_FILES: Final = {"direct_url.json"}
INTERNAL_FORBIDDEN_SUFFIXES: Final = {
    ".bat",
    ".c",
    ".cmd",
    ".cpp",
    ".db",
    ".dmp",
    ".dump",
    ".exe",
    ".h",
    ".lnk",
    ".log",
    ".map",
    ".msi",
    ".ods",
    ".pdb",
    ".ps1",
    ".py",
    ".pyi",
    ".reg",
    ".sqlite",
    ".url",
    ".vbs",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
}
INTERNAL_SENSITIVE_TOKENS: Final = {
    "credential",
    "credentials",
    "debug",
    "passwd",
    "password",
    "secret",
    "secrets",
}
REQUIRED_INTERNAL_FILES: Final = {
    "base_library.zip",
    "python312.dll",
    "workbooklens/diff/templates/diff.html.j2",
    "workbooklens/reports/templates/scan.html.j2",
}
MAX_LICENSE_FILE_SIZE: Final = 4 * 1024 * 1024


class PortableArtifactError(ValueError):
    pass


@dataclass(frozen=True)
class ArtifactLimits:
    max_archive_size: int = 250 * 1024 * 1024
    max_members: int = 10_000
    max_total_uncompressed: int = 500 * 1024 * 1024
    max_member_uncompressed: int = 100 * 1024 * 1024
    max_compression_ratio: float = 100.0


DEFAULT_ARTIFACT_LIMITS: Final = ArtifactLimits()


@dataclass(frozen=True)
class ArtifactReport:
    archive: str
    version: str
    root: str
    members: int
    total_uncompressed: int
    sha256: str


def _fail(message: str) -> NoReturn:
    raise PortableArtifactError(message)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validate_checksum(archive: Path, expected_digest: str) -> None:
    sidecar = archive.with_suffix(archive.suffix + ".sha256")
    if not sidecar.is_file():
        _fail(f"missing SHA256 sidecar: {sidecar.name}")
    try:
        text = sidecar.read_text(encoding="ascii").strip()
    except (OSError, UnicodeError) as exc:
        _fail(f"cannot read SHA256 sidecar: {exc}")
    match = re.fullmatch(r"([0-9A-Fa-f]{64})  ([^\r\n]+)", text)
    if match is None:
        _fail("SHA256 sidecar must contain '<digest>  <archive-name>'")
    digest, name = match.groups()
    if name != archive.name:
        _fail(f"SHA256 sidecar names {name!r}, expected {archive.name!r}")
    if digest.casefold() != expected_digest.casefold():
        _fail("SHA256 sidecar does not match the archive")


def _validate_archive_envelope(
    archive: Path,
    limits: ArtifactLimits,
) -> tuple[int, int]:
    try:
        archive_size = archive.stat().st_size
    except OSError as exc:
        _fail(f"cannot stat portable archive: {exc}")
    if archive_size > limits.max_archive_size:
        _fail(
            f"portable archive is too large: {archive_size} bytes (limit {limits.max_archive_size})"
        )
    if archive_size < 22:
        _fail("portable archive is too small to contain a ZIP end record")
    try:
        with archive.open("rb") as stream:
            if stream.read(4) != b"PK\x03\x04":
                _fail("ZIP preamble or prepended executable data is not allowed")
            stream.seek(-22, os.SEEK_END)
            end_record = stream.read(22)
    except OSError as exc:
        _fail(f"cannot inspect ZIP end record: {exc}")
    if (
        len(end_record) != 22
        or end_record[:4] != b"PK\x05\x06"
        or struct.unpack_from("<H", end_record, 20)[0] != 0
    ):
        _fail("ZIP must have an uncommented end record exactly at EOF")
    (
        _signature,
        disk_number,
        directory_disk,
        entries_on_disk,
        total_entries,
        directory_size,
        directory_offset,
        _comment_size,
    ) = struct.unpack("<4s4H2LH", end_record)
    if (
        disk_number == 0xFFFF
        or directory_disk == 0xFFFF
        or entries_on_disk == 0xFFFF
        or total_entries == 0xFFFF
        or directory_size == 0xFFFFFFFF
        or directory_offset == 0xFFFFFFFF
    ):
        _fail("ZIP64 end records are unnecessary and not allowed for this bounded artifact")
    if disk_number or directory_disk or entries_on_disk != total_entries:
        _fail("multi-disk or inconsistent ZIP directory is not allowed")
    if directory_offset + directory_size != archive_size - 22:
        _fail("ZIP central directory is not contiguous with the end record")
    return int(directory_offset), int(total_entries)


def _validate_windows_component(component: str, member_name: str) -> str:
    normalized = unicodedata.normalize("NFC", component)
    if normalized in {"", ".", ".."}:
        _fail(f"unsafe path component in {member_name!r}")
    if normalized != normalized.rstrip(" ."):
        _fail(f"Windows-trimmed path component in {member_name!r}")
    if ":" in normalized:
        _fail(f"alternate data stream or drive syntax in {member_name!r}")
    if any(character in WINDOWS_INVALID_CHARACTERS for character in normalized):
        _fail(f"Windows-invalid character in ZIP member: {member_name!r}")
    if any(unicodedata.category(character) == "Cc" for character in normalized):
        _fail(f"control character in ZIP member: {member_name!r}")
    if len(normalized.encode("utf-16-le")) // 2 > 255:
        _fail(f"Windows path component is too long: {member_name!r}")
    device_stem = normalized.split(".", 1)[0].upper()
    if device_stem in WINDOWS_DEVICES:
        _fail(f"Windows device name in {member_name!r}")
    return normalized


def _normalized_member_parts(info: zipfile.ZipInfo) -> tuple[str, ...]:
    name = info.filename
    if info.orig_filename != name:
        _fail(f"ZIP member filename was normalized or truncated: {name!r}")
    if not name or "\x00" in name:
        _fail("ZIP member has an empty name or NUL byte")
    if "\\" in name:
        _fail(f"ZIP member uses an ambiguous backslash path: {name!r}")
    if name.startswith(("/", "//")) or WINDOWS_DRIVE_RE.match(name):
        _fail(f"ZIP member has an absolute path: {name!r}")

    raw_parts = name[:-1].split("/") if name.endswith("/") else name.split("/")
    if any(part == ".." for part in raw_parts):
        _fail(f"ZIP member traverses outside the archive root: {name!r}")
    parts = tuple(_validate_windows_component(part, name) for part in raw_parts)
    if len("/".join(parts).encode("utf-16-le")) // 2 > 240:
        _fail(f"ZIP member path exceeds the Windows portable limit: {name!r}")
    return parts


def _validate_member_type(info: zipfile.ZipInfo) -> None:
    if info.flag_bits & 0x1:
        _fail(f"encrypted ZIP member is not allowed: {info.filename!r}")
    if info.compress_type not in ALLOWED_COMPRESSION:
        _fail(f"unsupported ZIP compression method for {info.filename!r}")
    if info.comment:
        _fail(f"ZIP member comments are not allowed: {info.filename!r}")
    if info.extra:
        _fail(f"central-directory ZIP extra fields are not allowed: {info.filename!r}")

    if info.create_system != 3:
        return
    mode = (info.external_attr >> 16) & 0xFFFF
    file_type = stat.S_IFMT(mode)
    if file_type == stat.S_IFLNK:
        _fail(f"symbolic link is not allowed: {info.filename!r}")
    if file_type not in {0, stat.S_IFREG, stat.S_IFDIR}:
        _fail(f"special filesystem entry is not allowed: {info.filename!r}")


def _validate_local_zip64_extra(
    extra: bytes,
    *,
    compressed_size: int,
    uncompressed_size: int,
    info: zipfile.ZipInfo,
) -> None:
    if not extra:
        if compressed_size == 0xFFFFFFFF or uncompressed_size == 0xFFFFFFFF:
            _fail(f"local ZIP64 sizes are missing for {info.filename!r}")
        return
    if len(extra) < 4:
        _fail(f"truncated local ZIP extra field: {info.filename!r}")
    header_id, field_size = struct.unpack_from("<HH", extra)
    if header_id != 0x0001 or field_size + 4 != len(extra):
        _fail(f"unexpected local ZIP extra field: {info.filename!r}")
    values = extra[4:]
    cursor = 0
    if uncompressed_size == 0xFFFFFFFF:
        if cursor + 8 > len(values):
            _fail(f"truncated ZIP64 uncompressed size: {info.filename!r}")
        actual = struct.unpack_from("<Q", values, cursor)[0]
        cursor += 8
        if actual != info.file_size:
            _fail(f"ZIP64 uncompressed size mismatch: {info.filename!r}")
    if compressed_size == 0xFFFFFFFF:
        if cursor + 8 > len(values):
            _fail(f"truncated ZIP64 compressed size: {info.filename!r}")
        actual = struct.unpack_from("<Q", values, cursor)[0]
        cursor += 8
        if actual != info.compress_size:
            _fail(f"ZIP64 compressed size mismatch: {info.filename!r}")
    if cursor != len(values):
        _fail(f"unused or hidden bytes in local ZIP64 extra field: {info.filename!r}")


def _descriptor_matches(payload: bytes, info: zipfile.ZipInfo) -> bool:
    if len(payload) == 12:
        crc, compressed, uncompressed = struct.unpack("<LLL", payload)
    elif len(payload) == 16 and payload[:4] == b"PK\x07\x08":
        _signature, crc, compressed, uncompressed = struct.unpack("<4sLLL", payload)
    elif len(payload) == 20:
        crc, compressed, uncompressed = struct.unpack("<LQQ", payload)
    elif len(payload) == 24 and payload[:4] == b"PK\x07\x08":
        _signature, crc, compressed, uncompressed = struct.unpack("<4sLQQ", payload)
    else:
        return False
    return crc == info.CRC and compressed == info.compress_size and uncompressed == info.file_size


def _validate_local_records(
    archive: Path,
    infos: list[zipfile.ZipInfo],
    directory_offset: int,
    limits: ArtifactLimits,
) -> None:
    ordered = sorted(infos, key=lambda info: info.header_offset)
    if not ordered or ordered[0].header_offset != 0:
        _fail("ZIP preamble or prepended executable data is not allowed")
    try:
        stream = archive.open("rb")
    except OSError as exc:
        _fail(f"cannot inspect local ZIP records: {exc}")
    with stream:
        for index, info in enumerate(ordered):
            boundary = (
                ordered[index + 1].header_offset if index + 1 < len(ordered) else directory_offset
            )
            stream.seek(info.header_offset)
            header = stream.read(30)
            if len(header) != 30 or header[:4] != b"PK\x03\x04":
                _fail(f"invalid local ZIP header for {info.filename!r}")
            (
                _signature,
                _extract_version,
                flags,
                compression,
                _modified_time,
                _modified_date,
                crc,
                compressed_size,
                uncompressed_size,
                name_size,
                extra_size,
            ) = struct.unpack("<4s5H3L2H", header)
            if flags != info.flag_bits or compression != info.compress_type:
                _fail(f"local/central ZIP metadata mismatch: {info.filename!r}")
            raw_name = stream.read(name_size)
            local_extra = stream.read(extra_size)
            encoding = "utf-8" if flags & 0x800 else "cp437"
            try:
                local_name = raw_name.decode(encoding)
            except UnicodeDecodeError as exc:
                _fail(f"invalid local ZIP filename encoding: {info.filename!r}: {exc}")
            if local_name != info.filename:
                _fail(f"local/central ZIP filename mismatch: {info.filename!r}")
            _validate_local_zip64_extra(
                local_extra,
                compressed_size=compressed_size,
                uncompressed_size=uncompressed_size,
                info=info,
            )

            payload_start = info.header_offset + 30 + name_size + extra_size
            payload_end = payload_start + info.compress_size
            if payload_end > boundary:
                _fail(f"overlapping local ZIP records: {info.filename!r}")
            stream.seek(payload_start)
            compressed_payload = stream.read(info.compress_size)
            if len(compressed_payload) != info.compress_size:
                _fail(f"truncated ZIP payload: {info.filename!r}")
            if info.compress_type == zipfile.ZIP_STORED:
                expanded = compressed_payload
            else:
                decompressor = zlib.decompressobj(-15)
                try:
                    expanded = decompressor.decompress(
                        compressed_payload,
                        limits.max_member_uncompressed + 1,
                    )
                    expanded += decompressor.flush()
                except zlib.error as exc:
                    _fail(f"invalid deflate stream for {info.filename!r}: {exc}")
                if not decompressor.eof or decompressor.unused_data or decompressor.unconsumed_tail:
                    _fail(f"deflate stream has hidden or trailing data: {info.filename!r}")
            if len(expanded) != info.file_size or zlib.crc32(expanded) & 0xFFFFFFFF != info.CRC:
                _fail(f"local ZIP payload metadata mismatch: {info.filename!r}")

            if flags & 0x08:
                stream.seek(payload_end)
                descriptor = stream.read(boundary - payload_end)
                if not _descriptor_matches(descriptor, info):
                    _fail(f"invalid or non-contiguous data descriptor: {info.filename!r}")
            else:
                if payload_end != boundary:
                    _fail(f"hidden data between local ZIP records: {info.filename!r}")
                if crc != info.CRC:
                    _fail(f"local ZIP CRC mismatch: {info.filename!r}")
                if compressed_size not in {info.compress_size, 0xFFFFFFFF}:
                    _fail(f"local ZIP compressed size mismatch: {info.filename!r}")
                if uncompressed_size not in {info.file_size, 0xFFFFFFFF}:
                    _fail(f"local ZIP uncompressed size mismatch: {info.filename!r}")


def _validate_member_size(info: zipfile.ZipInfo, limits: ArtifactLimits) -> None:
    if info.file_size < 0 or info.compress_size < 0:
        _fail(f"negative ZIP member size: {info.filename!r}")
    if info.file_size > limits.max_member_uncompressed:
        _fail(f"ZIP member is too large: {info.filename!r}")
    if info.file_size == 0:
        return
    if info.compress_size == 0:
        _fail(f"non-empty ZIP member has zero compressed size: {info.filename!r}")
    ratio = info.file_size / info.compress_size
    if ratio > limits.max_compression_ratio:
        _fail(f"ZIP member compression ratio {ratio:.1f} exceeds the limit: {info.filename!r}")


def _validate_path_topology(
    entries: dict[tuple[str, ...], tuple[str, bool]],
    parts: tuple[str, ...],
    info: zipfile.ZipInfo,
) -> None:
    canonical = tuple(part.casefold() for part in parts)
    for index in range(1, len(canonical)):
        ancestor = canonical[:index]
        previous = entries.get(ancestor)
        if previous is not None and not previous[1]:
            _fail(
                "file/directory ancestor conflict: "
                f"{previous[0]!r} is a file above {info.filename!r}"
            )
    if not info.is_dir():
        for previous_parts, (previous_name, _) in entries.items():
            if (
                len(previous_parts) > len(canonical)
                and previous_parts[: len(canonical)] == canonical
            ):
                _fail(
                    "file/directory ancestor conflict: "
                    f"{info.filename!r} is a file above {previous_name!r}"
                )
    entries[canonical] = (info.filename, info.is_dir())


def _validate_sensitive_path(parts: tuple[str, ...], member_name: str) -> None:
    folded = tuple(part.casefold() for part in parts)
    if any(part in SENSITIVE_DIRECTORY_NAMES for part in folded):
        _fail(f"sensitive directory included in portable archive: {member_name!r}")
    leaf = folded[-1]
    if leaf in SENSITIVE_FILE_NAMES or leaf.startswith(".env."):
        _fail(f"sensitive file included in portable archive: {member_name!r}")
    if any(leaf.endswith(suffix) for suffix in SENSITIVE_SUFFIXES):
        _fail(f"private-key or credential file included: {member_name!r}")


def _validate_scoped_member(
    zf: zipfile.ZipFile,
    info: zipfile.ZipInfo,
    parts: tuple[str, ...],
) -> None:
    if len(parts) < 3:
        return
    scope = parts[1].casefold()
    scoped_parts = parts[2:]

    if scope == "licenses":
        if info.is_dir():
            return
        if info.file_size > MAX_LICENSE_FILE_SIZE:
            _fail(f"license file exceeds 4 MiB: {info.filename!r}")
        if Path(scoped_parts[-1]).suffix.casefold() in LICENSE_FORBIDDEN_SUFFIXES:
            _fail(f"executable content is forbidden under LICENSES: {info.filename!r}")
        try:
            _read_member(zf, info.filename).decode("utf-8-sig")
        except UnicodeDecodeError as exc:
            _fail(f"license file is not decodable UTF-8 text: {info.filename!r}: {exc}")
        return

    if scope != "_internal":
        return
    directory_parts = scoped_parts if info.is_dir() else scoped_parts[:-1]
    if any(part.casefold() in INTERNAL_FORBIDDEN_DIRECTORIES for part in directory_parts):
        _fail(f"development directory is forbidden under _internal: {info.filename!r}")
    if info.is_dir():
        return

    leaf = scoped_parts[-1].casefold()
    if leaf in INTERNAL_FORBIDDEN_FILES:
        _fail(f"build metadata is forbidden under _internal: {info.filename!r}")
    suffix = Path(leaf).suffix.casefold()
    if suffix in INTERNAL_FORBIDDEN_SUFFIXES:
        _fail(f"unexpected source, data, or executable file under _internal: {info.filename!r}")
    tokens = {token for token in re.split(r"[^a-z0-9]+", leaf) if token}
    if tokens & INTERNAL_SENSITIVE_TOKENS:
        _fail(f"sensitive filename token under _internal: {info.filename!r}")


def _validate_pe_x64(
    data: bytes,
    *,
    member_name: str,
    version: str | None = None,
) -> None:
    if len(data) < 0x40 or data[:2] != b"MZ":
        _fail(f"{member_name} is not a PE executable")
    pe_offset = struct.unpack_from("<I", data, 0x3C)[0]
    if pe_offset + 26 > len(data) or data[pe_offset : pe_offset + 4] != b"PE\0\0":
        _fail(f"{member_name} has an invalid PE header")
    machine = struct.unpack_from("<H", data, pe_offset + 4)[0]
    optional_magic = struct.unpack_from("<H", data, pe_offset + 24)[0]
    if machine != 0x8664 or optional_magic != 0x20B:
        _fail(f"{member_name} is not a Windows x64 PE32+ executable")
    if version is not None and (
        version.encode("ascii") not in data and version.encode("utf-16le") not in data
    ):
        _fail(f"{member_name} does not contain the expected product version")


def _read_member(zf: zipfile.ZipFile, name: str) -> bytes:
    try:
        return zf.read(name)
    except (KeyError, OSError, RuntimeError, zipfile.BadZipFile) as exc:
        _fail(f"cannot read required ZIP member {name!r}: {exc}")


def _compare_repository_templates(
    zf: zipfile.ZipFile,
    root: str,
    repository_root: Path,
    version: str,
) -> None:
    expected = {
        "LICENSE": repository_root / "LICENSE",
        "Start-WorkbookLens.cmd": repository_root
        / "packaging"
        / "windows"
        / "Start-WorkbookLens.cmd",
        "workbooklens.example.yml": repository_root / "workbooklens.example.yml",
    }
    for archive_name, source in expected.items():
        if not source.is_file():
            _fail(f"repository template is missing: {source}")
        if _read_member(zf, f"{root}/{archive_name}") != source.read_bytes():
            _fail(f"archive member does not match repository template: {archive_name}")

    readme_template = repository_root / "packaging" / "windows" / "README-PORTABLE.txt"
    if not readme_template.is_file():
        _fail(f"repository template is missing: {readme_template}")
    try:
        template_text = readme_template.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        _fail(f"cannot read repository README template: {exc}")
    if template_text.count("@VERSION@") != 1:
        _fail("README-PORTABLE.txt template must contain exactly one @VERSION@")
    expected_readme = template_text.replace("@VERSION@", version).encode("utf-8")
    if _read_member(zf, f"{root}/README-PORTABLE.txt") != expected_readme:
        _fail("README-PORTABLE.txt does not exactly match the repository template")


def inspect_artifact(
    archive: Path,
    *,
    expected_version: str | None = None,
    repository_root: Path | None = None,
    limits: ArtifactLimits = DEFAULT_ARTIFACT_LIMITS,
    require_checksum: bool = True,
) -> ArtifactReport:
    archive = archive.resolve()
    if not archive.is_file():
        _fail(f"portable archive does not exist: {archive}")
    directory_offset, expected_entries = _validate_archive_envelope(archive, limits)
    match = ARCHIVE_RE.fullmatch(archive.name)
    if match is None:
        _fail(f"unexpected portable archive name: {archive.name!r}")
    version = match.group("version")
    if expected_version is not None and version != expected_version:
        _fail(f"archive version {version!r} does not match {expected_version!r}")
    root = f"WorkbookLens-{version}-windows-x64"
    digest = _sha256(archive)
    if require_checksum:
        _validate_checksum(archive, digest)

    try:
        zf = zipfile.ZipFile(archive)
    except (OSError, zipfile.BadZipFile) as exc:
        _fail(f"cannot open portable archive: {exc}")

    with zf:
        if zf.comment:
            _fail("ZIP comments are not allowed")
        infos = zf.infolist()
        if not infos:
            _fail("portable archive is empty")
        if len(infos) > limits.max_members:
            _fail(f"portable archive has too many members: {len(infos)}")
        if len(infos) != expected_entries:
            _fail("ZIP end record member count does not match the central directory")
        _validate_local_records(archive, infos, directory_offset, limits)

        canonical_names: dict[str, str] = {}
        path_entries: dict[tuple[str, ...], tuple[str, bool]] = {}
        files: set[str] = set()
        root_directories: set[str] = set()
        internal_files: set[str] = set()
        license_distribution_dirs: set[str] = set()
        pe_members: list[str] = []
        total_uncompressed = 0

        for info in infos:
            _validate_member_type(info)
            _validate_member_size(info, limits)
            parts = _normalized_member_parts(info)
            _validate_sensitive_path(parts, info.filename)
            canonical = "/".join(parts).casefold()
            previous = canonical_names.get(canonical)
            if previous is not None:
                _fail(
                    "duplicate or case/Unicode-colliding ZIP members: "
                    f"{previous!r} and {info.filename!r}"
                )
            canonical_names[canonical] = info.filename
            _validate_path_topology(path_entries, parts, info)

            if parts[0] != root:
                _fail(f"ZIP member is outside the expected root {root!r}: {info.filename!r}")
            relative = parts[1:]
            if not relative:
                if not info.is_dir():
                    _fail("archive root entry must be a directory")
            elif len(relative) == 1 and not info.is_dir():
                if relative[0] not in ALLOWED_ROOT_FILES:
                    _fail(f"unexpected file at archive root: {relative[0]!r}")
                files.add(relative[0])
            else:
                top = relative[0]
                if top not in ALLOWED_ROOT_DIRECTORIES:
                    _fail(f"unexpected directory at archive root: {top!r}")
                root_directories.add(top)
                if top == "_internal" and not info.is_dir():
                    internal_name = "/".join(relative[1:]).casefold()
                    internal_files.add(internal_name)
                    if Path(relative[-1]).suffix.casefold() in {".dll", ".pyd"}:
                        pe_members.append(info.filename)
                elif top == "LICENSES" and (
                    len(relative) >= 3 or (len(relative) == 2 and info.is_dir())
                ):
                    license_distribution_dirs.add(relative[1])

            _validate_scoped_member(zf, info, parts)

            total_uncompressed += info.file_size
            if total_uncompressed > limits.max_total_uncompressed:
                _fail("portable archive exceeds the total uncompressed-size limit")

        missing_files = ALLOWED_ROOT_FILES - files
        if missing_files:
            _fail(f"portable archive is missing root files: {sorted(missing_files)!r}")
        missing_directories = ALLOWED_ROOT_DIRECTORIES - root_directories
        if missing_directories:
            _fail(
                "portable archive is missing populated directories: "
                f"{sorted(missing_directories)!r}"
            )
        missing_internal = REQUIRED_INTERNAL_FILES - internal_files
        if missing_internal:
            _fail(
                f"portable archive is missing required runtime files: {sorted(missing_internal)!r}"
            )

        try:
            corrupt_member = zf.testzip()
        except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
            _fail(f"cannot verify ZIP member CRCs: {exc}")
        if corrupt_member is not None:
            _fail(f"ZIP member failed CRC verification: {corrupt_member!r}")

        exe_name = f"{root}/WorkbookLens.exe"
        _validate_pe_x64(
            _read_member(zf, exe_name),
            member_name=exe_name,
            version=version,
        )
        for member_name in pe_members:
            _validate_pe_x64(
                _read_member(zf, member_name),
                member_name=member_name,
            )

        readme = _read_member(zf, f"{root}/README-PORTABLE.txt")
        try:
            readme_text = readme.decode("utf-8")
        except UnicodeDecodeError as exc:
            _fail(f"README-PORTABLE.txt is not UTF-8: {exc}")
        if f"WorkbookLens {version}" not in readme_text or "@VERSION@" in readme_text:
            _fail("README-PORTABLE.txt does not contain the expected version")

        notices = _read_member(zf, f"{root}/THIRD-PARTY-NOTICES.txt")
        try:
            notices_text = notices.decode("utf-8")
        except UnicodeDecodeError as exc:
            _fail(f"THIRD-PARTY-NOTICES.txt is not UTF-8: {exc}")
        for required_notice in ("CPython", "PyInstaller"):
            if required_notice not in notices_text:
                _fail(f"THIRD-PARTY-NOTICES.txt does not mention {required_notice}")
        for directory in sorted(license_distribution_dirs):
            if directory not in notices_text:
                _fail(f"THIRD-PARTY-NOTICES.txt does not reference license directory {directory!r}")
        cpython_license = f"{root}/LICENSES/CPython-3.12-LICENSE.txt".casefold()
        if cpython_license not in canonical_names:
            _fail("portable archive is missing the CPython 3.12 license")

        if repository_root is not None:
            _compare_repository_templates(
                zf,
                root,
                repository_root.resolve(),
                version,
            )

    return ArtifactReport(
        archive=str(archive),
        version=version,
        root=root,
        members=len(infos),
        total_uncompressed=total_uncompressed,
        sha256=digest,
    )


def _is_reparse_point(path: Path) -> bool:
    if path.is_symlink():
        return True
    try:
        attributes = getattr(os.lstat(path), "st_file_attributes", 0)
    except OSError:
        return False
    reparse_flag = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    return bool(attributes & reparse_flag)


def _prepare_extraction_destination(destination: Path) -> Path:
    destination = destination.absolute()
    if destination.exists() or destination.is_symlink():
        if _is_reparse_point(destination):
            _fail(f"extraction destination is a reparse point: {destination}")
        if not destination.is_dir():
            _fail(f"extraction destination is not a directory: {destination}")
        try:
            if next(destination.iterdir(), None) is not None:
                _fail(f"extraction destination must be empty: {destination}")
        except OSError as exc:
            _fail(f"cannot inspect extraction destination: {exc}")
    else:
        try:
            destination.mkdir(parents=True, exist_ok=False)
        except OSError as exc:
            _fail(f"cannot create extraction destination: {exc}")
    if _is_reparse_point(destination):
        _fail(f"extraction destination is a reparse point: {destination}")
    return destination


def _hash_extracted_file(path: Path) -> tuple[int, str]:
    digest = hashlib.sha256()
    try:
        size = path.stat().st_size
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as exc:
        _fail(f"cannot verify extracted file {path}: {exc}")
    return size, digest.hexdigest()


def _remove_failed_extraction(destination: Path) -> None:
    with suppress(OSError):
        shutil.rmtree(destination)


def extract_checked_artifact(
    archive: Path,
    destination: Path,
    *,
    expected_version: str,
    repository_root: Path | None = None,
) -> Path:
    report = inspect_artifact(
        archive,
        expected_version=expected_version,
        repository_root=repository_root,
    )
    destination = _prepare_extraction_destination(destination)
    destination_resolved = destination.resolve()
    expected_files: dict[str, tuple[int, str]] = {}

    try:
        with zipfile.ZipFile(archive) as zf:
            for info in zf.infolist():
                parts = _normalized_member_parts(info)
                relative_name = "/".join(parts)
                target = destination.joinpath(*parts)
                resolved_target = target.resolve()
                if not resolved_target.is_relative_to(destination_resolved):
                    _fail(f"extraction target escapes destination: {info.filename!r}")
                if info.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                digest = hashlib.sha256()
                try:
                    with zf.open(info, "r") as source, target.open("xb") as output:
                        while chunk := source.read(1024 * 1024):
                            output.write(chunk)
                            digest.update(chunk)
                except FileExistsError:
                    _fail(f"extraction would overwrite an existing file: {target}")
                expected_files[relative_name] = (info.file_size, digest.hexdigest())

        actual_files: dict[str, tuple[int, str]] = {}
        for path in destination.rglob("*"):
            if _is_reparse_point(path):
                _fail(f"extraction created a reparse point: {path}")
            if path.is_file():
                relative_name = path.relative_to(destination).as_posix()
                actual_files[relative_name] = _hash_extracted_file(path)
            elif not path.is_dir():
                _fail(f"extraction created an unsupported filesystem entry: {path}")
        if actual_files.keys() != expected_files.keys():
            _fail("extracted file set does not match the validated archive")
        for relative_name, expected in expected_files.items():
            if actual_files[relative_name] != expected:
                _fail(
                    f"extracted file size or SHA256 does not match the archive: {relative_name!r}"
                )
    except PortableArtifactError:
        _remove_failed_extraction(destination)
        raise
    except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
        _remove_failed_extraction(destination)
        _fail(f"portable artifact extraction failed: {exc}")

    return destination / report.root


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Validate a WorkbookLens Windows portable ZIP before release."
    )
    parser.add_argument("archive", type=Path)
    parser.add_argument("--expected-version")
    parser.add_argument(
        "--repository-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
        help="Repository root used for byte-for-byte template checks.",
    )
    parser.add_argument("--json", action="store_true", dest="as_json")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        report = inspect_artifact(
            args.archive,
            expected_version=args.expected_version,
            repository_root=args.repository_root,
        )
    except PortableArtifactError as exc:
        print(f"portable artifact check failed: {exc}")
        return 1
    if args.as_json:
        print(json.dumps(asdict(report), indent=2, sort_keys=True))
    else:
        print(
            f"portable artifact OK: {report.archive} "
            f"({report.members} members, sha256={report.sha256})"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
