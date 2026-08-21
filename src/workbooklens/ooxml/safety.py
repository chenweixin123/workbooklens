"""Fail-closed inspection of untrusted Office Open XML ZIP packages."""

from __future__ import annotations

import posixpath
import re
import stat
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from urllib.parse import urlsplit

from lxml import etree

from workbooklens.exceptions import UnsafeWorkbookError, UsageError

RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REQUIRED_PARTS = {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"}
FORBIDDEN_XML_DECLARATION_RE = re.compile(rb"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE)
XLSX_WORKBOOK_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)
XLSM_WORKBOOK_CONTENT_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
VBA_RELATIONSHIP_NAMES = {"vbaproject", "vbaprojectsignature"}


@dataclass(frozen=True, slots=True)
class PackageLimits:
    """Resource ceilings applied before XML or workbook parsing."""

    max_file_bytes: int = 100 * 1024 * 1024
    max_entries: int = 10_000
    max_total_uncompressed_bytes: int = 1_000 * 1024 * 1024
    max_entry_uncompressed_bytes: int = 100 * 1024 * 1024
    max_xml_bytes: int = 50 * 1024 * 1024
    max_compression_ratio: float = 100.0


@dataclass(frozen=True, slots=True)
class PackageInspection:
    """Auditable summary returned after a package passes all safety checks."""

    path: Path
    format: str
    entry_count: int
    total_compressed_bytes: int
    total_uncompressed_bytes: int
    part_names: frozenset[str]
    external_relationships: tuple[str, ...] = field(default_factory=tuple)

    content_format: str | None = None
    extension_format: str = ""
    has_vba: bool = False
    format_mismatch: bool = False
    repairable: bool = False


def _canonical_member_name(name: str) -> str:
    if not name or "\x00" in name:
        raise UnsafeWorkbookError("ZIP package contains an empty or NUL-containing entry name")
    if "\\" in name:
        raise UnsafeWorkbookError(f"ZIP entry uses a non-portable backslash path: {name!r}")
    if name.startswith("/") or urlsplit(name).scheme:
        raise UnsafeWorkbookError(f"ZIP entry is absolute or URI-like: {name!r}")
    path = PurePosixPath(name)
    if any(part in {"", ".", ".."} for part in path.parts):
        raise UnsafeWorkbookError(f"ZIP entry contains unsafe path components: {name!r}")
    canonical = path.as_posix().rstrip("/")
    if not canonical:
        raise UnsafeWorkbookError(f"ZIP entry has no usable package path: {name!r}")
    return canonical


def _is_symlink(info: zipfile.ZipInfo) -> bool:
    unix_mode = info.external_attr >> 16
    return stat.S_ISLNK(unix_mode)


def _read_bounded(archive: zipfile.ZipFile, info: zipfile.ZipInfo, limit: int) -> bytes:
    if info.file_size > limit:
        raise UnsafeWorkbookError(f"Package part {info.filename!r} exceeds the {limit}-byte limit")
    with archive.open(info, "r") as handle:
        data = handle.read(limit + 1)
    if len(data) > limit:
        raise UnsafeWorkbookError(f"Package part {info.filename!r} expanded beyond its limit")
    return data


def secure_xml_parser() -> etree.XMLParser:
    """Create a fresh hardened lxml parser; parser instances are not shared across threads."""

    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        dtd_validation=False,
        huge_tree=False,
        recover=False,
        remove_comments=False,
        remove_blank_text=False,
    )


def parse_xml_part(data: bytes, part_name: str) -> etree._Element:
    """Parse one already-bounded OOXML part without DTDs or entity expansion."""

    if FORBIDDEN_XML_DECLARATION_RE.search(data):
        raise UnsafeWorkbookError(f"DTD or entity declarations are forbidden in {part_name!r}")
    try:
        return etree.fromstring(data, parser=secure_xml_parser())
    except (etree.XMLSyntaxError, ValueError) as exc:
        raise UnsafeWorkbookError(f"Malformed XML in package part {part_name!r}: {exc}") from exc


def _relationship_source(rels_name: str) -> str:
    if rels_name == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in rels_name or not rels_name.endswith(".rels"):
        raise UnsafeWorkbookError(f"Malformed relationship part name: {rels_name!r}")
    directory, filename = rels_name.split(marker, maxsplit=1)
    return posixpath.join(directory, filename[: -len(".rels")])


def _resolve_internal_target(rels_name: str, target: str) -> str:
    if not target or "\x00" in target or "\\" in target:
        raise UnsafeWorkbookError(f"Unsafe relationship target {target!r} in {rels_name!r}")
    split = urlsplit(target)
    if split.scheme or split.netloc:
        raise UnsafeWorkbookError(
            f"URI relationship target lacks TargetMode=External in {rels_name!r}: {target!r}"
        )
    if target.startswith("/"):
        candidate = target.lstrip("/")
    else:
        source = _relationship_source(rels_name)
        candidate = posixpath.join(posixpath.dirname(source), target)
    normalized = posixpath.normpath(candidate)
    if normalized in {"", ".", ".."} or normalized.startswith("../"):
        raise UnsafeWorkbookError(f"Relationship escapes the package root: {target!r}")
    return _canonical_member_name(normalized)


def _inspect_relationships(
    archive: zipfile.ZipFile,
    infos: dict[str, zipfile.ZipInfo],
    limits: PackageLimits,
) -> tuple[tuple[str, ...], bool]:
    external: list[str] = []
    has_vba_relationship = False
    for name, info in infos.items():
        if not name.endswith(".rels"):
            continue
        data = _read_bounded(archive, info, limits.max_xml_bytes)
        root = parse_xml_part(data, name)
        if root.tag != f"{{{RELATIONSHIPS_NS}}}Relationships":
            raise UnsafeWorkbookError(f"Unexpected root element in relationship part {name!r}")
        for relationship in root:
            if relationship.tag != f"{{{RELATIONSHIPS_NS}}}Relationship":
                continue
            target = relationship.get("Target", "")
            mode = relationship.get("TargetMode", "Internal")
            relationship_type = relationship.get("Type", "")
            relationship_name = relationship_type.rsplit("/", maxsplit=1)[-1].lower()
            if relationship_name in VBA_RELATIONSHIP_NAMES:
                has_vba_relationship = True
            if mode == "External":
                external.append(target)
                continue
            resolved = _resolve_internal_target(name, target)
            if resolved not in infos:
                raise UnsafeWorkbookError(
                    f"Internal relationship from {name!r} points to missing part {resolved!r}"
                )
    return tuple(sorted(set(external))), has_vba_relationship


def _inspect_content_types(
    archive: zipfile.ZipFile,
    infos: dict[str, zipfile.ZipInfo],
    limits: PackageLimits,
) -> tuple[str | None, bool]:
    """Return the workbook format declared by OOXML and any macro content marker."""

    part_name = "[Content_Types].xml"
    root = parse_xml_part(_read_bounded(archive, infos[part_name], limits.max_xml_bytes), part_name)
    if root.tag != f"{{{CONTENT_TYPES_NS}}}Types":
        raise UnsafeWorkbookError("Unexpected root element in '[Content_Types].xml'")

    overrides: dict[str, set[str]] = {}
    defaults: dict[str, set[str]] = {}
    has_macro_content_type = False
    for declaration in root:
        content_type = declaration.get("ContentType", "").strip().lower()
        if "vbaproject" in content_type or "macroenabled" in content_type:
            has_macro_content_type = True
        if declaration.tag == f"{{{CONTENT_TYPES_NS}}}Override":
            declared_part = declaration.get("PartName", "").strip().lstrip("/")
            if declared_part:
                overrides.setdefault(declared_part, set()).add(content_type)
        elif declaration.tag == f"{{{CONTENT_TYPES_NS}}}Default":
            extension = declaration.get("Extension", "").strip().lstrip(".").lower()
            if extension:
                defaults.setdefault(extension, set()).add(content_type)

    for declared_part, content_types in overrides.items():
        if len(content_types) > 1:
            raise UnsafeWorkbookError(
                f"Part {declared_part!r} has conflicting Override content-type declarations"
            )
    for extension, content_types in defaults.items():
        if len(content_types) > 1:
            raise UnsafeWorkbookError(
                f"Extension {extension!r} has conflicting Default content-type declarations"
            )

    workbook_content_types = overrides.get("xl/workbook.xml")
    if workbook_content_types is None:
        workbook_content_types = defaults.get("xml")
    if not workbook_content_types:
        return None, has_macro_content_type

    workbook_content_type = next(iter(workbook_content_types))
    if workbook_content_type == XLSX_WORKBOOK_CONTENT_TYPE.lower():
        return "xlsx", has_macro_content_type
    if workbook_content_type == XLSM_WORKBOOK_CONTENT_TYPE.lower() or "macroenabled" in (
        workbook_content_type
    ):
        return "xlsm", True
    return None, has_macro_content_type


def _has_vba_part(part_names: frozenset[str]) -> bool:
    """Detect VBA project/signature parts even when their declarations are misleading."""

    for part_name in part_names:
        leaf = PurePosixPath(part_name).name.lower()
        if leaf in {"vbaproject.bin", "vbaprojectsignature.bin"}:
            return True
    return False


def inspect_package(path: Path, limits: PackageLimits | None = None) -> PackageInspection:
    """Validate extension, ZIP resources, XML, and internal relationships before use."""

    active_limits = limits or PackageLimits()
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        raise UsageError(f"Workbook does not exist or is not a file: {path}")
    suffix = resolved.suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise UsageError("WorkbookLens 2.2 accepts only .xlsx and read-only .xlsm inputs")
    file_size = resolved.stat().st_size
    if file_size > active_limits.max_file_bytes:
        raise UnsafeWorkbookError(
            f"Workbook is {file_size} bytes; configured limit is {active_limits.max_file_bytes}"
        )
    if not zipfile.is_zipfile(resolved):
        raise UnsafeWorkbookError("Input is not a valid ZIP-based Office Open XML package")

    try:
        with zipfile.ZipFile(resolved, "r") as archive:
            raw_infos = archive.infolist()
            if len(raw_infos) > active_limits.max_entries:
                raise UnsafeWorkbookError(
                    f"Workbook has {len(raw_infos)} ZIP entries; limit is {active_limits.max_entries}"
                )
            infos: dict[str, zipfile.ZipInfo] = {}
            total_compressed = 0
            total_uncompressed = 0
            for info in raw_infos:
                name = _canonical_member_name(info.filename)
                if name in infos:
                    raise UnsafeWorkbookError(f"ZIP package contains duplicate entry {name!r}")
                if info.flag_bits & 0x1:
                    raise UnsafeWorkbookError(f"Encrypted ZIP entry is unsupported: {name!r}")
                if _is_symlink(info):
                    raise UnsafeWorkbookError(f"Symbolic-link ZIP entry is forbidden: {name!r}")
                if info.file_size > active_limits.max_entry_uncompressed_bytes:
                    raise UnsafeWorkbookError(
                        f"ZIP entry {name!r} expands to {info.file_size} bytes, over the per-entry limit"
                    )
                if info.file_size and info.compress_size == 0:
                    raise UnsafeWorkbookError(
                        f"ZIP entry {name!r} has an impossible compression size"
                    )
                if info.compress_size:
                    ratio = info.file_size / info.compress_size
                    if ratio > active_limits.max_compression_ratio:
                        raise UnsafeWorkbookError(
                            f"ZIP entry {name!r} has suspicious compression ratio {ratio:.1f}"
                        )
                infos[name] = info
                total_compressed += info.compress_size
                total_uncompressed += info.file_size
                if total_uncompressed > active_limits.max_total_uncompressed_bytes:
                    raise UnsafeWorkbookError(
                        "Workbook exceeds the total uncompressed package limit"
                    )

            missing = REQUIRED_PARTS - infos.keys()
            if missing:
                raise UnsafeWorkbookError(
                    "Workbook is missing required OOXML parts: " + ", ".join(sorted(missing))
                )

            for name, info in infos.items():
                if name.endswith((".xml", ".rels")):
                    data = _read_bounded(archive, info, active_limits.max_xml_bytes)
                    parse_xml_part(data, name)
            content_format, macro_content_type = _inspect_content_types(
                archive,
                infos,
                active_limits,
            )
            external_relationships, macro_relationship = _inspect_relationships(
                archive,
                infos,
                active_limits,
            )
    except (zipfile.BadZipFile, RuntimeError, OSError) as exc:
        raise UnsafeWorkbookError(f"Unable to inspect workbook package safely: {exc}") from exc

    extension_format = suffix.lstrip(".")
    has_vba = macro_content_type or macro_relationship or _has_vba_part(frozenset(infos))
    detected_format = "xlsm" if has_vba else content_format
    format_mismatch = detected_format is None or detected_format != extension_format
    repairable = detected_format == "xlsx" and extension_format == "xlsx" and not format_mismatch
    # Existing callers use format == "xlsm" as the conservative read-only signal.
    effective_format = "xlsx" if repairable else "xlsm"

    return PackageInspection(
        path=resolved,
        format=effective_format,
        entry_count=len(infos),
        total_compressed_bytes=total_compressed,
        total_uncompressed_bytes=total_uncompressed,
        part_names=frozenset(infos),
        external_relationships=external_relationships,
        content_format=detected_format,
        extension_format=extension_format,
        has_vba=has_vba,
        format_mismatch=format_mismatch,
        repairable=repairable,
    )
