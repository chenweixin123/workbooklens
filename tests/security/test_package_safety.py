from __future__ import annotations

import stat
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

from workbooklens.exceptions import UnsafeWorkbookError, UsageError
from workbooklens.ooxml.safety import (
    PackageLimits,
    _canonical_member_name,
    inspect_package,
    parse_xml_part,
)

CONTENT_TYPES = b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
WORKBOOK = b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
ROOT_RELS = b"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def _minimal_package(
    path: Path,
    extras: list[tuple[str | zipfile.ZipInfo, bytes]] | None = None,
) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("xl/workbook.xml", WORKBOOK)
        for name, data in extras or []:
            archive.writestr(name, data)


def test_valid_openpyxl_package_passes(tmp_path: Path) -> None:
    workbook = Workbook()
    workbook.active["A1"] = "safe"  # type: ignore[index]
    path = tmp_path / "valid.xlsx"
    workbook.save(path)
    workbook.close()
    inspection = inspect_package(path)
    assert inspection.format == "xlsx"
    assert "xl/workbook.xml" in inspection.part_names


@pytest.mark.parametrize("name", ["../escape.xml", "/absolute.xml"])
def test_zip_member_path_traversal_is_rejected(tmp_path: Path, name: str) -> None:
    path = tmp_path / "traversal.xlsx"
    _minimal_package(path, [(name, b"x")])
    with pytest.raises(UnsafeWorkbookError, match="ZIP entry"):
        inspect_package(path)


def test_backslash_member_path_is_rejected_before_normalization() -> None:
    with pytest.raises(UnsafeWorkbookError, match="backslash"):
        _canonical_member_name("xl\\evil.xml")


def test_duplicate_member_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "duplicate.xlsx"
    with (
        pytest.warns(UserWarning, match="Duplicate name"),
        zipfile.ZipFile(path, "w") as archive,
    ):
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("xl/workbook.xml", WORKBOOK)
    with pytest.raises(UnsafeWorkbookError, match="duplicate"):
        inspect_package(path)


def test_symlink_member_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "symlink.xlsx"
    info = zipfile.ZipInfo("xl/link.xml")
    info.create_system = 3
    info.external_attr = (stat.S_IFLNK | 0o777) << 16
    _minimal_package(path, [(info, b"target")])
    with pytest.raises(UnsafeWorkbookError, match="Symbolic-link"):
        inspect_package(path)


def test_entry_count_limit_is_checked_before_parsing(tmp_path: Path) -> None:
    path = tmp_path / "entries.xlsx"
    _minimal_package(path, [(f"custom/{index}.bin", b"x") for index in range(5)])
    with pytest.raises(UnsafeWorkbookError, match="ZIP entries"):
        inspect_package(path, PackageLimits(max_entries=4))


def test_compression_ratio_limit_rejects_highly_repetitive_entry(tmp_path: Path) -> None:
    path = tmp_path / "ratio.xlsx"
    _minimal_package(path, [("custom/repetitive.bin", b"A" * 100_000)])
    with pytest.raises(UnsafeWorkbookError, match="compression ratio"):
        inspect_package(path, PackageLimits(max_compression_ratio=5))


def test_total_uncompressed_limit_is_enforced(tmp_path: Path) -> None:
    path = tmp_path / "total.xlsx"
    _minimal_package(path, [("custom/data.bin", b"0123456789" * 100)])
    with pytest.raises(UnsafeWorkbookError, match="total uncompressed"):
        inspect_package(path, PackageLimits(max_total_uncompressed_bytes=500))


def test_xml_dtd_and_entities_are_rejected() -> None:
    payload = b'<!DOCTYPE x [<!ENTITY boom "boom">]><x>&boom;</x>'
    with pytest.raises(UnsafeWorkbookError, match="DTD or entity"):
        parse_xml_part(payload, "custom.xml")


def test_xml_dtd_after_large_prefix_is_rejected() -> None:
    payload = b" " * 5000 + b'<!DOCTYPE x [<!ENTITY boom "boom">]><x>&boom;</x>'
    with pytest.raises(UnsafeWorkbookError, match="DTD or entity"):
        parse_xml_part(payload, "custom.xml")


def test_malformed_xml_is_rejected(tmp_path: Path) -> None:
    path = tmp_path / "malformed.xlsx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", b"<Types>")
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("xl/workbook.xml", WORKBOOK)
    with pytest.raises(UnsafeWorkbookError, match="Malformed XML"):
        inspect_package(path)


def test_internal_relationship_cannot_escape_package(tmp_path: Path) -> None:
    bad_rels = b"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="worksheet" Target="../../../outside.xml"/>
    </Relationships>"""
    path = tmp_path / "rels.xlsx"
    _minimal_package(path, [("xl/_rels/workbook.xml.rels", bad_rels)])
    with pytest.raises(UnsafeWorkbookError, match="escapes the package root"):
        inspect_package(path)


def test_external_relationship_is_recorded_but_not_opened(tmp_path: Path) -> None:
    external_rels = b"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="externalLink" Target="https://example.invalid/book.xlsx" TargetMode="External"/>
    </Relationships>"""
    path = tmp_path / "external.xlsx"
    _minimal_package(path, [("xl/_rels/workbook.xml.rels", external_rels)])
    inspection = inspect_package(path)
    assert inspection.external_relationships == ("https://example.invalid/book.xlsx",)


def test_wrong_extension_and_non_zip_fail_clearly(tmp_path: Path) -> None:
    wrong = tmp_path / "book.xls"
    wrong.write_bytes(b"not a workbook")
    with pytest.raises(UsageError, match=r"only \.xlsx"):
        inspect_package(wrong)
    fake = tmp_path / "book.xlsx"
    fake.write_bytes(b"not a zip")
    with pytest.raises(UnsafeWorkbookError, match="not a valid ZIP"):
        inspect_package(fake)


def test_file_size_limit_precedes_zip_parsing(tmp_path: Path) -> None:
    path = tmp_path / "large.xlsx"
    path.write_bytes(b"x" * 11)
    with pytest.raises(UnsafeWorkbookError, match="configured limit"):
        inspect_package(path, PackageLimits(max_file_bytes=10))
