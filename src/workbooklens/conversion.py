"""Best-effort local conversion of legacy binary Excel workbooks to OOXML."""

from __future__ import annotations

import base64
import contextlib
import copy
import ctypes
import importlib
import json
import ntpath
import os
import posixpath
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import zipfile
from ctypes import wintypes
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, Literal, assert_never, cast

from lxml import etree
from openpyxl.styles.numbers import BUILTIN_FORMATS

from workbooklens.exceptions import UsageError, WorkbookLensError
from workbooklens.ooxml.safety import PackageLimits, inspect_package, parse_xml_part

_OLE_COMPOUND_FILE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
_DEFAULT_TIMEOUT_SECONDS = 180
_MAX_FORMAT_RECORDS = 200_000
_MAX_FORMAT_MAP_BYTES = 64 * 1024 * 1024
_CELL_REFERENCE = re.compile(r"[A-Z]{1,3}[1-9][0-9]{0,6}\Z")
_PROCESS_DRAIN_TIMEOUT_SECONDS = 5
_MAX_EXCEL_IDENTITY_BYTES = 4096
_EXCEL_STARTUP_TIMEOUT_SECONDS = 30.0
_EXCEL_STARTUP_POLL_SECONDS = 0.05
_EXCEL_CANDIDATE_DISCOVERY_SECONDS = 1.0
_EXCEL_PROCESS_QUERY_TIMEOUT_SECONDS = 5
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_PROCESS_TERMINATE = 0x0001
_SYNCHRONIZE = 0x00100000
_WAIT_OBJECT_0 = 0x00000000
_WAIT_TIMEOUT = 0x00000102
_WAIT_FAILED = 0xFFFFFFFF
_ERROR_INVALID_PARAMETER = 87
_WINDOWS_TICKS_PER_SECOND = 10_000_000
_WINDOWS_EPOCH = datetime(1601, 1, 1, tzinfo=UTC)
_SIGTERM = int(getattr(signal, "SIGTERM", 15))
_SIGKILL = int(getattr(signal, "SIGKILL", 9))
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

_EXCEL_CONVERSION_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$inputPath = [Environment]::GetEnvironmentVariable('WORKBOOKLENS_XLS_INPUT', 'Process')
$outputPath = [Environment]::GetEnvironmentVariable('WORKBOOKLENS_XLSX_OUTPUT', 'Process')
$formatMapPath = [Environment]::GetEnvironmentVariable('WORKBOOKLENS_FORMAT_MAP', 'Process')
$excelIdentityPath = [Environment]::GetEnvironmentVariable(
    'WORKBOOKLENS_EXCEL_IDENTITY',
    'Process'
)
$excelGoPath = [Environment]::GetEnvironmentVariable('WORKBOOKLENS_EXCEL_GO', 'Process')
if (
    [string]::IsNullOrWhiteSpace($inputPath) -or
    [string]::IsNullOrWhiteSpace($outputPath) -or
    [string]::IsNullOrWhiteSpace($formatMapPath) -or
    [string]::IsNullOrWhiteSpace($excelIdentityPath) -or
    [string]::IsNullOrWhiteSpace($excelGoPath)
) {
    throw 'WorkbookLens did not provide conversion paths.'
}

$excelIdentityTempPath = "$excelIdentityPath.tmp"
$excel = $null
$workbook = $null
$calculationWorkbook = $null
$ownsExcelProcess = $false
$existingExcelPids = [System.Collections.Generic.HashSet[int]]::new()
foreach ($runningExcel in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
    [void]$existingExcelPids.Add([int]$runningExcel.Id)
}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class WorkbookLensNativeMethods {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
$script:formatRecords = New-Object System.Collections.ArrayList
function Release-ComObject($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
function Materialize-NumberFormats($sourceWorkbook) {
    $worksheets = $null
    try {
        $worksheets = $sourceWorkbook.Worksheets
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $worksheet = $null
            try {
                $worksheet = $worksheets.Item($sheetIndex)
                foreach ($cellType in @(2, -4123)) {
                    $usedRange = $null
                    $matchingCells = $null
                    $areas = $null
                    try {
                        $usedRange = $worksheet.UsedRange
                        try {
                            $matchingCells = $usedRange.SpecialCells($cellType)
                        } catch {
                            continue
                        }
                        $areas = $matchingCells.Areas
                        for ($areaIndex = 1; $areaIndex -le $areas.Count; $areaIndex++) {
                            $area = $null
                            $areaCells = $null
                            try {
                                $area = $areas.Item($areaIndex)
                                $areaCells = $area.Cells
                                for ($cellIndex = 1; $cellIndex -le $areaCells.Count; $cellIndex++) {
                                    $cell = $null
                                    try {
                                        $cell = $areaCells.Item($cellIndex)
                                        $format = [string]$cell.NumberFormatLocal
                                        if (
                                            -not [string]::IsNullOrWhiteSpace($format) -and
                                            $format -notin @('General', 'G/通用格式', '通用格式')
                                        ) {
                                            if ($script:formatRecords.Count -ge 200000) {
                                                throw 'Workbook has too many formatted non-empty cells to convert safely.'
                                            }
                                            [void]$script:formatRecords.Add([pscustomobject]@{
                                                sheet_index = $sheetIndex
                                                cell = [string]$cell.Address($false, $false)
                                                number_format = [string]$cell.NumberFormat
                                            })
                                        }
                                    } finally {
                                        Release-ComObject $cell
                                    }
                                }
                            } finally {
                                Release-ComObject $areaCells
                                Release-ComObject $area
                            }
                        }
                    } finally {
                        Release-ComObject $areas
                        Release-ComObject $matchingCells
                        Release-ComObject $usedRange
                    }
                }
            } finally {
                Release-ComObject $worksheet
            }
        }
    } finally {
        Release-ComObject $worksheets
    }
}
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelPid = 0
    $windowThread = [WorkbookLensNativeMethods]::GetWindowThreadProcessId(
        [IntPtr]([long]$excel.Hwnd),
        [ref]$excelPid
    )
    if ($windowThread -eq 0 -or $excelPid -eq 0) {
        throw 'WorkbookLens could not identify the Microsoft Excel process safely.'
    }
    if ($existingExcelPids.Contains([int]$excelPid)) {
        throw 'Microsoft Excel reused an existing user process; conversion was refused.'
    }
    $ownsExcelProcess = $true
    $ownedExcelProcessInfo = Get-Process -Id ([int]$excelPid) -ErrorAction Stop
    try {
        $startTimeUtc = $ownedExcelProcessInfo.StartTime.ToUniversalTime()
        $normalizedExecutablePath = [IO.Path]::GetFullPath(
            [string]$ownedExcelProcessInfo.Path
        ).ToLowerInvariant()
        if (
            [string]::IsNullOrWhiteSpace($normalizedExecutablePath) -or
            [IO.Path]::GetFileName($normalizedExecutablePath).ToUpperInvariant() -ne
                'EXCEL.EXE'
        ) {
            throw 'WorkbookLens could not bind the Microsoft Excel executable path safely.'
        }
        $identityJson = ConvertTo-Json -InputObject ([ordered]@{
            process_id = [int]$excelPid
            creation_utc = $startTimeUtc.ToString('o')
            creation_filetime = [string]$startTimeUtc.ToFileTimeUtc()
            session_id = [int]$ownedExcelProcessInfo.SessionId
            normalized_executable_path = $normalizedExecutablePath
        }) -Compress
    } finally {
        $ownedExcelProcessInfo.Dispose()
    }
    [IO.File]::WriteAllText(
        $excelIdentityTempPath,
        $identityJson,
        [Text.UTF8Encoding]::new($false)
    )
    [IO.File]::Move($excelIdentityTempPath, $excelIdentityPath)
    $goDeadline = [DateTime]::UtcNow.AddSeconds(35)
    while (-not [IO.File]::Exists($excelGoPath)) {
        if ([DateTime]::UtcNow -ge $goDeadline) {
            throw 'WorkbookLens did not confirm ownership of the Microsoft Excel process.'
        }
        Start-Sleep -Milliseconds 25
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $calculationWorkbook = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $excel.CalculateBeforeSave = $false
    $calculationWorkbook.Close($false)
    Release-ComObject $calculationWorkbook
    $calculationWorkbook = $null
    $workbook = $excel.Workbooks.Open($inputPath, 0, $true)
    $workbook.CheckCompatibility = $false
    Materialize-NumberFormats $workbook
    $workbook.SaveAs($outputPath, 51, $null, $null, $false, $false, 1, 2, $false, $null, $null, $true)
    $formatJson = if ($script:formatRecords.Count -eq 0) {
        '[]'
    } else {
        ConvertTo-Json -InputObject ($script:formatRecords.ToArray()) -Compress -Depth 3
    }
    [IO.File]::WriteAllText(
        $formatMapPath,
        $formatJson,
        [Text.UTF8Encoding]::new($false)
    )
} finally {
    if ([IO.File]::Exists($excelIdentityTempPath)) {
        try { [IO.File]::Delete($excelIdentityTempPath) } catch {}
    }
    if ($null -ne $calculationWorkbook) {
        try { $calculationWorkbook.Close($false) } catch {}
        Release-ComObject $calculationWorkbook
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        if ($ownsExcelProcess) {
            try { $excel.Quit() } catch {}
        }
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""

_EXCEL_PROCESS_QUERY_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$caller = Get-Process -Id $PID
$records = New-Object System.Collections.ArrayList
foreach ($item in @(Get-CimInstance Win32_Process -Filter "Name = 'EXCEL.EXE'")) {
    $process = $null
    $creationTime = $null
    $creationFileTime = $null
    $normalizedExecutablePath = $null
    $mainWindowHandle = $null
    try {
        $process = Get-Process -Id ([int]$item.ProcessId) -ErrorAction Stop
        $startTimeUtc = $process.StartTime.ToUniversalTime()
        $creationTime = $startTimeUtc.ToString('o')
        $creationFileTime = [string]$startTimeUtc.ToFileTimeUtc()
        if (-not [string]::IsNullOrWhiteSpace([string]$item.ExecutablePath)) {
            $normalizedExecutablePath = [IO.Path]::GetFullPath(
                [string]$item.ExecutablePath
            ).ToLowerInvariant()
        }
        $mainWindowHandle = [int64]$process.MainWindowHandle
    } catch {}
    [void]$records.Add([pscustomobject]@{
        process_id = [int]$item.ProcessId
        session_id = if ($null -eq $item.SessionId) { $null } else { [int]$item.SessionId }
        creation_utc = $creationTime
        creation_filetime = $creationFileTime
        normalized_executable_path = $normalizedExecutablePath
        command_line = if ($null -eq $item.CommandLine) { $null } else { [string]$item.CommandLine }
        main_window_handle = $mainWindowHandle
    })
}
[pscustomobject]@{
    caller_session_id = [int]$caller.SessionId
    processes = $records.ToArray()
} | ConvertTo-Json -Compress -Depth 4
"""


@dataclass(frozen=True, slots=True)
class ConversionProvider:
    """A locally installed application capable of opening legacy ``.xls`` files."""

    kind: Literal["excel", "libreoffice"]
    label: str
    runner: Path


@dataclass(frozen=True, slots=True)
class ConversionResult:
    """Verified OOXML output and the local application that produced it."""

    output: Path
    provider: ConversionProvider


@dataclass(frozen=True, slots=True)
class _NumberFormatRecord:
    sheet_index: int
    cell: str
    number_format: str


@dataclass(frozen=True, slots=True)
class _ExcelProcessIdentity:
    process_id: int
    creation_utc: str
    creation_filetime: int
    session_id: int
    normalized_executable_path: str


@dataclass(frozen=True, slots=True)
class _ExcelProcessRecord:
    process_id: int
    identity: _ExcelProcessIdentity | None
    creation_time: datetime | None
    command_line: str | None
    main_window_handle: int | None


@dataclass(frozen=True, slots=True)
class _ExcelProcessSnapshot:
    caller_session_id: int
    processes: tuple[_ExcelProcessRecord, ...]


@dataclass(frozen=True, slots=True)
class _OwnedExcelProcess:
    identity: _ExcelProcessIdentity
    handle: int


class _ProviderFailure(Exception):
    pass


class _ExcelStartupTimeout(Exception):
    pass


def _excel_com_registered() -> bool:
    if sys.platform != "win32":
        return False
    try:
        registry = importlib.import_module("winreg")
        with registry.OpenKey(registry.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID") as key:
            return bool(registry.QueryValueEx(key, "")[0])
    except (ImportError, OSError):
        return False


def _existing_path(candidates: list[str | Path | None]) -> Path | None:
    for candidate in candidates:
        if candidate is None:
            continue
        path = Path(candidate).expanduser()
        if path.is_file():
            return path.resolve()
    return None


def _find_powershell() -> Path | None:
    system_root = os.environ.get("SYSTEMROOT")
    system_powershell = (
        Path(system_root) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
        if system_root
        else None
    )
    return _existing_path(
        [
            shutil.which("powershell.exe"),
            system_powershell,
            shutil.which("pwsh.exe"),
        ]
    )


def _find_libreoffice() -> Path | None:
    program_files = os.environ.get("PROGRAMFILES")
    program_files_x86 = os.environ.get("PROGRAMFILES(X86)")
    local_app_data = os.environ.get("LOCALAPPDATA")
    return _existing_path(
        [
            shutil.which("soffice.exe" if sys.platform == "win32" else "soffice"),
            shutil.which("libreoffice"),
            Path(program_files) / "LibreOffice" / "program" / "soffice.exe"
            if program_files
            else None,
            Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe"
            if program_files_x86
            else None,
            Path(local_app_data) / "Programs" / "LibreOffice" / "program" / "soffice.exe"
            if local_app_data
            else None,
        ]
    )


def available_conversion_providers() -> tuple[ConversionProvider, ...]:
    """Return usable local converters in fidelity-preference order."""

    providers: list[ConversionProvider] = []
    powershell = _find_powershell()
    if powershell is not None and _excel_com_registered():
        providers.append(ConversionProvider("excel", "Microsoft Excel", powershell))
    libreoffice = _find_libreoffice()
    if libreoffice is not None:
        providers.append(ConversionProvider("libreoffice", "LibreOffice", libreoffice))
    return tuple(providers)


def _process_detail(completed: subprocess.CompletedProcess[str]) -> str:
    output = "\n".join(
        part.strip() for part in (completed.stderr, completed.stdout) if part.strip()
    )
    if not output:
        return f"process exited with code {completed.returncode}"
    return " ".join(output.split())[-600:]


def _load_number_format_records(path: Path) -> tuple[_NumberFormatRecord, ...]:
    try:
        if path.is_symlink() or not path.is_file() or path.stat().st_size > _MAX_FORMAT_MAP_BYTES:
            raise _ProviderFailure("Excel returned missing, unsafe, or excessive format metadata")
        payload: Any = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise _ProviderFailure(f"could not read Excel number-format metadata: {exc}") from exc
    if isinstance(payload, dict):
        payload = [payload]
    if not isinstance(payload, list) or len(payload) > _MAX_FORMAT_RECORDS:
        raise _ProviderFailure("Excel returned invalid or excessive number-format metadata")
    records: dict[tuple[int, str], _NumberFormatRecord] = {}
    for raw_record in payload:
        if not isinstance(raw_record, dict):
            raise _ProviderFailure("Excel returned a malformed number-format record")
        sheet_index = raw_record.get("sheet_index")
        cell = raw_record.get("cell")
        number_format = raw_record.get("number_format")
        if isinstance(sheet_index, bool) or not isinstance(sheet_index, int) or sheet_index < 1:
            raise _ProviderFailure("Excel returned an invalid worksheet index")
        if not isinstance(cell, str) or _CELL_REFERENCE.fullmatch(cell) is None:
            raise _ProviderFailure("Excel returned an invalid formatted-cell reference")
        if (
            not isinstance(number_format, str)
            or not number_format
            or "\x00" in number_format
            or len(number_format) > 1024
        ):
            raise _ProviderFailure("Excel returned an invalid number-format code")
        record = _NumberFormatRecord(sheet_index, cell, number_format)
        key = (sheet_index, cell)
        existing = records.get(key)
        if existing is not None and existing != record:
            raise _ProviderFailure("Excel returned conflicting number formats for one cell")
        records[key] = record
    return tuple(records[key] for key in sorted(records))


def _xml_namespace(root: etree._Element) -> str:
    namespace = etree.QName(root).namespace
    if namespace is None:
        raise _ProviderFailure("Excel generated OOXML without a namespace")
    return namespace


def _qname(namespace: str, local_name: str) -> str:
    return f"{{{namespace}}}{local_name}"


def _direct_child(root: etree._Element, local_name: str) -> etree._Element | None:
    return next(
        (child for child in root if etree.QName(child).localname == local_name),
        None,
    )


def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def _worksheet_parts_by_index(archive: zipfile.ZipFile) -> list[str]:
    workbook_part = "xl/workbook.xml"
    relationships_part = "xl/_rels/workbook.xml.rels"
    try:
        workbook_root = parse_xml_part(archive.read(workbook_part), workbook_part)
        relationships_root = parse_xml_part(
            archive.read(relationships_part),
            relationships_part,
        )
    except KeyError as exc:
        raise _ProviderFailure("Excel output is missing workbook relationships") from exc
    relationships: dict[str, tuple[str, str]] = {}
    for relationship in relationships_root:
        if etree.QName(relationship).localname != "Relationship":
            continue
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        relationship_type = relationship.get("Type", "")
        if relationship_id and target:
            relationships[relationship_id] = (relationship_type, target)
    parts: list[str] = []
    for sheet in workbook_root.iter():
        if etree.QName(sheet).localname != "sheet":
            continue
        relationship_id = cast(
            str | None,
            next(
                (
                    value
                    for attribute, value in sheet.attrib.items()
                    if etree.QName(attribute).localname == "id"
                ),
                None,
            ),
        )
        if relationship_id is None or relationship_id not in relationships:
            raise _ProviderFailure("Excel output contains a malformed sheet relationship")
        relationship_type, target = relationships[relationship_id]
        if relationship_type.endswith("/chartsheet"):
            continue
        if not relationship_type.endswith("/worksheet"):
            raise _ProviderFailure("Excel output contains an unsupported sheet relationship")
        candidate = target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
        normalized = posixpath.normpath(candidate)
        if normalized in {"", ".", ".."} or normalized.startswith("../") or "\\" in normalized:
            raise _ProviderFailure("Excel output contains an unsafe sheet relationship")
        if normalized not in archive.namelist():
            raise _ProviderFailure("Excel output is missing a worksheet part")
        parts.append(normalized)
    return parts


def _rewrite_package(path: Path, modified: dict[str, bytes]) -> None:
    with tempfile.NamedTemporaryFile(
        prefix=f".{path.stem}.", suffix=".xlsx", dir=path.parent, delete=False
    ) as temporary_handle:
        temporary_path = Path(temporary_handle.name)
    try:
        with (
            zipfile.ZipFile(path, "r") as source,
            zipfile.ZipFile(temporary_path, "w", allowZip64=True) as target,
        ):
            target.comment = source.comment
            for info in source.infolist():
                copied_info = copy.copy(info)
                replacement = modified.get(info.filename)
                if replacement is not None:
                    target.writestr(copied_info, replacement)
                    continue
                with (
                    source.open(info, "r") as source_handle,
                    target.open(copied_info, "w") as target_handle,
                ):
                    shutil.copyfileobj(source_handle, target_handle, length=1024 * 1024)
        os.replace(temporary_path, path)
    except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
        raise _ProviderFailure(f"could not restore Excel number formats safely: {exc}") from exc
    finally:
        temporary_path.unlink(missing_ok=True)


def _restore_number_formats(output: Path, records: tuple[_NumberFormatRecord, ...]) -> None:
    if not records:
        return
    try:
        with zipfile.ZipFile(output, "r") as archive:
            worksheet_parts = _worksheet_parts_by_index(archive)
            styles_part = "xl/styles.xml"
            try:
                styles_root = parse_xml_part(archive.read(styles_part), styles_part)
            except KeyError as exc:
                raise _ProviderFailure("Excel output is missing styles.xml") from exc
            cell_xfs = _direct_child(styles_root, "cellXfs")
            if cell_xfs is None or not len(cell_xfs):
                raise _ProviderFailure("Excel output has no cell style table")
            num_fmts = _direct_child(styles_root, "numFmts")
            format_to_id: dict[str, int] = {}
            id_to_format: dict[int, str] = {}
            if num_fmts is not None:
                for element in num_fmts:
                    if etree.QName(element).localname != "numFmt":
                        continue
                    try:
                        number_format_id = int(element.get("numFmtId", ""))
                    except ValueError as exc:
                        raise _ProviderFailure(
                            "Excel output has an invalid custom format id"
                        ) from exc
                    format_code = element.get("formatCode")
                    if format_code:
                        format_to_id.setdefault(format_code, number_format_id)
                        id_to_format[number_format_id] = format_code
            next_format_id = max({163, *id_to_format}) + 1
            by_part: dict[str, list[_NumberFormatRecord]] = {}
            for record in records:
                if record.sheet_index > len(worksheet_parts):
                    raise _ProviderFailure("Excel format metadata references a missing worksheet")
                by_part.setdefault(worksheet_parts[record.sheet_index - 1], []).append(record)

            modified: dict[str, bytes] = {}
            style_cache: dict[tuple[int, int], int] = {}
            styles_changed = False
            for part, part_records in by_part.items():
                sheet_root = parse_xml_part(archive.read(part), part)
                cells = {
                    reference: element
                    for element in sheet_root.iter()
                    if etree.QName(element).localname == "c"
                    and (reference := element.get("r")) is not None
                }
                sheet_changed = False
                for record in part_records:
                    cell = cells.get(record.cell)
                    if cell is None:
                        raise _ProviderFailure(
                            f"Excel format metadata references missing cell {record.cell}"
                        )
                    try:
                        style_index = int(cell.get("s", "0"))
                    except ValueError as exc:
                        raise _ProviderFailure(
                            "Excel output has an invalid cell style index"
                        ) from exc
                    if not 0 <= style_index < len(cell_xfs):
                        raise _ProviderFailure("Excel output cell references a missing style")
                    source_xf = cell_xfs[style_index]
                    try:
                        current_format_id = int(source_xf.get("numFmtId", "0"))
                    except ValueError as exc:
                        raise _ProviderFailure(
                            "Excel output style has an invalid number format id"
                        ) from exc
                    current_format = id_to_format.get(
                        current_format_id,
                        BUILTIN_FORMATS.get(current_format_id),
                    )
                    if current_format == record.number_format:
                        continue
                    replacement_format_id = format_to_id.get(record.number_format)
                    if replacement_format_id is None:
                        if next_format_id > 65535:
                            raise _ProviderFailure(
                                "Excel output exhausted custom number format ids"
                            )
                        replacement_format_id = next_format_id
                        next_format_id += 1
                        if num_fmts is None:
                            namespace = _xml_namespace(styles_root)
                            num_fmts = etree.Element(_qname(namespace, "numFmts"), count="0")
                            styles_root.insert(0, num_fmts)
                        format_element = etree.Element(
                            _qname(_xml_namespace(styles_root), "numFmt"),
                            numFmtId=str(replacement_format_id),
                            formatCode=record.number_format,
                        )
                        num_fmts.append(format_element)
                        num_fmts.set("count", str(len(num_fmts)))
                        format_to_id[record.number_format] = replacement_format_id
                        id_to_format[replacement_format_id] = record.number_format
                        styles_changed = True
                    cache_key = (style_index, replacement_format_id)
                    replacement_style_index = style_cache.get(cache_key)
                    if replacement_style_index is None:
                        replacement_xf = copy.deepcopy(source_xf)
                        replacement_xf.set("numFmtId", str(replacement_format_id))
                        replacement_xf.set("applyNumberFormat", "1")
                        cell_xfs.append(replacement_xf)
                        cell_xfs.set("count", str(len(cell_xfs)))
                        replacement_style_index = len(cell_xfs) - 1
                        style_cache[cache_key] = replacement_style_index
                        styles_changed = True
                    cell.set("s", str(replacement_style_index))
                    sheet_changed = True
                if sheet_changed:
                    modified[part] = _serialize_xml(sheet_root)
            if styles_changed:
                modified[styles_part] = _serialize_xml(styles_root)
    except (OSError, RuntimeError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        raise _ProviderFailure(f"could not inspect Excel number formats safely: {exc}") from exc
    if modified:
        _rewrite_package(output, modified)


def _creation_flags() -> int:
    return int(getattr(subprocess, "CREATE_NO_WINDOW", 0))


def _windows_taskkill_executable() -> Path | None:
    system_root = os.environ.get("SYSTEMROOT")
    if not system_root:
        return None
    candidate = Path(system_root) / "System32" / "taskkill.exe"
    return candidate if candidate.is_file() else None


def _taskkill_process_tree(process_id: int) -> None:
    if sys.platform != "win32" or process_id <= 0:
        return
    executable = _windows_taskkill_executable()
    if executable is None:
        return
    with contextlib.suppress(OSError, subprocess.TimeoutExpired):
        subprocess.run(  # noqa: S603
            [str(executable), "/PID", str(process_id), "/T", "/F"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=_PROCESS_DRAIN_TIMEOUT_SECONDS,
            creationflags=_creation_flags(),
        )


def _filetime_to_creation_utc(filetime: int) -> str | None:
    if filetime <= 0:
        return None
    whole_seconds, fractional_ticks = divmod(filetime, _WINDOWS_TICKS_PER_SECOND)
    try:
        created = _WINDOWS_EPOCH + timedelta(seconds=whole_seconds)
    except OverflowError:
        return None
    return created.strftime("%Y-%m-%dT%H:%M:%S") + f".{fractional_ticks:07d}Z"


def _normalize_windows_executable_path(value: str) -> str | None:
    if not value or "\x00" in value or len(value) > 32767:
        return None
    normalized = ntpath.normcase(ntpath.normpath(value.strip()))
    if not ntpath.isabs(normalized) or ntpath.basename(normalized).casefold() != "excel.exe":
        return None
    return normalized


def _load_excel_process_identity(path: Path) -> _ExcelProcessIdentity | None:
    try:
        if (
            path.is_symlink()
            or not path.is_file()
            or not 1 <= path.stat().st_size <= _MAX_EXCEL_IDENTITY_BYTES
        ):
            return None
        payload: Any = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None
    required_fields = {
        "process_id",
        "creation_utc",
        "creation_filetime",
        "session_id",
        "normalized_executable_path",
    }
    if not isinstance(payload, dict) or set(payload) != required_fields:
        return None
    process_id = payload.get("process_id")
    session_id = payload.get("session_id")
    creation_utc = payload.get("creation_utc")
    raw_filetime = payload.get("creation_filetime")
    raw_path = payload.get("normalized_executable_path")
    if (
        isinstance(process_id, bool)
        or not isinstance(process_id, int)
        or not 1 <= process_id <= 0x7FFFFFFF
        or isinstance(session_id, bool)
        or not isinstance(session_id, int)
        or not 0 <= session_id <= 0xFFFFFFFF
        or not isinstance(creation_utc, str)
        or not isinstance(raw_filetime, str)
        or re.fullmatch(r"[1-9][0-9]{0,18}", raw_filetime) is None
        or not isinstance(raw_path, str)
    ):
        return None
    creation_filetime = int(raw_filetime)
    expected_creation_utc = _filetime_to_creation_utc(creation_filetime)
    normalized_path = _normalize_windows_executable_path(raw_path)
    if (
        expected_creation_utc is None
        or creation_utc != expected_creation_utc
        or normalized_path is None
        or raw_path != normalized_path
    ):
        return None
    return _ExcelProcessIdentity(
        process_id=process_id,
        creation_utc=creation_utc,
        creation_filetime=creation_filetime,
        session_id=session_id,
        normalized_executable_path=normalized_path,
    )


class _WindowsFileTime(ctypes.Structure):
    _fields_ = [
        ("low", wintypes.DWORD),
        ("high", wintypes.DWORD),
    ]


def _windows_kernel32() -> Any:
    if sys.platform != "win32":
        raise _ProviderFailure("Microsoft Excel process handles require Windows")
    loader = getattr(ctypes, "WinDLL", None)
    if loader is None:
        raise _ProviderFailure("Windows process-handle APIs are unavailable")
    return loader("kernel32", use_last_error=True)


def _windows_last_error() -> int:
    getter = getattr(ctypes, "get_last_error", None)
    return int(getter()) if getter is not None else 0


def _open_windows_process_handle(process_id: int) -> int | None:
    kernel32 = _windows_kernel32()
    open_process = kernel32.OpenProcess
    open_process.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    open_process.restype = wintypes.HANDLE
    raw_handle = open_process(
        _PROCESS_QUERY_LIMITED_INFORMATION | _PROCESS_TERMINATE | _SYNCHRONIZE,
        False,
        process_id,
    )
    if not raw_handle:
        error = _windows_last_error()
        if error == _ERROR_INVALID_PARAMETER:
            return None
        raise _ProviderFailure(f"could not open the Excel process safely (Windows error {error})")
    if isinstance(raw_handle, int):
        return raw_handle
    value = getattr(raw_handle, "value", None)
    if not isinstance(value, int) or value <= 0:
        raise _ProviderFailure("Windows returned an invalid Excel process handle")
    return value


def _close_windows_process_handle(handle: int) -> None:
    if handle <= 0 or sys.platform != "win32":
        return
    with contextlib.suppress(OSError, _ProviderFailure):
        kernel32 = _windows_kernel32()
        close_handle = kernel32.CloseHandle
        close_handle.argtypes = [wintypes.HANDLE]
        close_handle.restype = wintypes.BOOL
        close_handle(wintypes.HANDLE(handle))


def _read_windows_process_handle_identity(
    handle: int,
    process_id: int,
) -> _ExcelProcessIdentity:
    kernel32 = _windows_kernel32()
    get_process_id = kernel32.GetProcessId
    get_process_id.argtypes = [wintypes.HANDLE]
    get_process_id.restype = wintypes.DWORD
    handle_process_id = int(get_process_id(wintypes.HANDLE(handle)))
    if handle_process_id == 0:
        error = _windows_last_error()
        raise _ProviderFailure(f"could not bind the Excel handle to a PID (Windows error {error})")
    if handle_process_id != process_id:
        raise _ProviderFailure("Windows returned an Excel handle for an unexpected PID")
    get_process_times = kernel32.GetProcessTimes
    get_process_times.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(_WindowsFileTime),
        ctypes.POINTER(_WindowsFileTime),
        ctypes.POINTER(_WindowsFileTime),
        ctypes.POINTER(_WindowsFileTime),
    ]
    get_process_times.restype = wintypes.BOOL
    creation = _WindowsFileTime()
    exit_time = _WindowsFileTime()
    kernel_time = _WindowsFileTime()
    user_time = _WindowsFileTime()
    if not get_process_times(
        wintypes.HANDLE(handle),
        ctypes.byref(creation),
        ctypes.byref(exit_time),
        ctypes.byref(kernel_time),
        ctypes.byref(user_time),
    ):
        error = _windows_last_error()
        raise _ProviderFailure(
            f"could not read the Excel creation time safely (Windows error {error})"
        )
    creation_filetime = (int(creation.high) << 32) | int(creation.low)
    creation_utc = _filetime_to_creation_utc(creation_filetime)
    if creation_utc is None:
        raise _ProviderFailure("Windows returned an invalid Excel creation time")

    query_image = kernel32.QueryFullProcessImageNameW
    query_image.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    query_image.restype = wintypes.BOOL
    image_buffer = ctypes.create_unicode_buffer(32768)
    image_length = wintypes.DWORD(len(image_buffer))
    if not query_image(
        wintypes.HANDLE(handle),
        0,
        image_buffer,
        ctypes.byref(image_length),
    ):
        error = _windows_last_error()
        raise _ProviderFailure(
            f"could not read the Excel executable path safely (Windows error {error})"
        )
    normalized_path = _normalize_windows_executable_path(image_buffer.value[: image_length.value])
    if normalized_path is None:
        raise _ProviderFailure("Windows returned an invalid Excel executable path")

    process_id_to_session_id = kernel32.ProcessIdToSessionId
    process_id_to_session_id.argtypes = [
        wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD),
    ]
    process_id_to_session_id.restype = wintypes.BOOL
    session_id = wintypes.DWORD()
    if not process_id_to_session_id(handle_process_id, ctypes.byref(session_id)):
        error = _windows_last_error()
        raise _ProviderFailure(f"could not read the Excel session safely (Windows error {error})")
    return _ExcelProcessIdentity(
        process_id=handle_process_id,
        creation_utc=creation_utc,
        creation_filetime=creation_filetime,
        session_id=int(session_id.value),
        normalized_executable_path=normalized_path,
    )


def _open_verified_excel_process(
    expected: _ExcelProcessIdentity,
    *,
    missing_is_clean: bool,
) -> _OwnedExcelProcess | None:
    handle = _open_windows_process_handle(expected.process_id)
    if handle is None:
        if missing_is_clean:
            return None
        raise _ProviderFailure("Excel exited before its process identity could be bound")
    try:
        actual = _read_windows_process_handle_identity(handle, expected.process_id)
    except _ProviderFailure:
        _close_windows_process_handle(handle)
        raise
    if actual != expected:
        _close_windows_process_handle(handle)
        raise _ProviderFailure(
            "Excel process identity changed; refusing to authorize or terminate that PID"
        )
    return _OwnedExcelProcess(expected, handle)


def _wait_windows_process_handle(handle: int, timeout_milliseconds: int) -> int:
    kernel32 = _windows_kernel32()
    wait_for_single_object = kernel32.WaitForSingleObject
    wait_for_single_object.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    wait_for_single_object.restype = wintypes.DWORD
    return int(
        wait_for_single_object(
            wintypes.HANDLE(handle),
            max(0, timeout_milliseconds),
        )
    )


def _terminate_excel_process(process: _OwnedExcelProcess) -> None:
    wait_result = _wait_windows_process_handle(process.handle, 0)
    if wait_result == _WAIT_OBJECT_0:
        return
    if wait_result == _WAIT_FAILED:
        raise _ProviderFailure(
            f"could not inspect the Excel process handle (Windows error {_windows_last_error()})"
        )
    if wait_result != _WAIT_TIMEOUT:
        raise _ProviderFailure("Windows returned an unexpected Excel handle wait result")
    kernel32 = _windows_kernel32()
    terminate_process = kernel32.TerminateProcess
    terminate_process.argtypes = [wintypes.HANDLE, wintypes.UINT]
    terminate_process.restype = wintypes.BOOL
    if not terminate_process(wintypes.HANDLE(process.handle), 1):
        terminate_error = _windows_last_error()
        if _wait_windows_process_handle(process.handle, 0) == _WAIT_OBJECT_0:
            return
        raise _ProviderFailure(
            f"could not terminate the owned Excel process (Windows error {terminate_error})"
        )
    wait_result = _wait_windows_process_handle(
        process.handle,
        _PROCESS_DRAIN_TIMEOUT_SECONDS * 1000,
    )
    if wait_result == _WAIT_TIMEOUT:
        raise _ProviderFailure("the owned Excel process did not terminate in time")
    if wait_result != _WAIT_OBJECT_0:
        raise _ProviderFailure(
            f"could not wait for the owned Excel process (Windows error {_windows_last_error()})"
        )


def _close_excel_process(process: _OwnedExcelProcess | None) -> None:
    if process is not None:
        _close_windows_process_handle(process.handle)


def _terminate_excel_identity_if_present(identity: _ExcelProcessIdentity) -> None:
    process = _open_verified_excel_process(identity, missing_is_clean=True)
    if process is None:
        return
    try:
        _terminate_excel_process(process)
    finally:
        _close_excel_process(process)


def _terminate_process_tree(process: subprocess.Popen[str]) -> None:
    if sys.platform == "win32":
        _taskkill_process_tree(process.pid)
    else:
        with contextlib.suppress(OSError):
            os.killpg(process.pid, _SIGTERM)
    try:
        process.communicate(timeout=_PROCESS_DRAIN_TIMEOUT_SECONDS)
        return
    except (OSError, subprocess.TimeoutExpired):
        pass
    if sys.platform == "win32":
        with contextlib.suppress(OSError):
            process.kill()
    else:
        with contextlib.suppress(OSError):
            os.killpg(process.pid, _SIGKILL)
    try:
        process.communicate(timeout=_PROCESS_DRAIN_TIMEOUT_SECONDS)
    except (OSError, subprocess.TimeoutExpired):
        with contextlib.suppress(OSError):
            process.kill()


def _query_excel_processes(runner: Path) -> _ExcelProcessSnapshot | None:
    encoded_script = base64.b64encode(_EXCEL_PROCESS_QUERY_SCRIPT.encode("utf-16le")).decode(
        "ascii"
    )
    try:
        completed = subprocess.run(  # noqa: S603
            [
                str(runner),
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-EncodedCommand",
                encoded_script,
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=_EXCEL_PROCESS_QUERY_TIMEOUT_SECONDS,
            creationflags=_creation_flags(),
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    if completed.returncode != 0:
        return None
    try:
        payload: Any = json.loads(completed.stdout.lstrip("\ufeff"))
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict):
        return None
    caller_session_id = payload.get("caller_session_id")
    if (
        isinstance(caller_session_id, bool)
        or not isinstance(caller_session_id, int)
        or caller_session_id < 0
    ):
        return None
    raw_processes = payload.get("processes")
    if raw_processes is None:
        raw_processes = []
    elif isinstance(raw_processes, dict):
        raw_processes = [raw_processes]
    if not isinstance(raw_processes, list):
        return None
    processes: list[_ExcelProcessRecord] = []
    for raw_process in raw_processes:
        if not isinstance(raw_process, dict):
            return None
        process_id = raw_process.get("process_id")
        if (
            isinstance(process_id, bool)
            or not isinstance(process_id, int)
            or not 1 <= process_id <= 0x7FFFFFFF
        ):
            return None
        identity: _ExcelProcessIdentity | None = None
        creation_time: datetime | None = None
        session_id_value = raw_process.get("session_id")
        creation_utc_value = raw_process.get("creation_utc")
        creation_filetime_value = raw_process.get("creation_filetime")
        executable_path_value = raw_process.get("normalized_executable_path")
        normalized_path = (
            _normalize_windows_executable_path(executable_path_value)
            if isinstance(executable_path_value, str)
            else None
        )
        creation_filetime = (
            int(creation_filetime_value)
            if isinstance(creation_filetime_value, str)
            and re.fullmatch(r"[1-9][0-9]{0,18}", creation_filetime_value)
            else None
        )
        expected_creation_utc = (
            _filetime_to_creation_utc(creation_filetime) if creation_filetime is not None else None
        )
        if (
            isinstance(session_id_value, int)
            and not isinstance(session_id_value, bool)
            and 0 <= session_id_value <= 0xFFFFFFFF
            and isinstance(creation_utc_value, str)
            and creation_utc_value == expected_creation_utc
            and creation_filetime is not None
            and normalized_path is not None
            and executable_path_value == normalized_path
        ):
            identity = _ExcelProcessIdentity(
                process_id=process_id,
                creation_utc=creation_utc_value,
                creation_filetime=creation_filetime,
                session_id=session_id_value,
                normalized_executable_path=normalized_path,
            )
            try:
                parsed_time = datetime.fromisoformat(creation_utc_value.replace("Z", "+00:00"))
            except ValueError:
                pass
            else:
                if parsed_time.tzinfo is not None:
                    creation_time = parsed_time.astimezone(UTC)
        command_line_value = raw_process.get("command_line")
        command_line = (
            command_line_value.strip()
            if isinstance(command_line_value, str) and command_line_value.strip()
            else None
        )
        main_window_handle_value = raw_process.get("main_window_handle")
        main_window_handle = (
            main_window_handle_value
            if isinstance(main_window_handle_value, int)
            and not isinstance(main_window_handle_value, bool)
            and main_window_handle_value >= 0
            else None
        )
        processes.append(
            _ExcelProcessRecord(
                process_id=process_id,
                identity=identity,
                creation_time=creation_time,
                command_line=command_line,
                main_window_handle=main_window_handle,
            )
        )
    return _ExcelProcessSnapshot(caller_session_id, tuple(processes))


def _snapshot_excel_process_ids(runner: Path) -> frozenset[int] | None:
    snapshot = _query_excel_processes(runner)
    if snapshot is None:
        return None
    return frozenset(process.process_id for process in snapshot.processes)


def _has_windows_command_argument(command_line: str, argument: str) -> bool:
    return (
        re.search(
            rf"(?<!\S){re.escape(argument)}(?!\S)",
            command_line,
            flags=re.IGNORECASE,
        )
        is not None
    )


def _find_unique_new_excel_automation_process(
    runner: Path,
    baseline_process_ids: frozenset[int] | None,
    launched_at: datetime,
) -> _ExcelProcessIdentity | None:
    if baseline_process_ids is None:
        return None
    observed_candidates: set[_ExcelProcessIdentity] = set()
    deadline = time.monotonic() + _EXCEL_CANDIDATE_DISCOVERY_SECONDS
    while True:
        snapshot = _query_excel_processes(runner)
        if snapshot is None:
            return None
        for process in snapshot.processes:
            if process.process_id in baseline_process_ids:
                continue
            if (
                process.identity is None
                or process.creation_time is None
                or process.command_line is None
                or process.main_window_handle is None
            ):
                return None
            if (
                process.identity.session_id == snapshot.caller_session_id
                and process.creation_time >= launched_at
                and process.main_window_handle == 0
                and _has_windows_command_argument(process.command_line, "/automation")
                and _has_windows_command_argument(process.command_line, "-Embedding")
            ):
                observed_candidates.add(process.identity)
        if len(observed_candidates) > 1:
            return None
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            break
        time.sleep(min(_EXCEL_STARTUP_POLL_SECONDS, remaining))
    if len(observed_candidates) != 1:
        return None
    return next(iter(observed_candidates))


def _wait_for_owned_excel_process_identity(
    process: subprocess.Popen[str],
    path: Path,
    command: list[str],
) -> _ExcelProcessIdentity:
    deadline = time.monotonic() + _EXCEL_STARTUP_TIMEOUT_SECONDS
    while True:
        identity = _load_excel_process_identity(path)
        if identity is not None:
            return identity
        if os.path.lexists(path):
            raise _ProviderFailure("Excel returned an invalid process-identity handshake")
        try:
            returncode = process.poll()
        except OSError as exc:
            raise _ProviderFailure(f"could not monitor PowerShell: {exc}") from exc
        if returncode is not None:
            try:
                stdout, stderr = process.communicate(timeout=_PROCESS_DRAIN_TIMEOUT_SECONDS)
            except (OSError, subprocess.TimeoutExpired):
                stdout, stderr = "", ""
            completed = subprocess.CompletedProcess(
                command,
                returncode,
                stdout or "",
                stderr or "",
            )
            raise _ProviderFailure(_process_detail(completed))
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise _ExcelStartupTimeout
        time.sleep(min(_EXCEL_STARTUP_POLL_SECONDS, remaining))


def _signal_excel_go(path: Path) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    try:
        temporary.write_bytes(b"go")
        os.replace(temporary, path)
    except OSError as exc:
        raise _ProviderFailure(f"could not confirm Excel process ownership: {exc}") from exc
    finally:
        with contextlib.suppress(OSError):
            temporary.unlink()


def _run_excel(
    provider: ConversionProvider,
    source: Path,
    output: Path,
    timeout_seconds: int,
) -> tuple[_NumberFormatRecord, ...]:
    if output.exists():
        raise _ProviderFailure(f"conversion output already exists: {output}")
    encoded_script = base64.b64encode(_EXCEL_CONVERSION_SCRIPT.encode("utf-16le")).decode("ascii")
    with tempfile.TemporaryDirectory(prefix=".workbooklens-excel-", dir=output.parent) as raw:
        format_map = Path(raw) / "number-formats.json"
        staged_output = Path(raw) / "converted.xlsx"
        excel_identity = Path(raw) / "excel-process.json"
        excel_go = Path(raw) / "excel.go"
        environment = os.environ.copy()
        environment["WORKBOOKLENS_XLS_INPUT"] = str(source)
        environment["WORKBOOKLENS_XLSX_OUTPUT"] = str(staged_output)
        environment["WORKBOOKLENS_FORMAT_MAP"] = str(format_map)
        environment["WORKBOOKLENS_EXCEL_IDENTITY"] = str(excel_identity)
        environment["WORKBOOKLENS_EXCEL_GO"] = str(excel_go)
        command = [
            str(provider.runner),
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-EncodedCommand",
            encoded_script,
        ]
        baseline_process_ids = _snapshot_excel_process_ids(provider.runner)
        launched_at = datetime.now(UTC)
        try:
            process = subprocess.Popen(  # noqa: S603
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
                creationflags=_creation_flags(),
            )
        except OSError as exc:
            raise _ProviderFailure(f"could not start PowerShell: {exc}") from exc
        try:
            expected_identity = _wait_for_owned_excel_process_identity(
                process,
                excel_identity,
                command,
            )
        except _ExcelStartupTimeout as exc:
            _terminate_process_tree(process)
            late_identity = _load_excel_process_identity(excel_identity)
            try:
                if late_identity is not None:
                    if (
                        baseline_process_ids is None
                        or late_identity.process_id not in baseline_process_ids
                    ):
                        _terminate_excel_identity_if_present(late_identity)
                elif os.path.lexists(excel_identity):
                    raise _ProviderFailure(
                        "Excel returned an invalid late process-identity handshake"
                    )
                else:
                    candidate = _find_unique_new_excel_automation_process(
                        provider.runner,
                        baseline_process_ids,
                        launched_at,
                    )
                    if candidate is not None:
                        _terminate_excel_identity_if_present(candidate)
            except _ProviderFailure as safety_error:
                raise safety_error from exc
            raise _ProviderFailure(
                f"Excel startup handshake timed out after "
                f"{int(_EXCEL_STARTUP_TIMEOUT_SECONDS)} seconds"
            ) from exc
        except _ProviderFailure:
            _terminate_process_tree(process)
            raise
        if (
            baseline_process_ids is not None
            and expected_identity.process_id in baseline_process_ids
        ):
            _terminate_process_tree(process)
            raise _ProviderFailure(
                "Microsoft Excel reused an existing user process; conversion was refused"
            )
        try:
            owned_excel = _open_verified_excel_process(
                expected_identity,
                missing_is_clean=False,
            )
        except _ProviderFailure:
            _terminate_process_tree(process)
            raise
        if owned_excel is None:
            _terminate_process_tree(process)
            raise _ProviderFailure("Excel exited before its process handle could be held")
        try:
            try:
                _signal_excel_go(excel_go)
            except _ProviderFailure:
                try:
                    _terminate_excel_process(owned_excel)
                finally:
                    _terminate_process_tree(process)
                raise
            try:
                stdout, stderr = process.communicate(timeout=timeout_seconds)
            except subprocess.TimeoutExpired as exc:
                try:
                    _terminate_excel_process(owned_excel)
                finally:
                    _terminate_process_tree(process)
                raise _ProviderFailure(f"timed out after {timeout_seconds} seconds") from exc
            except OSError as exc:
                try:
                    _terminate_excel_process(owned_excel)
                finally:
                    _terminate_process_tree(process)
                raise _ProviderFailure(f"could not monitor PowerShell: {exc}") from exc
            completed = subprocess.CompletedProcess(
                command,
                process.returncode if process.returncode is not None else -1,
                stdout or "",
                stderr or "",
            )
            if completed.returncode != 0:
                _terminate_excel_process(owned_excel)
                raise _ProviderFailure(_process_detail(completed))
            _terminate_excel_process(owned_excel)
            records = _load_number_format_records(format_map)
            if not staged_output.is_file():
                raise _ProviderFailure("Excel did not create the expected .xlsx file")
            if output.exists():
                raise _ProviderFailure(
                    "conversion output appeared while Excel was running; refusing to overwrite it"
                )
            staged_output.rename(output)
            return records
        finally:
            _close_excel_process(owned_excel)


def _run_libreoffice(
    provider: ConversionProvider,
    source: Path,
    output: Path,
    timeout_seconds: int,
) -> tuple[_NumberFormatRecord, ...]:
    if output.exists():
        raise _ProviderFailure(f"conversion output already exists: {output}")
    with tempfile.TemporaryDirectory(prefix=".workbooklens-libreoffice-", dir=output.parent) as raw:
        workspace = Path(raw)
        staged_input = workspace / "source.xls"
        converted = workspace / "source.xlsx"
        profile = workspace / "profile"
        output_dir = workspace / "output"
        profile.mkdir()
        output_dir.mkdir()
        shutil.copyfile(source, staged_input)
        command = [
            str(provider.runner),
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--norestore",
            "--convert-to",
            "xlsx:Calc MS Excel 2007 XML",
            "--outdir",
            str(output_dir),
            str(staged_input),
        ]
        popen_options: dict[str, Any] = {
            "stdout": subprocess.PIPE,
            "stderr": subprocess.PIPE,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
        }
        if sys.platform == "win32":
            popen_options["creationflags"] = _creation_flags() | int(
                getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            )
        else:
            popen_options["start_new_session"] = True
        try:
            process = subprocess.Popen(command, **popen_options)  # noqa: S603
        except OSError as exc:
            raise _ProviderFailure(f"could not start LibreOffice: {exc}") from exc
        try:
            stdout, stderr = process.communicate(timeout=timeout_seconds)
        except subprocess.TimeoutExpired as exc:
            _terminate_process_tree(process)
            raise _ProviderFailure(f"timed out after {timeout_seconds} seconds") from exc
        return_code = process.returncode if process.returncode is not None else -1
        completed = subprocess.CompletedProcess(command, return_code, stdout, stderr)
        generated = output_dir / converted.name
        if completed.returncode != 0:
            raise _ProviderFailure(_process_detail(completed))
        if not generated.is_file():
            raise _ProviderFailure(
                "LibreOffice reported success but did not create the expected .xlsx file"
            )
        if output.exists():
            raise _ProviderFailure(
                "conversion output appeared while LibreOffice was running; refusing to overwrite it"
            )
        generated.rename(output)
    return ()


def _run_provider(
    provider: ConversionProvider,
    source: Path,
    output: Path,
    timeout_seconds: int,
) -> tuple[_NumberFormatRecord, ...]:
    if provider.kind == "excel":
        return _run_excel(provider, source, output, timeout_seconds)
    elif provider.kind == "libreoffice":
        return _run_libreoffice(provider, source, output, timeout_seconds)
    else:
        assert_never(provider.kind)


def _validate_source(source: Path) -> Path:
    resolved = source.expanduser().resolve()
    if source.is_symlink() or not resolved.is_file():
        raise UsageError(f"Legacy workbook does not exist or is not a regular file: {source}")
    if resolved.suffix.lower() != ".xls":
        raise UsageError("Legacy conversion accepts only .xls input files")
    with resolved.open("rb") as handle:
        signature = handle.read(len(_OLE_COMPOUND_FILE_SIGNATURE))
    if signature != _OLE_COMPOUND_FILE_SIGNATURE:
        raise UsageError("Upload is not a recognized binary Excel .xls workbook")
    return resolved


def _validate_output(output: Path, maximum: int) -> None:
    inspection = inspect_package(output, PackageLimits(max_file_bytes=maximum))
    if not inspection.repairable:
        raise UsageError("The local converter did not produce a verified macro-free .xlsx workbook")


def convert_xls_to_xlsx(
    source: Path,
    output: Path,
    *,
    max_output_bytes: int = 100 * 1024 * 1024,
    timeout_seconds: int = _DEFAULT_TIMEOUT_SECONDS,
) -> ConversionResult:
    """Convert one legacy workbook locally and validate the resulting OOXML package.

    This function never attempts to parse BIFF itself. Microsoft Excel is preferred for fidelity;
    LibreOffice is an explicit local fallback. Both providers run without a shell.
    """

    if max_output_bytes <= 0:
        raise ValueError("max_output_bytes must be positive")
    if timeout_seconds <= 0:
        raise ValueError("timeout_seconds must be positive")
    resolved_source = _validate_source(source)
    resolved_output = output.expanduser().resolve()
    if resolved_output.suffix.lower() != ".xlsx":
        raise UsageError("Legacy conversion output must use the .xlsx extension")
    if resolved_output.exists():
        raise UsageError(f"Conversion output already exists: {output}")
    if not resolved_output.parent.is_dir():
        raise UsageError(f"Conversion output directory does not exist: {output.parent}")

    providers = available_conversion_providers()
    if not providers:
        raise UsageError(
            "No local .xls converter is available. Install Microsoft Excel or LibreOffice, "
            "then restart WorkbookLens. No workbook was uploaded to a cloud service."
        )

    failures: list[str] = []
    with tempfile.TemporaryDirectory(
        prefix=".workbooklens-convert-",
        dir=resolved_output.parent,
    ) as raw:
        staged_output = Path(raw) / "staged.xlsx"
        for provider in providers:
            try:
                format_records = _run_provider(
                    provider,
                    resolved_source,
                    staged_output,
                    timeout_seconds,
                )
                _validate_output(staged_output, max_output_bytes)
                _restore_number_formats(staged_output, format_records)
                _validate_output(staged_output, max_output_bytes)
            except (_ProviderFailure, WorkbookLensError, OSError) as exc:
                staged_output.unlink(missing_ok=True)
                failures.append(f"{provider.label}: {exc}")
                continue
            try:
                os.link(staged_output, resolved_output)
            except FileExistsError as exc:
                raise UsageError(
                    "Conversion output appeared while conversion was running; "
                    "WorkbookLens refused to overwrite or delete it"
                ) from exc
            except OSError as exc:
                raise UsageError(
                    f"Could not publish the verified conversion without overwriting "
                    f"another file: {exc}"
                ) from exc
            return ConversionResult(resolved_output, provider)

    detail = "; ".join(failures)
    raise UsageError(f"Every available local converter failed. {detail}")


__all__ = [
    "XLSX_MEDIA_TYPE",
    "ConversionProvider",
    "ConversionResult",
    "available_conversion_providers",
    "convert_xls_to_xlsx",
]
