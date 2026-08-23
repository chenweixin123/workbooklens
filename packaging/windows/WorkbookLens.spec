# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


SPEC_DIR = Path(SPEC).resolve().parent
GUI_VERSION_FILE = os.environ.get("WORKBOOKLENS_GUI_VERSION_FILE")
CLI_VERSION_FILE = os.environ.get("WORKBOOKLENS_CLI_VERSION_FILE")
ICON_FILE = SPEC_DIR / "assets" / "WorkbookLens.ico"
if not GUI_VERSION_FILE or not CLI_VERSION_FILE:
    raise RuntimeError(
        "WORKBOOKLENS_GUI_VERSION_FILE and WORKBOOKLENS_CLI_VERSION_FILE are required"
    )
if not ICON_FILE.is_file():
    raise RuntimeError(f"WorkbookLens icon is required: {ICON_FILE}")

datas = []
datas += collect_data_files(
    "workbooklens.diff",
    includes=["templates/diff.html.j2"],
)
datas += collect_data_files(
    "workbooklens.reports",
    includes=["templates/scan.html.j2"],
)

hiddenimports = sorted(
    {
        "python_multipart.multipart",
        "winreg",
        "workbooklens.rules.builtin",
        *collect_submodules("uvicorn"),
    }
)
desktop_hiddenimports = sorted({*hiddenimports, "webview"})

cli_analysis = Analysis(
    [str(SPEC_DIR / "entry.py")],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
desktop_analysis = Analysis(
    [str(SPEC_DIR / "desktop_entry.py")],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=desktop_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)


def keep_desktop_entry(entry):
    destination = entry[0].replace("\\", "/").casefold()
    return not (
        destination.endswith("/webbrowserinterop.x86.dll")
        or "/clr_loader/ffi/dlls/x86/" in f"/{destination}"
    )


desktop_analysis.binaries = [
    entry for entry in desktop_analysis.binaries if keep_desktop_entry(entry)
]
desktop_analysis.datas = [
    entry for entry in desktop_analysis.datas if keep_desktop_entry(entry)
]

cli_pyz = PYZ(cli_analysis.pure)
desktop_pyz = PYZ(desktop_analysis.pure)

cli_exe = EXE(
    cli_pyz,
    cli_analysis.scripts,
    [("X utf8", None, "OPTION")],
    exclude_binaries=True,
    name="WorkbookLensCLI",
    contents_directory="_internal",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=False,
    uac_uiaccess=False,
    version=CLI_VERSION_FILE,
)

desktop_exe = EXE(
    desktop_pyz,
    desktop_analysis.scripts,
    [("X utf8", None, "OPTION")],
    exclude_binaries=True,
    name="WorkbookLens",
    contents_directory="_internal",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=True,
    argv_emulation=False,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICON_FILE),
    uac_admin=False,
    uac_uiaccess=False,
    version=GUI_VERSION_FILE,
)

bundle = COLLECT(
    desktop_exe,
    cli_exe,
    desktop_analysis.binaries,
    desktop_analysis.datas,
    cli_analysis.binaries,
    cli_analysis.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="WorkbookLens",
)
