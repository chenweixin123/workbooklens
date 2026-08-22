# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


SPEC_DIR = Path(SPEC).resolve().parent
VERSION_FILE = os.environ.get("WORKBOOKLENS_VERSION_FILE")
if not VERSION_FILE:
    raise RuntimeError("WORKBOOKLENS_VERSION_FILE is required")

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

analysis = Analysis(
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
pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    [("X utf8", None, "OPTION")],
    exclude_binaries=True,
    name="WorkbookLens",
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
    version=VERSION_FILE,
)

bundle = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="WorkbookLens",
)
