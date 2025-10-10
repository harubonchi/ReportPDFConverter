# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build specification for the ReportPDFConverter project."""

import os
import pathlib

from PyInstaller.utils.hooks import collect_submodules


PROJECT_ROOT = pathlib.Path(os.getcwd())
ENTRY_SCRIPT = PROJECT_ROOT / "tray_launcher.py"

DATAS = [
    (PROJECT_ROOT / "templates", "templates"),
    (PROJECT_ROOT / "static", "static"),
    (PROJECT_ROOT / "fonts", "fonts"),
    (PROJECT_ROOT / "order.json", "."),
]


def _normalize_datas(entries):
    normalized = []
    for source, target in entries:
        if not source.exists():
            continue
        normalized.append((str(source), str(target)))
    return normalized


a = Analysis(
    [str(ENTRY_SCRIPT)],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=_normalize_datas(DATAS),
    hiddenimports=collect_submodules("PyQt6"),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ReportPDFConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)