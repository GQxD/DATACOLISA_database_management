# -*- mode: python ; coding: utf-8 -*-

import os

SPEC_DIR = os.path.abspath(globals().get("SPECPATH", os.getcwd()))
PROJECT_ROOT = os.path.abspath(os.path.join(SPEC_DIR, ".."))
TEMPLATE_FILE = os.path.join(PROJECT_ROOT, "Mapping", "COLISA_template_interne.xlsx")
EXTERNAL_DIR = os.path.join(PROJECT_ROOT, "external")
APP_ENTRY = os.path.join(SPEC_DIR, "ui_pyside6_poc.py")

datas = [(TEMPLATE_FILE, ".")]
if os.path.isdir(EXTERNAL_DIR):
    # Bundle optional external resources.
    datas.append((EXTERNAL_DIR, "external"))

a = Analysis(
    [APP_ENTRY],
    pathex=[PROJECT_ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    exclude_binaries=False,
    name='DATACOLISA_PySide6',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/datacolisa_logo4.ico',
)
