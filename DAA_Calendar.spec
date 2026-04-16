# ============================================================
# File: DAA_Calendar.spec
# Place beside: DAA_Calendar.py
# ============================================================
# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# PyInstaller can execute specs without __file__. SPECPATH is reliable.
spec_dir = globals().get("SPECPATH", os.getcwd())
project_dir = os.path.abspath(spec_dir)

script_path = os.path.join(project_dir, "DAA_Calendar.py")

xlsx_path = os.path.join(project_dir, "DAA Claiming Spreadsheet.xlsx")
ico_path = os.path.join(project_dir, "DAACal.ico")
png_path = os.path.join(project_dir, "DAACal.png")
icons_dir = os.path.join(project_dir, "icons")

datas = []


def add_if_exists(src: str, dest: str) -> None:
    if os.path.exists(src):
        datas.append((src, dest))


add_if_exists(xlsx_path, ".")
add_if_exists(ico_path, ".")        # for runtime window icon usage
add_if_exists(png_path, ".")
add_if_exists(icons_dir, "icons")

hiddenimports = []
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("fitz")   # PyMuPDF
hiddenimports += collect_submodules("PyQt5")

a = Analysis(
    [script_path],
    pathex=[project_dir],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="DAA_Calendar",
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
    icon=ico_path if os.path.exists(ico_path) else None,  # FORCE exe icon
)