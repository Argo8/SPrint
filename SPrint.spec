# -*- mode: python ; coding: utf-8 -*-
# Builds SPrint as a folder, dist/SPrint/ (SPrint.exe + its libraries).
# installer/SPrint.iss turns that folder into SPrint-Setup-<version>.exe.
import os
import re
import sys

with open('sprint/__init__.py', encoding='utf-8') as fh:
    version = re.search(r'__version__ = "([\d.]+)"', fh.read()).group(1)
version_tuple = (tuple(int(part) for part in version.split('.')) + (0, 0, 0))[:4]

version_info = None
if sys.platform == 'win32':   # version resource, shown in the file's Properties > Details
    from PyInstaller.utils.win32.versioninfo import (FixedFileInfo, StringFileInfo, StringStruct,
                                                     StringTable, VarFileInfo, VarStruct,
                                                     VSVersionInfo)
    version_info = VSVersionInfo(
        ffi=FixedFileInfo(filevers=version_tuple, prodvers=version_tuple),
        kids=[
            StringFileInfo([StringTable('040904B0', [
                StringStruct('CompanyName', 'Filip Kozina'),
                StringStruct('FileDescription', 'SPrint - PDF print router'),
                StringStruct('FileVersion', version),
                StringStruct('InternalName', 'SPrint'),
                StringStruct('LegalCopyright', 'Filip Kozina'),
                StringStruct('OriginalFilename', 'SPrint.exe'),
                StringStruct('ProductName', 'SPrint'),
                StringStruct('ProductVersion', version),
            ])]),
            VarFileInfo([VarStruct('Translation', [0x0409, 1200])]),
        ],
    )

a = Analysis(
    ['PRINT.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'PIL', 'numpy'],   # not used at run time (pypdfium2 imports PIL only on request)
    noarchive=False,
)
# The window draws with plain Qt widgets: no software OpenGL or shader compiler needed.
UNUSED = ('opengl32sw', 'd3dcompiler', 'qt6quick', 'qt6qml', 'qt6pdf', 'qt6virtualkeyboard')
a.binaries = [b for b in a.binaries if not os.path.basename(b[0]).lower().startswith(UNUSED)]
# Qt's own translations are not needed: every text, dialog buttons included, comes from sprint/i18n.py.
a.datas = [d for d in a.datas if 'translations' not in d[0].replace('\\', '/').split('/')]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SPrint',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX-packed executables are often flagged by antivirus
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
    version=version_info,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='SPrint',
)
