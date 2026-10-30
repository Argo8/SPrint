# -*- mode: python ; coding: utf-8 -*-
# Builds SPrint as a folder, dist/SPrint/ (SPrint.exe + its libraries).
# installer/SPrint.iss turns that folder into SPrint-Setup-<version>.exe.
import re
import sys

from PyInstaller.utils.hooks import collect_all

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

datas = [('icon.ico', '.')]
binaries = []
hiddenimports = ['PIL._tkinter_finder']   # Pillow's Tk bridge, used by ttkbootstrap's images
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['PRINT.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
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
