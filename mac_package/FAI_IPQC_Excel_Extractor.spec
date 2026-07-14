# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all, collect_submodules

openpyxl_datas, openpyxl_binaries, openpyxl_hidden = collect_all('openpyxl')

datas = [
    ('../fai_excel_extractor.py', '.'),
]
datas += openpyxl_datas
binaries = openpyxl_binaries

hiddenimports = []
hiddenimports += openpyxl_hidden
hiddenimports += collect_submodules('openpyxl')
hiddenimports += ['tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox']

block_cipher = None


a = Analysis(
    ['../fai_excel_extractor_desktop.py'],
    pathex=[],
    binaries=binaries,
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
    [],
    exclude_binaries=True,
    name='FAI IPQC Excel Extractor',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='FAI IPQC Excel Extractor',
)
app = BUNDLE(
    coll,
    name='FAI IPQC Excel Extractor.app',
    icon=None,
    bundle_identifier='com.nordbo.fai-ipqc-excel-extractor',
    info_plist={
        'CFBundleDisplayName': 'FAI IPQC Excel Extractor',
        'CFBundleName': 'FAI IPQC Excel Extractor',
        'NSHighResolutionCapable': 'True',
    },
)
