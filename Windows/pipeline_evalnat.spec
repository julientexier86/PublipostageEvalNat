# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pipeline_evalnat.py'],
    pathex=[],
    binaries=[],
    datas=[('split_4C.py', '.'), ('merge_parents_4e.py', '.'), ('normalize.py', '.'), ('build_mailmerge_4e_from_merged_v5.py', '.'), ('tb_mailmerge_mac.py', '.'), ('ocr_helper.py', '.')],
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
    name='pipeline_evalnat',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
