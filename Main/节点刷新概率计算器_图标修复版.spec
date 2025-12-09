# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['excel_filter.py'],
    pathex=[],
    binaries=[('wood.ico', '.')],
    datas=[('target.xlsx', '.')],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='节点刷新概率计算器_图标修复版',
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
    icon=['wood.ico'],
    resources=['wood.ico,1,ICONGROUP', 'wood.ico,2,ICONGROUP'],
)
