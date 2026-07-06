# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec ファイル — Windows と Mac の両方に対応

import sys

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['lxml.etree'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if sys.platform == 'win32':
    # ---- Windows: 単一 .exe ファイルを生成 ----
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        name='PPT変換ツール',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,      # コンソールウィンドウを非表示にする
        icon=None,          # アイコンを使う場合は 'icon.ico' に変更
    )

else:
    # ---- Mac: .app バンドルを生成 ----
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='PPT変換ツール',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,
        argv_emulation=False,
        icon=None,          # アイコンを使う場合は 'icon.icns' に変更
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='PPT変換ツール',
    )
    app = BUNDLE(
        coll,
        name='PPT変換ツール.app',
        icon=None,
        bundle_identifier='com.example.ppt-converter',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '1.0.0',
        },
    )
