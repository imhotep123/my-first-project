# -*- mode: python ; coding: utf-8 -*-
"""
kiyaku_viewer.spec
PyInstaller ビルド設定ファイル

使い方:
    pyinstaller kiyaku_viewer.spec --clean --noconfirm

出力:
    dist/kiyaku_viewer/kiyaku_viewer.exe  (--onedir)
    dist/kiyaku_viewer/_internal/         (依存DLL・ライブラリ)
"""

import os
import sys
import glob

spec_dir = os.path.dirname(os.path.abspath(SPEC))

# ── Python DLL を明示的に同梱 ──────────────────────────────────────
# Python 3.12+ では python3XX.dll が自動収集されないことがある。
# 実行中の Python インタープリタのディレクトリから python3*.dll を探して同梱する。
_python_dir = os.path.dirname(sys.executable)
_py_dlls    = glob.glob(os.path.join(_python_dir, 'python3*.dll'))
_binaries   = [(dll, '.') for dll in _py_dlls]

block_cipher = None

a = Analysis(
    [os.path.join(spec_dir, 'kiyaku_viewer.py')],

    pathex=[spec_dir],

    binaries=_binaries,   # Python DLL を明示的に含める

    datas=[],

    # 自動検出できない隠れインポート
    hiddenimports=[
        # python-docx / lxml
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.oxml.parser',
        'docx.oxml.shared',
        'docx.opc',
        'docx.opc.constants',
        'lxml',
        'lxml.etree',
        'lxml._elementpath',
        # tkinter
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
    ],

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
    name='kiyaku_viewer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,       # GUI アプリのためコンソールウィンドウなし
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# --onedir: exe + 依存ファイルを dist/kiyaku_viewer/ フォルダに出力
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='kiyaku_viewer',
)
