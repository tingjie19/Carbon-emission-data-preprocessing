# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包設定

macOS：pyinstaller build.spec  → dist/碳健檢前處理工具.app
Windows：pyinstaller build.spec → dist/碳健檢前處理工具.exe
"""

import sys

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates',  'templates'),   # HTML 前端
        ('範例',        '範例'),        # 範本 xlsx
        ('scripts',    'scripts'),     # Python 處理腳本
    ],
    hiddenimports=[
        # Flask 相關
        'flask', 'werkzeug', 'jinja2', 'click', 'itsdangerous',
        # PDF 處理
        'pdfplumber', 'pdfminer', 'pdfminer.high_level', 'pdfminer.layout',
        'pypdf', 'pypdf.generic',
        # Excel 處理
        'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
        # 圖片 / 蓋章
        'PIL', 'PIL.Image', 'reportlab', 'reportlab.pdfgen',
        'reportlab.lib.units',
        # 腳本模組
        'split_ledger', 'process_utilities', 'process_gasoline',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'numpy', 'scipy'],
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
    name='碳健檢前處理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 不顯示黑色終端視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # 可換成 .icns (macOS) 或 .ico (Windows)
)

# macOS：額外產生 .app bundle
if sys.platform == 'darwin':
    app_bundle = BUNDLE(
        exe,
        name='碳健檢前處理工具.app',
        icon=None,
        bundle_identifier='com.carbon.audit.tool',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '1.0.0',
        },
    )
