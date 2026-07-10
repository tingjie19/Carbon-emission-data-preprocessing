# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包設定

macOS：pyinstaller build.spec  → dist/碳健檢前處理工具.app
Windows：pyinstaller build.spec → dist/碳健檢前處理工具.exe
"""

import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# reportlab 子模組／字型資料檔很多，PyInstaller 的靜態掃描會漏收
# （實測會出現 ImportError: cannot import name 'canvas' from 'reportlab.pdfgen'，
# 導致蓋章功能在打包後的執行檔內被跳過）。用 collect_submodules/collect_data_files
# 直接掃整個已安裝套件目錄，取代原本只列 'reportlab.pdfgen' 這種漏網的寫法。
reportlab_hidden = collect_submodules('reportlab')
reportlab_datas  = collect_data_files('reportlab')

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates',  'templates'),   # HTML 前端
        ('範例',        '範例'),        # 範本 xlsx
        ('scripts',    'scripts'),     # Python 處理腳本
    ] + reportlab_datas,
    hiddenimports=[
        # Flask 相關
        'flask', 'werkzeug', 'jinja2', 'click', 'itsdangerous',
        # PDF 處理
        'pdfplumber', 'pdfminer', 'pdfminer.high_level', 'pdfminer.layout',
        'pypdf', 'pypdf.generic',
        # Excel 處理
        'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
        # 圖片 / 蓋章
        'PIL', 'PIL.Image',
        # 腳本模組
        'split_ledger', 'process_utilities', 'process_gasoline',
    ] + reportlab_hidden,
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

# onedir 模式：EXE 只放腳本本身，binaries/datas 由 COLLECT 另外收集成資料夾。
# onefile + macOS .app windowed bundle 這個組合會被 PyInstaller 標示為不建議，
# 官方建議一律改用 onedir。
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='碳健檢前處理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,          # 不顯示黑色終端視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # 可換成 .icns (macOS) 或 .ico (Windows)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='碳健檢前處理工具',
)

# macOS：額外產生 .app bundle
if sys.platform == 'darwin':
    app_bundle = BUNDLE(
        coll,
        name='碳健檢前處理工具.app',
        icon=None,
        bundle_identifier='com.carbon.audit.tool',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '1.0.0',
        },
    )
