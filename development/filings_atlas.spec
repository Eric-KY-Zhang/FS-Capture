# PyInstaller spec for Filings Atlas
#
# Build with:
#   pyinstaller --noconfirm filings_atlas.spec
#
# Result: dist/Filings Atlas/  (one-folder mode for faster cold start)

# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files

block_cipher = None

# Collect everything for libraries that ship runtime data files
ak_datas, ak_binaries, ak_hiddenimports = collect_all("akshare")
odr_datas, odr_binaries, odr_hiddenimports = collect_all("OpenDartReader")
pykrx_datas = collect_data_files("pykrx")
certifi_datas = collect_data_files("certifi")
playwright_datas, playwright_binaries, playwright_hiddenimports = collect_all("playwright")
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all("openpyxl")

extra_datas = [
    ("app/ui/styles/app.qss", "app/ui/styles"),
    ("app/assets/filings_atlas.ico", "app/assets"),
    ("app/assets/filings_atlas_logo.png", "app/assets"),
]

a = Analysis(
    ["app/main.py"],
    pathex=["."],
    binaries=ak_binaries + odr_binaries + playwright_binaries + openpyxl_binaries,
    datas=ak_datas + odr_datas + pykrx_datas + certifi_datas + playwright_datas + openpyxl_datas + extra_datas,
    hiddenimports=ak_hiddenimports + odr_hiddenimports + [
        "plugins.ashare", "plugins.hk", "plugins.us", "plugins.kr",
        "plugins.tw", "plugins.jp", "plugins.uk",
        "tomli_w", "certifi",
    ] + playwright_hiddenimports + openpyxl_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Heavy unused stacks; keep package small.
        "matplotlib", "scipy", "torch", "tensorflow",
        "PySide6.QtQuick", "PySide6.QtMultimedia", "PySide6.QtWebEngineCore",
        "PySide6.Qt3D", "PySide6.QtCharts",
    ],
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
    name="Filings Atlas",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="app/assets/filings_atlas.ico",
    version="version_info.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Filings Atlas",
)
