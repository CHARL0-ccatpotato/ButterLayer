# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files

# Collect all sub-packages and data for google-genai
genai_datas, genai_binaries, genai_hiddenimports = collect_all("google.genai")
genai_datas += collect_all("google.ai")[0]
genai_datas += collect_data_files("google.auth")

# Collect fitz (PyMuPDF) native libraries
fitz_datas, fitz_binaries, fitz_hiddenimports = collect_all("fitz")

# Collect certifi CA bundle (needed for HTTPS in frozen exe)
certifi_datas = collect_data_files("certifi")

block_cipher = None

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=fitz_binaries + genai_binaries,
    datas=[
        ("templates", "templates"),
        ("cover", "cover"),
        *fitz_datas,
        *genai_datas,
        *certifi_datas,
    ],
    hiddenimports=[
        *fitz_hiddenimports,
        *genai_hiddenimports,
        "google.auth",
        "google.auth.transport",
        "google.auth.transport.requests",
        "google.auth.credentials",
        "docx",
        "docx.oxml",
        "docx.oxml.ns",
        "docx.shared",
        "PIL._tkinter_finder",
        "pytesseract",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "scipy", "pandas", "jupyter"],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ButterLayer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon="cover/raw toas..ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="ButterLayer",
)
