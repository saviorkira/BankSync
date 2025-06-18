# -*- mode: python ; coding: utf-8 -*-
import os
import pathlib

project_root = str(pathlib.Path().resolve())

def recursive_collect_files(src_folder, dest_folder):
    datas = []
    for root, dirs, files in os.walk(src_folder):
        for f in files:
            abs_path = os.path.join(root, f)
            rel_path = os.path.relpath(abs_path, src_folder)
            dest_path = os.path.join(dest_folder, os.path.dirname(rel_path))
            datas.append((abs_path, dest_path))
    return datas

datas = [
    (os.path.join(project_root, "config.txt"), "."),
]

datas += recursive_collect_files(
    os.path.join(project_root, "playwright-browsers"),
    "playwright-browsers"
)

block_cipher = None

a = Analysis(
    ['ui_launcher.py'],
    pathex=[project_root],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
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
    name='BankDownload',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BankDownload',
)
