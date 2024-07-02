# myapp.spec
# -*- mode: python ; coding: utf-8 -*-

import os

block_cipher = None

project_dir = 'D:\\Project\\PycharmProject\\论文阅读记录器'

a = Analysis(
    [os.path.join(project_dir, 'main.py')],
    pathex=[project_dir],
    binaries=[],
    datas=[(os.path.join(project_dir, 'api.txt'), 'api.txt')],
    hiddenimports=[],
    hookspath=[os.path.join(project_dir, 'hooks')],
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
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 设置为False，不显示控制台窗口
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
