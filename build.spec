# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置
将 GUI 应用和 OCR 模型打包为单个 exe
"""
import os
import sys
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# rapidocr 数据文件（模型 + 配置）
rapidocr_datas = collect_data_files('rapidocr_onnxruntime')

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=rapidocr_datas,
    hiddenimports=[
        'rapidocr_onnxruntime',
        'onnxruntime',
        'PIL',
        'pandas',
        'openpyxl',
        'numpy',
        'windnd',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'torch', 'tensorflow',
        'IPython', 'jupyter', 'notebook',
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
    name='股票图片OCR识别工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='股票图片OCR识别工具',
)
