# -*- mode: python ; coding: utf-8 -*-
"""
AST (AJS Support Tool) - PyInstaller Build Spec
--onefile mode
"""

block_cipher = None

a = Analysis(
    ['ajs_main.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'ajs_constants',
        'ajs_print_logic',
        'ajs_define_logic',
        'ajs_inout_logic',
        'ajs_rel_logic',
        'ajs_depend_logic',
        'ajs_dep_tracer',
        'ajs_dep_report',
        'ajs_shell_parser',
        'ajs_inout_report',
        'ajs_exception_editor',
        'ajs_bridge_editor',
        'ajs_utils',
        'paramiko',
        'networkx',
        'openpyxl',
        'openpyxl.styles',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'scipy', 'pandas',
        'PIL', 'PyQt5', 'wx', 'unittest',
        'xmlrpc', 'pydoc', 'doctest',
    ],
    noarchive=False,
    optimize=0,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='AST',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=None,
)
