# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None
# SPECPATH may not resolve correctly on CI; use current working directory
script_dir = os.getcwd()

a = Analysis(
    [os.path.join(script_dir, 'proposal_builder.py')],
    pathex=[script_dir],
    binaries=[],
    datas=[
        (os.path.join(script_dir, 'scope_database.db'), '.'),
        (os.path.join(script_dir, 'Estimate & Clarifications Template.docx'), '.'),
    ],
    hiddenimports=[
        'docx',
        'docx.opc',
        'docx.opc.constants',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.shared',
        'lxml',
        'lxml.etree',
        'lxml._elementpath',
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
    name='Proposal Builder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=None,
    target_arch=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Proposal Builder',
)
