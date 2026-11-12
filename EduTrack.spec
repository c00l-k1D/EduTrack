# -*- mode: python -*-
from pathlib import Path

block_cipher = None
base_dir = Path('.').resolve()

a = Analysis(
    ['main.py'],
    pathex=[str(base_dir)],
    binaries=[],
    datas=[
        (str(base_dir / 'image.png'), '.'),
        (str(base_dir / 'icon.ico'), '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EduTrack',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EduTrack',
)
