# -*- mode: python ; coding: utf-8 -*-

# PyInstaller spec file for Report Writer
# Includes app.manifest for drag-and-drop fix

block_cipher = None

a = Analysis(
    ['ReportWriter.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('_assets/splash.png', '_assets'),
        ('_assets/icon.ico', '_assets'),
        ('_assets/brinelling.png', '_assets'),
        ('_assets/cage.png', '_assets'),
        ('_assets/corrosion.png', '_assets'),
        ('_assets/fluting.png', '_assets'),
        ('_assets/heat.png', '_assets'),
        ('_assets/lubrication.png', '_assets'),
        ('_assets/misalignment.png', '_assets'),
        ('_assets/spalling.png', '_assets'),
    ],
    hiddenimports=[],  # Removed docx2pdf - no longer needed
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Report Writer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Windowed mode (no console)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='_assets/icon.ico',
    version='_assets/version.txt',
    manifest='app.manifest',  # CRITICAL: Fixes drag-and-drop on all PCs
)
