# -*- mode: python ; coding: utf-8 -*-

import sys
import os

# Read version from _assets/version.txt
APP_VERSION = "1.0.0"
version_file = os.path.join('_assets', 'version.txt')
if os.path.exists(version_file):
    with open(version_file, 'r', encoding='utf-8') as f:
        APP_VERSION = f.read().strip()
    print(f"✓ Read version from _assets/version.txt: {APP_VERSION}")
else:
    print(f"⚠ Warning: _assets/version.txt not found, using default: {APP_VERSION}")

# Find Python DLL explicitly
python_dll = None
if sys.platform == 'win32':
    python_version = f"{sys.version_info.major}{sys.version_info.minor}"
    dll_name = f"python{python_version}.dll"
    
    # Search multiple locations
    search_paths = [
        sys.base_prefix,
        os.path.dirname(sys.executable),
        os.path.join(sys.base_prefix, 'DLLs'),
        sys.prefix,
    ]
    
    for path in search_paths:
        dll_path = os.path.join(path, dll_name)
        if os.path.exists(dll_path):
            python_dll = (dll_path, '.')
            print(f"✓ Found Python DLL: {dll_path}")
            break
    
    if not python_dll:
        print(f"✗ WARNING: Could not find {dll_name} in any location!")
        print(f"Searched: {search_paths}")

# Collect assets (includes version.txt)
datas = []
assets_dir = '_assets'
if os.path.exists(assets_dir):
    for file in os.listdir(assets_dir):
        file_path = os.path.join(assets_dir, file)
        if os.path.isfile(file_path):
            datas.append((file_path, '_assets'))
    print(f"✓ Found {len(datas)} asset files")

# Version info for Windows exe
version_info = None
if sys.platform == 'win32':
    version_parts = APP_VERSION.split('.')
    while len(version_parts) < 4:
        version_parts.append('0')
    
    version_tuple = tuple(int(x) for x in version_parts[:4])
    
    version_info = f'''
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={version_tuple},
    prodvers={version_tuple},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'Daniel Fraser'),
        StringStruct(u'FileDescription', u'Report Writer - Samsung Quality Lab Reports'),
        StringStruct(u'FileVersion', u'{APP_VERSION}'),
        StringStruct(u'InternalName', u'ReportWriter'),
        StringStruct(u'LegalCopyright', u'Daniel Fraser'),
        StringStruct(u'OriginalFilename', u'Report Writer.exe'),
        StringStruct(u'ProductName', u'Report Writer'),
        StringStruct(u'ProductVersion', u'{APP_VERSION}')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
'''
    
    # Write version info to file
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_info)
    
    print(f"✓ Created version info for .exe: {APP_VERSION}")

a = Analysis(
    ['reportwriter.py'],
    pathex=[],
    binaries=[python_dll] if python_dll else [],
    datas=datas,
    hiddenimports=['subprocess', 'tempfile', 'urllib.request', 'urllib.error', 'json'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# --- Exclude bundled Microsoft VC++ runtime DLLs ---
a.binaries = [
    b for b in a.binaries
    if not any(
        dll in b[0].lower()
        for dll in (
            'msvcp140',
            'vcruntime140',
            'vcruntime140_1',
            'ucrtbase'
        )
    )
]

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

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
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt' if os.path.exists('version_info.txt') else None,
    icon='_assets/icon.ico' if os.path.exists('_assets/icon.ico') else None,
    manifest='app.manifest' if os.path.exists('app.manifest') else None,
)
