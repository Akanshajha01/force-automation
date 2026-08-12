# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for Force Automation / Calibration Report
# Automation System.
#
# Build with:
#   pyinstaller Force_Automation.spec
#
# Produces a "onedir" build in dist/Force_Automation/ containing
# Force_Automation.exe plus all required files. Onedir is used
# instead of onefile because it starts faster and makes resource
# resolution (templates/assets) fully predictable -- the installer
# below packages the whole folder for the end user, so they never
# see or need to know about this.

import sys
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Dash, dash-bootstrap-components and Flask each ship their own
# static assets (JS/CSS bundled inside the packages themselves).
# These are NOT the same as our project's ./assets folder and must
# be collected separately or the UI will fail to load in the built app.
datas = []
datas += collect_data_files('dash')
datas += collect_data_files('dash_bootstrap_components')
datas += collect_data_files('flask')

# Our own project resources: certificate templates and app assets
# (logo, style.css). These are read at runtime via resource_path().
datas += [
    ('templates', 'templates'),
    ('assets', 'assets'),
]

hiddenimports = [
    'dash',
    'dash.dash_table',
    'dash_bootstrap_components',
    'flask',
    'werkzeug',
    'jinja2',
    'dateutil.relativedelta',
    'docx',
    'openpyxl',
    'pandas',
    'numpy',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='Force_Automation_Debug',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,           # DEBUG BUILD: shows console for troubleshooting
    disable_windowed_traceback=False,
    icon='installer/app_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Force_Automation_Debug',
)
