# -*- mode: python ; coding: utf-8 -*-
# DebtMail Pro - Single EXE build

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
    ],
    hiddenimports=[
        'win32com.client',
        'win32com.server',
        'pythoncom',
        'pywintypes',
        'email_template',
        'outlook_sender',
        'convert_to_template',
        'license_check',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'pandas', 'matplotlib', 'tkinter', 'PIL', 'scipy', 'pytest'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Remove unnecessary binaries to reduce size
a.binaries = [x for x in a.binaries if not x[0].startswith('numpy')]
a.binaries = [x for x in a.binaries if not x[0].startswith('_tk')]
a.binaries = [x for x in a.binaries if not x[0].startswith('tcl')]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DebtMailPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    icon=None,
)
