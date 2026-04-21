# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['cable_design_system_v4.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('cable_config.json', '.'),
        ('cable_config_example.json', '.'),
        ('cable_products_v4.db', '.'),
        ('cable_codes.db', '.'),
        ('演示项目文件夹', '演示项目文件夹'),
        ('文档说明', '文档说明'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'sqlite3',
        'json',
        'hashlib',
        're',
        'datetime',
        'os',
        'pathlib'
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='电缆设计系统',
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
    icon=None,
)
