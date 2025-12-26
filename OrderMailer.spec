# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('.env', '.'),
        ('email_accounts.json', '.'),
        ('app_icon.ico', '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # サイズ削減: 実行時に不要な開発用/テスト用依存を除外（機能・UIは変更しない）
    excludes=[
        # pytest stack (dev only)
        'pytest',
        '_pytest',
        'pluggy',
        'py',
        # occasionally pulled-in but unused here
        'numpy',
        'numpy.core',
        'numpy.testing',
        'pygments',
        'IPython',
        'jedi',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OrderMailer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # コンソールウィンドウを表示しない
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',  # アイコンファイルを指定
)

