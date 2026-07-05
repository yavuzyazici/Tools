# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller yapılandırması — Toplu Fatura Mail Gönderici

Kullanım:
    pip install pyinstaller
    pyinstaller build.spec

Çıktı: dist/TopluFaturaMailer.exe  (tek dosya, konsolsuz)

Not: Ayar dosyası (ayarlar.json) exe'nin yanında DEĞİL,
%APPDATA%\TopluFaturaMailer\ klasöründe tutulur; exe'nin içine de gömülmez.
Bu yüzden burada 'datas' ile paketlenmez.
"""

block_cipher = None


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # Outlook (pywin32) gönderimi için — dinamik import olduğundan elle belirtilir
        'win32com',
        'win32com.client',
        'win32timezone',
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
    name='TopluFaturaMailer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI uygulaması: konsol penceresi açılmasın
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app.ico',       # Simge eklemek isterseniz app.ico koyup bu satırı açın
)
