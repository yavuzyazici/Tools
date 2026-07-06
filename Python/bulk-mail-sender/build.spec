# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller yapılandırması — Toplu Fatura Mail Gönderici

Kullanım:
    pip install pyinstaller
    pyinstaller build.spec

Çıktı: dist/TopluFaturaMailer.exe  (tek dosya, konsolsuz)

Arayüz (web/ klasörü: HTML/CSS/JS + Inter fontu) exe'nin içine gömülür.
Not: Ayar dosyası (ayarlar.json) exe'nin yanında DEĞİL,
%APPDATA%/TopluFaturaMailer/ klasöründe tutulur; exe'nin içine gömülmez.
"""

from PyInstaller.utils.hooks import collect_all

# pywebview + pythonnet(clr) için tüm alt modül/veri/binary'leri topla
_datas, _binaries, _hidden = collect_all('webview')
for _pkg in ('clr_loader', 'pythonnet'):
    try:
        _d, _b, _h = collect_all(_pkg)
        _datas += _d; _binaries += _b; _hidden += _h
    except Exception:
        pass

block_cipher = None


a = Analysis(
    ['ui.py'],
    pathex=[],
    binaries=_binaries,
    datas=[('web', 'web')] + _datas,   # web/ klasörünü exe içine göm
    hiddenimports=[
        # Outlook (pywin32) gönderimi için — dinamik import olduğundan elle belirtilir
        'win32com',
        'win32com.client',
        'win32timezone',
        # pywebview Windows arka ucu (Edge WebView2 / clr)
        'webview.platforms.winforms',
        'webview.platforms.edgechromium',
        'clr',
    ] + _hidden,
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
    upx=False,              # UPX kapalı: başlatma hızı + antivirüs uyumu (pywebview/clr için önerilir)
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
