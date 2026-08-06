# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller yapılandırması — Toplu Fatura Mail Gönderici

Kullanım:
    pip install pyinstaller
    pyinstaller build.spec              -> dist/TopluFaturaMailer/  (KLASÖR, hızlı açılır)

Ortam değişkenleriyle:
    set TFM_TEKDOSYA=1  -> tek .exe (taşıması kolay, AÇILIŞI YAVAŞ)
    set TFM_KONSOL=1    -> arkasında konsol penceresi olan tanılama sürümü

NEDEN VARSAYILAN KLASÖR (onedir)?
    Tek dosya (onefile) sürümü her çalıştırmada ~28 MB'lık içeriği %TEMP% altına
    açar; Windows Defender de bu binlerce dosyayı her seferinde tarar. Bu sırada
    pencere görünür ama mesaj döngüsü henüz çalışmadığından Windows "Yanıt
    vermiyor" der. Klasör sürümünde açma/tarama adımı yoktur; uygulama saniyeler
    yerine anında açılır. Kod tamamen aynıdır.

Arayüz (web/ klasörü: HTML/CSS/JS + Inter fontu) çıktının içine gömülür.
Not: Ayar dosyası (ayarlar.json) ve tanılama günlüğü (calisma.log)
%APPDATA%/TopluFaturaMailer/ klasöründe tutulur.
"""

import os

from PyInstaller.utils.hooks import collect_all

TEKDOSYA = os.environ.get("TFM_TEKDOSYA") == "1"
KONSOL = os.environ.get("TFM_KONSOL") == "1"

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
    datas=[('web', 'web')] + _datas,   # web/ klasörünü çıktının içine göm
    hiddenimports=[
        # Outlook (pywin32) gönderimi için — dinamik import olduğundan elle belirtilir
        'win32com',
        'win32com.client',
        'win32timezone',
        'pythoncom',
        'pywintypes',
        # pywebview Windows arka ucu (Edge WebView2 / clr)
        'webview.platforms.winforms',
        'webview.platforms.edgechromium',
        'clr',
    ] + _hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Kullanılmayan ağır paketler çıktıya girmesin (açılışı ve boyutu şişirir)
        'tkinter', 'unittest', 'pydoc_data', 'test',
        'PyQt5', 'PySide2', 'PySide6', 'gi', 'cefpython3',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

_ad = 'TopluFaturaMailer' + ('-tanilama' if KONSOL else '')

if TEKDOSYA:
    # --- Tek dosya: taşıması kolay, açılışı yavaş ---
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name=_ad,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,              # UPX kapalı: başlatma hızı + antivirüs uyumu
        upx_exclude=[],
        runtime_tmpdir=None,
        console=KONSOL,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        # icon='app.ico',
    )
else:
    # --- Klasör (varsayılan): anında açılır ---
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=_ad,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=KONSOL,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        # icon='app.ico',
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name=_ad,
    )
