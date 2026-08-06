@echo off
REM ============================================================
REM  Toplu Fatura Mail Gonderici - .exe olusturma betigi
REM
REM  Kullanim:
REM    build.bat              -> KLASOR surumu (varsayilan, ANINDA acilir)
REM    build.bat tanilama     -> konsollu tanilama surumu (hata mesajlari gorunur)
REM    build.bat tekdosya     -> tek .exe (tasimasi kolay, acilisi yavas)
REM ============================================================
setlocal

cd /d "%~dp0"

set "MOD=%~1"
set "TFM_TEKDOSYA="
set "TFM_KONSOL="

if /i "%MOD%"=="tekdosya" set "TFM_TEKDOSYA=1"
if /i "%MOD%"=="tanilama" set "TFM_KONSOL=1"

echo [1/3] Gerekli paketler kuruluyor...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo HATA: Paket kurulumu basarisiz. Python kurulu mu?
    pause
    exit /b 1
)

echo.
echo [2/3] Onceki build ciktilari temizleniyor...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [3/3] .exe olusturuluyor (PyInstaller)...
python -m PyInstaller build.spec --noconfirm
if errorlevel 1 (
    echo.
    echo HATA: Derleme basarisiz oldu.
    pause
    exit /b 1
)

echo.
echo ============================================================
if defined TFM_TEKDOSYA (
    echo  TAMAMLANDI!  Cikti:  dist\TopluFaturaMailer.exe
    echo  NOT: Tek dosya surumu her acilista kendini %%TEMP%% altina acar,
    echo       bu yuzden acilis 10-30 sn surebilir ve Windows bu sirada
    echo       "Yanit vermiyor" diyebilir. Gunluk kullanim icin klasor
    echo       surumunu tercih edin:  build.bat
) else if defined TFM_KONSOL (
    echo  TANILAMA SURUMU:  dist\TopluFaturaMailer-tanilama\TopluFaturaMailer-tanilama.exe
    echo  Arkasindaki siyah konsol penceresini KAPATMAYIN; hatalar orada gorunur.
    echo  Ayrica: %%APPDATA%%\TopluFaturaMailer\calisma.log
) else (
    echo  TAMAMLANDI!  Cikti:  dist\TopluFaturaMailer\TopluFaturaMailer.exe
    echo  Klasorun TAMAMINI kopyalayin (exe tek basina calismaz).
    echo  Kisayol olusturmak icin exe'ye sag tik ^> Kisayol olustur.
)
echo ============================================================
echo  Ayarlar ve gunluk:  %%APPDATA%%\TopluFaturaMailer\
echo ============================================================
pause
endlocal
