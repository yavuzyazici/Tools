@echo off
REM ============================================================
REM  Toplu Fatura Mail Gonderici - .exe olusturma betigi
REM  Cift tiklayarak veya komut satirindan calistirin.
REM ============================================================
setlocal

cd /d "%~dp0"

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
echo  TAMAMLANDI!  Cikti:  dist\TopluFaturaMailer.exe
echo ============================================================
echo  Not: Ayarlar (ayarlar.json) exe'nin yaninda olusur.
echo  Exe'yi tasinabilir bir klasorde tutup oradan calistirin.
echo ============================================================
pause
endlocal
