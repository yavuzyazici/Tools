@echo off
setlocal

echo === EXE olusturma basladi ===
python -m pip install --upgrade pip
python -m pip install pyinstaller dnspython openpyxl

pyinstaller ^
  --onefile ^
  --noconsole ^
  --name TopluMXKontrol ^
  mx_checker_gui.py

echo.
echo EXE hazir:
echo dist\TopluMXKontrol.exe
pause
