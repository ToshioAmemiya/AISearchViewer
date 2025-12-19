@echo off
chcp 65001 >nul
setlocal

REM 事前に: pip install -U pyinstaller

REM クリーン
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

REM EXE化（名前を AISearchViewer1.2.exe にする）
pyinstaller --onefile --noconsole ^
  --name AISearchViewer1.2 ^
  --icon AISearchViewer.ico ^
  --add-data "AISearchViewer.ico;." ^
  AISearchViewer.py

echo.
echo Done. dist\AISearchViewer1.2.exe を確認してください。
pause
