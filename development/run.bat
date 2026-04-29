@echo off
REM Launch FS Capture from source (no EXE build required).
REM Use this for development. For end-user distribution build dist/FS Capture/ via:
REM   pyinstaller --noconfirm fs_capture.spec
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\kaiyu\AppData\Local\Python\bin\python.exe" -X utf8 -m app.main
