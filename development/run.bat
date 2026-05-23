@echo off
REM Launch Filings Atlas from source (no EXE build required).
REM Use this for development. For end-user distribution build dist/Filings Atlas/ via:
REM   pyinstaller --noconfirm filings_atlas.spec
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\kaiyu\AppData\Local\Python\bin\python.exe" -X utf8 -m app.main
