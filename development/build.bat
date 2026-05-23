@echo off
REM Build the EXE distribution and refresh the root-level runnable EXE.
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\kaiyu\AppData\Local\Python\bin\python.exe" -m PyInstaller --noconfirm filings_atlas.spec
if exist "dist\Filings Atlas\Filings Atlas.exe" (
    copy /Y "dist\Filings Atlas\Filings Atlas.exe" "..\Filings Atlas.exe" >nul
    if exist "..\_internal" rmdir /S /Q "..\_internal"
    xcopy /E /I /Y "dist\Filings Atlas\_internal" "..\_internal" >nul
)
echo.
echo Build complete. Run ..\Filings Atlas.exe
pause
