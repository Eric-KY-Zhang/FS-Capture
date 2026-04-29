@echo off
REM Build the EXE distribution and refresh the root-level runnable EXE.
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\kaiyu\AppData\Local\Python\bin\python.exe" -m PyInstaller --noconfirm fs_capture.spec
if exist "dist\FS Capture\FS Capture.exe" (
    copy /Y "dist\FS Capture\FS Capture.exe" "..\FS Capture.exe" >nul
    if exist "..\_internal" rmdir /S /Q "..\_internal"
    xcopy /E /I /Y "dist\FS Capture\_internal" "..\_internal" >nul
)
echo.
echo Build complete. Run ..\FS Capture.exe
pause
