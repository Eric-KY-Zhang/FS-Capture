@echo off
REM Build the EXE distribution and refresh the root-level runnable EXE.
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\kaiyu\AppData\Local\Python\bin\python.exe" -m PyInstaller --noconfirm --clean filings_atlas.spec
if errorlevel 1 exit /b 1

REM Keep only Chinese / English Qt translations; the app ships its own UI strings.
if exist "dist\Filings Atlas\_internal\PySide6\translations" (
    for /r "dist\Filings Atlas\_internal\PySide6\translations" %%F in (*.qm) do (
        echo %%~nF | findstr /i "zh_CN en_US en" >nul || del /q "%%F"
    )
)

REM Remove Qt developer tools that are not used by the packaged desktop app.
if exist "dist\Filings Atlas\_internal\PySide6" (
    del /q "dist\Filings Atlas\_internal\PySide6\designer*.exe" 2>nul
    del /q "dist\Filings Atlas\_internal\PySide6\linguist*.exe" 2>nul
)

if exist "dist\Filings Atlas\Filings Atlas.exe" (
    copy /Y "dist\Filings Atlas\Filings Atlas.exe" "..\Filings Atlas.exe" >nul
    if exist "..\_internal" rmdir /S /Q "..\_internal"
    xcopy /E /I /Y "dist\Filings Atlas\_internal" "..\_internal" >nul
)
echo.
echo Build complete. Run ..\Filings Atlas.exe
pause
