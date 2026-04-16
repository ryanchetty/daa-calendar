@echo off
setlocal
cd /d "%~dp0"

REM --- Clean old outputs ---
if exist "build" rmdir /s /q "build"
if exist "dist"  rmdir /s /q "dist"
if exist "output" rmdir /s /q "output"

REM --- Build using the SPEC (no --onedir/--onefile allowed with .spec) ---
pyinstaller --noconfirm "DAA_Calendar.spec"
if errorlevel 1 (
  echo PyInstaller build failed.
  pause
  exit /b 1
)

REM --- Compile installer (Inno Setup) ---
where ISCC >nul 2>nul
if errorlevel 1 (
  set "ISCC_EXE=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
) else (
  set "ISCC_EXE=ISCC"
)

"%ISCC_EXE%" "installer.iss"
if errorlevel 1 (
  echo Inno Setup compile failed.
  pause
  exit /b 1
)

echo.
echo Done. Installer is in: output\
pause