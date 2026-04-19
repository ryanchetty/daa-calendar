@echo off
setlocal
cd /d "%~dp0"

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "output" rmdir /s /q "output"

pyinstaller --noconfirm "DAA_Calendar.spec"
if errorlevel 1 (
  echo PyInstaller build failed.
  pause
  exit /b 1
)

if not exist "dist\DAA_Calendar\DAA_Calendar.exe" (
  echo.
  echo ERROR: Expected ONEDIR build was not produced.
  echo Expected:
  echo   dist\DAA_Calendar\DAA_Calendar.exe
  echo.
  echo Your DAA_Calendar.spec is still building ONEFILE.
  pause
  exit /b 1
)

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