@echo off
setlocal
cd /d "%~dp0"
set ICONO=assets\icon.ico

set PY=python
if exist "venv\Scripts\python.exe" set PY=venv\Scripts\python.exe
for /f "delims=" %%v in ('%PY% -c "from version import __version__; print(__version__)"') do set VERSION=%%v
echo Compilando version %VERSION%...

%PY% -m nuitka main.py ^
  --standalone ^
  --msvc=latest ^
  --assume-yes-for-downloads ^
  --enable-plugin=tk-inter ^
  --include-package=ui.pages ^
  --include-package-data=reportlab ^
  --nofollow-import-to=PyInstaller,wx,pytest,pymupdf,fitz ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=%ICONO% ^
  --include-data-dir=assets=assets ^
  --output-filename=Auditor.exe ^
  --output-dir=dist ^
  --remove-output ^
  --company-name="Asesores Contables y Tributarios" ^
  --product-name="Auditor" ^
  --file-description="Conciliador DIAN vs Contabilidad" ^
  --product-version=%VERSION% ^
  --file-version=%VERSION%

if errorlevel 1 (
  echo.
  echo *** La compilacion fallo. Revisa los mensajes anteriores. ***
  pause
  exit /b 1
)

set ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" set ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe
if not exist "%ISCC%" set ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" (
  echo.
  echo Programa compilado en dist\main.dist. No se encontro Inno Setup: compila instalador.iss manualmente.
  pause
  exit /b 0
)
"%ISCC%" /DVersion=%VERSION% instalador.iss
if errorlevel 1 (
  echo *** Fallo la creacion del instalador. ***
) else (
  echo.
  echo Instalador listo: instaladores\Auditor_Setup_%VERSION%.exe
)
pause