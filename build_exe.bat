@echo off
cd /d "%~dp0"

REM ===============================
REM venv (ne la recree PAS a chaque fois)
REM ===============================
if not exist ".venv\Scripts\python.exe" (
  echo Creation de l'environnement virtuel...
  python -m venv .venv
)

call .venv\Scripts\activate

echo === MAJ pip ===
python -m pip install --upgrade pip

echo === Installation dependances ===
python -m pip install -r requirements.txt
python -m pip install --upgrade streamlit pyinstaller

echo === Nettoyage build/dist ===
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul

REM ===============================
REM Build mode
REM ===============================
set "BUILD_CONSOLE=1"

if "%BUILD_CONSOLE%"=="1" (
  set "CONSOLE_FLAG=--console"
  echo [INFO] Build en mode console (diagnostic)
) else (
  set "CONSOLE_FLAG=--noconsole"
  echo [INFO] Build en mode noconsole (production)
)

echo === Build EXE (Streamlit stable) ===
python -m PyInstaller ^
  --noconfirm ^
  --onedir ^
  --clean ^
  --name Gestion_cuisine_centrale ^
  %CONSOLE_FLAG% ^
  --collect-all streamlit ^
  --add-data "app.py;." ^
  --add-data "launcher.py;." ^
  --add-data "src;src" ^
  --add-data "assets;assets" ^
  --add-data "templates;templates" ^
  --add-data ".streamlit;.streamlit" ^
  launcher.py

echo.
echo Build termine !
echo EXE ici : dist\Gestion_cuisine_centrale\Gestion_cuisine_centrale.exe
pause
