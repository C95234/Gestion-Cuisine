@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ================================
echo   Gestion-Cuisine - Lancement
echo ================================

REM 1) Verifie Python via "py"
py -V >nul 2>&1
if errorlevel 1 (
  echo.
  echo ERREUR : Python n'est pas detecte via la commande "py".
  echo Installe Python depuis python.org (cocher "Add Python to PATH")
  echo puis relance ce fichier.
  echo.
  pause
  exit /b 1
)

REM 2) Cree / active un venv local
if not exist ".venv" (
  echo Creation de l'environnement .venv ...
  py -m venv .venv
  if errorlevel 1 (
    echo ERREUR : impossible de creer .venv
    pause
    exit /b 1
  )
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
  echo ERREUR : activation de .venv impossible
  pause
  exit /b 1
)

REM 3) Installe les dependances
echo Installation / mise a jour des dependances ...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo ERREUR : installation des dependances echouee.
  echo Verifie ta connexion internet / proxy.
  echo.
  pause
  exit /b 1
)

REM 4) Lance Streamlit via -m (pas besoin de streamlit dans le PATH)
echo.
echo Lancement de l'application...
python -m streamlit run app.py
pause
