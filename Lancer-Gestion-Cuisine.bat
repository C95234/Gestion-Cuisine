@echo off
setlocal enabledelayedexpansion

REM === Gestion-Cuisine : lancement automatique (Windows) ===
REM Double-clique ce fichier. Il crée un environnement .venv, installe les dépendances, puis lance l'app.

cd /d "%~dp0"

REM 1) Trouver Python via le launcher "py" (recommandé sur Windows)
py --version >nul 2>&1
if %errorlevel%==0 (
    set PY=py
) else (
    REM Sinon, essayer "python"
    python --version >nul 2>&1
    if %errorlevel%==0 (
        set PY=python
    ) else (
        echo.
        echo ERREUR: Python n'est pas detecte.
        echo Installe Python depuis python.org (coche "Add Python to PATH"), puis relance.
        echo.
        pause
        exit /b 1
    )
)

REM 2) Creer .venv si absent
if not exist ".venv\Scripts\python.exe" (
    echo Creation de l'environnement virtuel...
    %PY% -m venv .venv
    if %errorlevel% neq 0 (
        echo.
        echo ERREUR: impossible de creer l'environnement virtuel.
        echo.
        pause
        exit /b 1
    )
)

REM 3) Activer venv
call ".venv\Scripts\activate.bat"

REM 4) Installer/mettre a jour dependances
echo Installation des dependances...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo ERREUR: l'installation des dependances a echoue.
    echo Verifie ta connexion Internet ou les droits d'installation.
    echo.
    pause
    exit /b 1
)

REM 5) Lancer Streamlit
echo.
echo Lancement de l'application...
python -m streamlit run app.py

endlocal
