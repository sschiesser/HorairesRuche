@echo off
chcp 65001 >nul
echo ============================================================
echo   VoeuxConcerts — Construction de l'executable standalone
echo ============================================================
echo.

if not exist ".venv\Scripts\pip.exe" (
    echo Erreur : virtualenv introuvable ^(.venv\^).
    echo Lancez d'abord : python -m venv .venv ^&^& .venv\Scripts\pip install -r requirements.txt
    pause
    exit /b 1
)

echo Installation de PyInstaller dans le virtualenv...
.venv\Scripts\pip install --quiet --upgrade pyinstaller
if errorlevel 1 (
    echo Erreur : PyInstaller n'a pas pu etre installe.
    pause
    exit /b 1
)

echo Construction en cours...
.venv\Scripts\pyinstaller ^
  --onefile ^
  --console ^
  --name VoeuxConcerts ^
  --collect-all ortools ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  --hidden-import openpyxl.worksheet.page ^
  main_standalone.py

if errorlevel 1 (
    echo.
    echo Erreur lors de la construction. Voir les messages ci-dessus.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Executable genere : dist\VoeuxConcerts.exe
echo ============================================================
echo.
echo Pour distribuer l'application, copiez ces elements ensemble :
echo   - dist\VoeuxConcerts.exe
echo   - le dossier data\     (fichiers Excel editables)
echo   - le dossier output\   (cree automatiquement si absent)
echo.
pause
