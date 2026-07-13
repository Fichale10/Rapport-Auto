@echo off
REM ============================================================
REM  Script de démarrage - ReportXCare (Production)
REM ============================================================

cd /d "%~dp0"

REM Variables d'environnement de production
SET DJANGO_DEBUG=False
SET DJANGO_ALLOWED_HOSTS=*

REM Secrets locaux (cle API IA, etc.) - fichier NON versionne dans .venv
IF EXIST ".venv\secrets.bat" (
    CALL ".venv\secrets.bat"
) ELSE (
    echo [WARN] .venv\secrets.bat introuvable : le chatbot ISOC_IA sera desactive.
)

REM Récupérer les dernières modifications
echo [INFO] Mise à jour depuis le dépôt Git...
git pull
if %ERRORLEVEL% NEQ 0 (
    echo [WARN] git pull a échoué, démarrage avec le code existant.
)

REM Appliquer les migrations si besoin
echo [INFO] Application des migrations...
python manage.py migrate --run-syncdb

REM Collecter les fichiers statiques si besoin
echo [INFO] Collecte des fichiers statiques...
python manage.py collectstatic --noinput

REM Démarrer le serveur Waitress
echo [INFO] Démarrage du serveur sur http://0.0.0.0:8000
echo [INFO] Accessible sur le réseau : http://10.234.24.30:8000
echo [INFO] Pour arrêter : Ctrl+C
echo.

python -m waitress --host=0.0.0.0 --port=8000 rapport_automatic.wsgi:application
