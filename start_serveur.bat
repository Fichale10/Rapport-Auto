@echo off
REM ============================================================
REM  Script de démarrage - Rapport Automatique (Production)
REM ============================================================

cd /d "%~dp0"

REM Variables d'environnement de production
SET DJANGO_DEBUG=False
SET DJANGO_ALLOWED_HOSTS=*

REM Appliquer les migrations si besoin
echo [INFO] Application des migrations...
python manage.py migrate --run-syncdb

REM Collecter les fichiers statiques si besoin
echo [INFO] Collecte des fichiers statiques...
python manage.py collectstatic --noinput

REM Démarrer le serveur Waitress
echo [INFO] Démarrage du serveur sur http://0.0.0.0:8000
echo [INFO] Accessible sur le réseau : http://10.234.24.32:8000
echo [INFO] Pour arrêter : Ctrl+C
echo.

python -m waitress --host=0.0.0.0 --port=8000 rapport_automatic.wsgi:application
