@echo off
REM ============================================================
REM  Script de demarrage - ReportXCare (Production)
REM  Autonome : venv, dependances, migrations, statics, serveur
REM ============================================================

cd /d "%~dp0"

REM Variables d'environnement de production
SET DJANGO_DEBUG=False
SET DJANGO_ALLOWED_HOSTS=*

REM ── Python du venv (cree si absent) ─────────────────────────
SET VENV_PY=.venv\Scripts\python.exe

IF NOT EXIST "%VENV_PY%" (
    echo [INFO] Environnement virtuel absent : creation de .venv...
    python -m venv .venv
    IF NOT EXIST "%VENV_PY%" (
        echo [ERREUR] Impossible de creer le venv. Verifiez que Python est installe.
        pause
        exit /b 1
    )
)

REM ── Fichier .env obligatoire ────────────────────────────────
IF NOT EXIST ".env" (
    echo [ERREUR] Fichier .env introuvable a la racine du projet.
    echo          Creez-le avec DJANGO_SECRET_KEY, DB_*, TICKETING_API_*,
    echo          SITE_DOWN_NETWORK_BASES, etc. puis relancez.
    pause
    exit /b 1
)

REM Secrets locaux (cle API IA, etc.) - fichier NON versionne dans .venv
IF EXIST ".venv\secrets.bat" (
    CALL ".venv\secrets.bat"
) ELSE (
    echo [WARN] .venv\secrets.bat introuvable : le chatbot ISOC_IA sera desactive.
)

REM ── Recuperer les dernieres modifications ───────────────────
echo [INFO] Mise a jour depuis le depot Git...
git pull
if %ERRORLEVEL% NEQ 0 (
    echo [WARN] git pull a echoue, demarrage avec le code existant.
)

REM ── Installer/mettre a jour les dependances (dans le venv) ──
echo [INFO] Installation des dependances...
"%VENV_PY%" -m pip install --upgrade pip --quiet
"%VENV_PY%" -m pip install -r requirements.txt --quiet
if %ERRORLEVEL% NEQ 0 (
    echo [ERREUR] pip install a echoue. Voir les messages ci-dessus.
    pause
    exit /b 1
)

REM ── Verification de la configuration Django ─────────────────
echo [INFO] Verification de la configuration...
"%VENV_PY%" manage.py check
if %ERRORLEVEL% NEQ 0 (
    echo [ERREUR] manage.py check a echoue. Deploiement interrompu.
    pause
    exit /b 1
)

REM ── Appliquer les migrations ────────────────────────────────
echo [INFO] Application des migrations...
"%VENV_PY%" manage.py migrate
if %ERRORLEVEL% NEQ 0 (
    echo [ERREUR] Les migrations ont echoue. Deploiement interrompu.
    pause
    exit /b 1
)

REM ── Collecter les fichiers statiques ────────────────────────
echo [INFO] Collecte des fichiers statiques...
"%VENV_PY%" manage.py collectstatic --noinput

REM ── Demarrer le serveur Waitress ────────────────────────────
echo.
echo [INFO] Demarrage du serveur sur http://10.234.24.30:8000
echo [INFO] Pour arreter : Ctrl+C
echo.

"%VENV_PY%" -m waitress --host=0.0.0.0 --port=8000 rapport_automatic.wsgi:application
