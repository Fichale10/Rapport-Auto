import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = os.environ.get('DJANGO_SECRET_KEY', 'fallback-dev-only')
DEBUG = os.environ.get('DJANGO_DEBUG', 'False') == 'True'
ALLOWED_HOSTS = os.environ.get('DJANGO_ALLOWED_HOSTS', 'localhost').split(',')

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'reports',
    'accounts',
    'axes',
    'django_apscheduler',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'axes.middleware.AxesMiddleware',
]

ROOT_URLCONF = 'rapport_automatic.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
                'accounts.context_processors.user_role',
            ],
        },
    },
]

WSGI_APPLICATION = 'rapport_automatic.wsgi.application'

DATABASES = {
    'default': {
        'ENGINE':   'django.db.backends.postgresql',
        'NAME':     os.environ.get('DB_NAME', 'rapport_db'),
        'USER':     os.environ.get('DB_USER', 'postgres'),
        'PASSWORD': os.environ.get('DB_PASSWORD', ''),
        'HOST':     os.environ.get('DB_HOST', 'localhost'),
        'PORT':     os.environ.get('DB_PORT', '5432'),
        'OPTIONS':  {'client_encoding': 'UTF8'},
    }
}

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]

LANGUAGE_CODE = 'fr-fr'
TIME_ZONE = 'Africa/Lome'
USE_I18N = True
USE_TZ = True

STATIC_URL = 'static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_STORAGE = 'whitenoise.storage.CompressedManifestStaticFilesStorage'

MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# ── API Ticketing ──────────────────────────────────────────────────────────────
TICKETING_API_URL      = os.environ.get('TICKETING_API_URL', '').strip().rstrip('/')
TICKETING_API_USERNAME = os.environ.get('TICKETING_API_USERNAME', '').strip()
TICKETING_API_PASSWORD = os.environ.get('TICKETING_API_PASSWORD', '').strip()
# Fenêtre (en jours) de récupération des tickets API par le chatbot ISOC_IA
try:
    ISOC_API_WINDOW_DAYS = int(os.environ.get('ISOC_API_WINDOW_DAYS', '180'))
except ValueError:
    ISOC_API_WINDOW_DAYS = 180

# ── Chatbot ISOC_IA (Rodium AI / OpenAI-compatible) ─────────────────────────────
RODIUM_API_KEY  = os.environ.get('RODIUM_API_KEY', '')
RODIUM_API_BASE = os.environ.get('RODIUM_API_BASE', 'https://api.rodiumai.io/v1').rstrip('/')
RODIUM_MODEL    = os.environ.get('RODIUM_MODEL', 'anthropic/claude-fable-5')

# ── Auth ───────────────────────────────────────────────────────────────────────
LOGIN_URL           = '/accounts/login/'
LOGIN_REDIRECT_URL  = '/'
LOGOUT_REDIRECT_URL = '/accounts/login/'

# ── Sessions ───────────────────────────────────────────────────────────────────
SESSION_COOKIE_AGE              = 8 * 3600
SESSION_EXPIRE_AT_BROWSER_CLOSE = True
SESSION_COOKIE_HTTPONLY         = True
SESSION_COOKIE_SAMESITE         = 'Lax'

# ── Rate limiting (django-axes) ────────────────────────────────────────────────
AXES_FAILURE_LIMIT      = 5
AXES_COOLOFF_TIME       = 1
AXES_LOCKOUT_PARAMETERS = ['username', 'ip_address']
AXES_RESET_ON_SUCCESS   = True
AUTHENTICATION_BACKENDS = [
    'axes.backends.AxesStandaloneBackend',
    'django.contrib.auth.backends.ModelBackend',
]

# ── Logging ────────────────────────────────────────────────────────────────────
LOGGING = {
    'version': 1,
    'disable_existing_loggers': False,
    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'formatter': 'simple',
        },
    },
    'formatters': {
        'simple': {
            'format': '[%(levelname)s] %(name)s: %(message)s',
        },
    },
    'loggers': {
        'treatement': {
            'handlers': ['console'],
            'level': 'DEBUG',
            'propagate': False,
        },
        'reports': {
            'handlers': ['console'],
            'level': 'DEBUG',
            'propagate': False,
        },
    },
}