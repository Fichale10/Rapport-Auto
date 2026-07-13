"""Tâche planifiée : import API toutes les N heures."""
import logging
from datetime import timedelta

from django.conf import settings
from django.utils import timezone

logger = logging.getLogger(__name__)


def auto_api_import():
    """Import automatique de la dernière tranche de N heures."""
    from .api_import import run_import

    api_url  = getattr(settings, 'TICKETING_API_URL', '')
    api_user = getattr(settings, 'TICKETING_API_USERNAME', '')
    api_pass = getattr(settings, 'TICKETING_API_PASSWORD', '')

    if not (api_url and api_user and api_pass):
        logger.warning("auto_api_import : identifiants API manquants, import ignoré.")
        return

    interval = getattr(settings, 'TICKETING_API_IMPORT_INTERVAL_HOURS', 3)
    now   = timezone.now()
    d_fin = now.strftime('%Y-%m-%d')
    d_deb = (now - timedelta(hours=interval)).strftime('%Y-%m-%d')

    logger.info("auto_api_import : import %s → %s", d_deb, d_fin)
    try:
        result = run_import(d_deb, d_fin)
        logger.info("auto_api_import : %s créé(s), %s ignoré(s), %s erreur(s)",
                    result['created'], result['skipped'], len(result['errors']))
        if result['errors']:
            for err in result['errors']:
                logger.error("auto_api_import erreur : %s", err)
    except Exception:
        logger.exception("auto_api_import : exception non gérée")


def auto_site_down():
    """Collecte réseau + traitement automatique des alarmes SITE DOWN."""
    from .site_down import run_auto

    try:
        summary = run_auto()
        logger.info(
            "auto_site_down : %s collecté(s), %s traité(s), %s erreur(s), "
            "%s créée(s) / %s maj en base",
            summary.get('collected', 0), summary['processed'], summary['errors'],
            summary['created'], summary['updated'])
        for msg in summary['messages']:
            logger.info("auto_site_down : %s", msg)
    except Exception:
        logger.exception("auto_site_down : exception non gérée")
