import os
from django.apps import AppConfig


class ReportsConfig(AppConfig):
    name = 'reports'

    def ready(self):
        from django.conf import settings
        # In dev (DEBUG=True), the autoreloader spawns a child process with RUN_MAIN=true.
        # Only start the scheduler in that child to avoid double registration.
        if settings.DEBUG and not os.environ.get('RUN_MAIN'):
            return
        self._start_scheduler()

    def _start_scheduler(self):
        try:
            from django.conf import settings
            from apscheduler.schedulers.background import BackgroundScheduler
            from apscheduler.triggers.interval import IntervalTrigger

            interval = getattr(settings, 'TICKETING_API_IMPORT_INTERVAL_HOURS', 3)

            scheduler = BackgroundScheduler()
            from .scheduler import auto_api_import, auto_site_down
            scheduler.add_job(
                auto_api_import,
                trigger=IntervalTrigger(hours=interval),
                id='auto_api_import',
                name='Import API automatique',
                replace_existing=True,
            )
            sd_interval = getattr(settings, 'SITE_DOWN_INTERVAL_HOURS', 0)
            if sd_interval and getattr(settings, 'SITE_DOWN_NETWORK_BASES', []):
                scheduler.add_job(
                    auto_site_down,
                    trigger=IntervalTrigger(hours=sd_interval),
                    id='auto_site_down',
                    name='Traitement SITE DOWN automatique',
                    replace_existing=True,
                )
            scheduler.start()
        except Exception:
            pass
