import uuid
from django.db import models
from django.contrib.auth.models import User


class UploadedReport(models.Model):
    PERIOD_DAY   = 'day'
    PERIOD_WEEK  = 'week'
    PERIOD_MONTH = 'month'
    PERIOD_YEAR  = 'year'
    PERIOD_CHOICES = [
        (PERIOD_DAY,   'Jour'),
        (PERIOD_WEEK,  'Semaine'),
        (PERIOD_MONTH, 'Mois'),
        (PERIOD_YEAR,  'Année'),
    ]

    id               = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    user             = models.ForeignKey(User, on_delete=models.CASCADE, null=True, blank=True)
    original_filename= models.CharField(max_length=255)
    file             = models.FileField(upload_to='uploads/', blank=True, null=True)
    date_rapport     = models.DateField()
    date_fin         = models.DateField(null=True, blank=True)
    period_type      = models.CharField(max_length=10, choices=PERIOD_CHOICES, default=PERIOD_DAY)
    uploaded_at      = models.DateTimeField(auto_now_add=True)
    processed        = models.BooleanField(default=False)

    detailed_file    = models.FileField(upload_to='results/', blank=True, null=True)
    synthesis_file   = models.FileField(upload_to='results/', blank=True, null=True)

    total_rows          = models.IntegerField(default=0)
    filtered_rows       = models.IntegerField(default=0)
    total_incidents     = models.IntegerField(default=0)
    unresolved_count    = models.IntegerField(default=0)
    total_duration_sec  = models.FloatField(default=0)
    processing_time_sec = models.FloatField(default=0)

    synthesis_json      = models.JSONField(default=list, blank=True)
    top_sites_json      = models.JSONField(default=list, blank=True)
    top_causes_json     = models.JSONField(default=list, blank=True)

    # ── Nouveau champ : outage journalier par escalade ──────────────────────
    # Structure :
    # {
    #   "ENERGIE":  {"2026-05-01": 24240, "2026-05-02": 6840, ...},
    #   "TRANS FH": {"2026-05-01": 1200, ...},
    #   ...
    # }
    # Valeurs en secondes d'outage cumulé par jour.
    outage_journalier_json = models.JSONField(default=dict, blank=True)
    region_sites_json      = models.JSONField(default=dict, blank=True)

    SOURCE_EXCEL = 'excel'
    SOURCE_API   = 'api'
    SOURCE_CHOICES = [(SOURCE_EXCEL, 'Excel'), (SOURCE_API, 'API')]
    source = models.CharField(max_length=10, choices=SOURCE_CHOICES, default=SOURCE_EXCEL)

    class Meta:
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"{self.original_filename} ({self.uploaded_at:%Y-%m-%d %H:%M})"


class Site(models.Model):
    site_name           = models.CharField(max_length=100, unique=True)
    date_mes            = models.DateField(null=True, blank=True)
    site_id             = models.CharField(max_length=20,  blank=True, default='')
    region              = models.CharField(max_length=50,  blank=True, default='')
    zone                = models.CharField(max_length=50,  blank=True, default='')
    base                = models.CharField(max_length=50,  blank=True, default='')
    olt                 = models.CharField(max_length=10,  blank=True, default='')
    longitude           = models.FloatField(null=True, blank=True)
    latitude            = models.FloatField(null=True, blank=True)
    config              = models.CharField(max_length=50,  blank=True, default='')
    techno              = models.CharField(max_length=30,  blank=True, default='')
    typ_trans           = models.CharField(max_length=30,  blank=True, default='')
    typ_energie         = models.CharField(max_length=50,  blank=True, default='')
    ge_auto             = models.CharField(max_length=20,  blank=True, default='')
    site_lithium        = models.CharField(max_length=20,  blank=True, default='')
    site_esm            = models.CharField(max_length=20,  blank=True, default='')
    site_solaire_neteco = models.CharField(max_length=20,  blank=True, default='')
    config_2g           = models.CharField(max_length=50,  blank=True, default='')
    config_3g           = models.CharField(max_length=50,  blank=True, default='')
    config_4g           = models.CharField(max_length=100, blank=True, default='')
    classif_tech        = models.CharField(max_length=30,  blank=True, default='')
    type_site           = models.CharField(max_length=30,  blank=True, default='')
    hauteur_pylone      = models.CharField(max_length=20,  blank=True, default='')
    typologie_pylone    = models.CharField(max_length=30,  blank=True, default='')
    numero_agent        = models.CharField(max_length=30,  blank=True, default='')
    societe_gardiens    = models.CharField(max_length=50,  blank=True, default='')
    contacts_surveillants = models.CharField(max_length=100, blank=True, default='')
    typologie_avant     = models.CharField(max_length=50,  blank=True, default='')
    typologie_apres     = models.CharField(max_length=50,  blank=True, default='')
    site_parent_1       = models.CharField(max_length=100, blank=True, default='')
    site_parent_2       = models.CharField(max_length=100, blank=True, default='')

    class Meta:
        ordering = ['site_name']

    def __str__(self):
        return f"{self.site_name} ({self.site_id})"


class Incident(models.Model):
    DOMAIN_MOBILE    = 'mobile'
    DOMAIN_FIXE      = 'fixe'
    DOMAIN_TRANSPORT = 'transport'
    DOMAIN_IGW       = 'igw'
    DOMAIN_CORE      = 'core'
    DOMAIN_CHOICES = [
        (DOMAIN_MOBILE,    'Réseau Mobile'),
        (DOMAIN_FIXE,      'Réseau Fixe'),
        (DOMAIN_TRANSPORT, 'Transport'),
        (DOMAIN_IGW,       'IGW'),
        (DOMAIN_CORE,      'Core'),
    ]

    domain              = models.CharField(max_length=20, choices=DOMAIN_CHOICES, db_index=True)
    mois_rapport        = models.DateField(null=True, blank=True, db_index=True)
    numero_ticket       = models.CharField(max_length=150, blank=True, default='', db_index=True)
    nature              = models.TextField(blank=True, default='')
    alarm_time          = models.DateTimeField(null=True, blank=True, db_index=True)
    cancel_time         = models.DateTimeField(null=True, blank=True)
    duration_sec        = models.FloatField(null=True, blank=True)
    site_parent         = models.CharField(max_length=150, blank=True, default='')
    site_name           = models.CharField(max_length=150, blank=True, default='', db_index=True)
    site_id             = models.CharField(max_length=100, blank=True, default='')
    region              = models.CharField(max_length=50,  blank=True, default='', db_index=True)
    base                = models.CharField(max_length=50,  blank=True, default='')
    plateforme          = models.CharField(max_length=150, blank=True, default='')
    technologies        = models.CharField(max_length=100, blank=True, default='')
    impact_equipement   = models.TextField(blank=True, default='')
    impact_service      = models.TextField(blank=True, default='')
    escalade            = models.CharField(max_length=80,  blank=True, default='', db_index=True)
    cause               = models.TextField(blank=True, default='')
    root_cause          = models.TextField(blank=True, default='')
    action              = models.TextField(blank=True, default='')
    technicien_informe  = models.TextField(blank=True, default='')
    technicien_maint    = models.TextField(blank=True, default='')
    point_bloquant      = models.TextField(blank=True, default='')
    observation         = models.TextField(blank=True, default='')
    status              = models.CharField(max_length=50,  blank=True, default='', db_index=True)
    nbre_clients        = models.CharField(max_length=50,  blank=True, default='')
    source_file         = models.CharField(max_length=255, blank=True, default='')
    imported_at         = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-alarm_time']
        indexes = [
            models.Index(fields=['domain', 'mois_rapport']),
            models.Index(fields=['domain', 'escalade']),
            models.Index(fields=['alarm_time', 'domain']),
        ]

    def __str__(self):
        return f"[{self.domain}] {self.numero_ticket or self.nature[:40]} ({self.alarm_time})"

    @property
    def is_dr2(self):
        """True si l'incident a duré ≥ 3h après la prochaine heure pleine suivant l'alarme."""
        if not self.alarm_time:
            return False
        from datetime import timedelta
        from django.utils import timezone as tz
        alarm = self.alarm_time
        if alarm.tzinfo is not None:
            alarm = alarm.replace(tzinfo=None)
        # Temps restant jusqu'à la prochaine heure pleine (toujours au moins 1h après)
        partial = alarm.minute * 60 + alarm.second
        secs_to_next_hour = (3600 - partial) if partial > 0 else 3600
        dr2_offset = secs_to_next_hour + 3 * 3600
        end = self.cancel_time if self.cancel_time else tz.now()
        if end.tzinfo is not None:
            end = end.replace(tzinfo=None)
        return (end - alarm).total_seconds() >= dr2_offset