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
    file             = models.FileField(upload_to='uploads/')
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

    class Meta:
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"{self.original_filename} ({self.uploaded_at:%Y-%m-%d %H:%M})"


class Site(models.Model):
    site_name           = models.CharField(max_length=100, unique=True)
    date_mes            = models.DateField(null=True, blank=True)
    site_id             = models.CharField(max_length=20,  blank=True, default='')
    region              = models.CharField(max_length=50,  blank=True, default='')
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
    config_2g           = models.CharField(max_length=50,  blank=True, default='')
    config_3g           = models.CharField(max_length=50,  blank=True, default='')
    config_4g           = models.CharField(max_length=100, blank=True, default='')
    classif_tech        = models.CharField(max_length=30,  blank=True, default='')
    type_site           = models.CharField(max_length=30,  blank=True, default='')
    numero_agent        = models.CharField(max_length=30,  blank=True, default='')
    societe_gardiens    = models.CharField(max_length=50,  blank=True, default='')
    contacts_surveillants = models.CharField(max_length=100, blank=True, default='')

    class Meta:
        ordering = ['site_name']

    def __str__(self):
        return f"{self.site_name} ({self.site_id})"