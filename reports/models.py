import uuid
from django.db import models
from django.contrib.auth.models import User


class UploadedReport(models.Model):
    PERIOD_DAY = 'day'
    PERIOD_WEEK = 'week'
    PERIOD_MONTH = 'month'
    PERIOD_YEAR = 'year'
    PERIOD_CHOICES = [
        (PERIOD_DAY, 'Jour'),
        (PERIOD_WEEK, 'Semaine'),
        (PERIOD_MONTH, 'Mois'),
        (PERIOD_YEAR, 'Année'),
    ]

    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    user = models.ForeignKey(User, on_delete=models.CASCADE, null=True, blank=True)  # ← nouveau
    original_filename = models.CharField(max_length=255)
    file = models.FileField(upload_to='uploads/')
    date_rapport = models.DateField()
    date_fin = models.DateField(null=True, blank=True)
    period_type = models.CharField(max_length=10, choices=PERIOD_CHOICES, default=PERIOD_DAY)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processed = models.BooleanField(default=False)

    detailed_file = models.FileField(upload_to='results/', blank=True, null=True)
    synthesis_file = models.FileField(upload_to='results/', blank=True, null=True)

    total_rows = models.IntegerField(default=0)
    filtered_rows = models.IntegerField(default=0)
    total_incidents = models.IntegerField(default=0)
    unresolved_count = models.IntegerField(default=0)
    total_duration_sec = models.FloatField(default=0)
    processing_time_sec = models.FloatField(default=0)

    synthesis_json = models.JSONField(default=list, blank=True)
    top_sites_json = models.JSONField(default=list, blank=True)

    class Meta:
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"{self.original_filename} ({self.uploaded_at:%Y-%m-%d %H:%M})"