from django.contrib import admin
from .models import UploadedReport, Site, Incident


@admin.register(UploadedReport)
class UploadedReportAdmin(admin.ModelAdmin):
    list_display = ('original_filename', 'date_rapport', 'date_fin', 'source', 'uploaded_at', 'processed')
    list_filter  = ('source', 'processed', 'period_type')
    search_fields = ('original_filename',)
    ordering = ('-uploaded_at',)


@admin.register(Site)
class SiteAdmin(admin.ModelAdmin):
    list_display  = ('site_name', 'site_id', 'region', 'zone', 'typ_energie', 'longitude', 'latitude')
    list_filter   = ('region', 'zone', 'typ_energie')
    search_fields = ('site_name', 'site_id', 'region', 'zone')
    ordering      = ('site_name',)


@admin.register(Incident)
class IncidentAdmin(admin.ModelAdmin):
    list_display  = ('domain', 'mois_rapport', 'numero_ticket', 'site_name', 'region', 'escalade', 'status', 'alarm_time')
    list_filter   = ('domain', 'mois_rapport', 'escalade', 'status', 'region')
    search_fields = ('numero_ticket', 'site_name', 'site_id', 'nature', 'cause')
    ordering      = ('-alarm_time',)
    date_hierarchy = 'alarm_time'
