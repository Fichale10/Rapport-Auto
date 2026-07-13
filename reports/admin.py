from django.contrib import admin
from .models import UploadedReport, Site, Incident, ChatInteraction


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


@admin.register(ChatInteraction)
class ChatInteractionAdmin(admin.ModelAdmin):
    list_display  = ('created_at', 'user', 'site_focus', 'status', 'is_refusal',
                     'feedback', 'latency_ms', 'total_tokens', 'eval_success')
    list_filter   = ('status', 'feedback', 'is_refusal', 'error_type',
                     'eval_success', 'eval_hallucination')
    search_fields = ('question', 'reply', 'site_focus', 'error_detail')
    ordering      = ('-created_at',)
    date_hierarchy = 'created_at'
    readonly_fields = ('created_at', 'feedback_at')
    fieldsets = (
        ('Requête', {
            'fields': ('created_at', 'user', 'session_key', 'site_focus',
                       'question', 'reply', 'model_name'),
        }),
        ('Télémétrie', {
            'fields': ('status', 'error_type', 'error_detail', 'is_refusal',
                       'latency_ms', 'prompt_tokens', 'completion_tokens',
                       'total_tokens', 'context_chars', 'history_len'),
        }),
        ('Satisfaction', {
            'fields': ('feedback', 'feedback_at'),
        }),
        ('Évaluation humaine', {
            'fields': ('eval_success', 'eval_faithful', 'eval_hallucination',
                       'eval_tool_ok', 'eval_needs_human', 'eval_rating', 'eval_notes'),
        }),
    )
