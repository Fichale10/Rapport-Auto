from django.urls import path
from django.contrib.auth.decorators import login_required
from . import views  # noqa

urlpatterns = [
    path('',                                      login_required(views.home),           name='home'),
    path('upload/',                               login_required(views.upload),          name='upload'),
    path('process/<uuid:pk>/',                    login_required(views.process_report),  name='process_report'),
    path('process/<uuid:pk>/status/',             login_required(views.process_status),  name='process_status'),
    path('results/<uuid:pk>/',                    login_required(views.results),         name='results'),
    path('results/<uuid:pk>/pdf/',                login_required(views.export_pdf),      name='export_pdf'),
    path('download/<uuid:pk>/<str:file_type>/',   login_required(views.download_file),   name='download_file'),
    path('delete/<uuid:pk>/',                     login_required(views.delete_report),   name='delete_report'),
    path('history/',                              login_required(views.history),         name='history'),
    path('statistiques/',                         login_required(views.statistiques),    name='statistiques'),
    path('notifications/',                        login_required(views.notifications),   name='notifications'),
    path('comparer/',                             login_required(views.comparer),        name='comparer'),
    path('incident-tracking/',                    login_required(views.incident_tracking), name='incident_tracking'),
    path('incident-tracking/process/',                      login_required(views.isocep_process),  name='isocep_process'),
    path('incident-tracking/process/download/<str:token>/', login_required(views.isocep_download), name='isocep_download'),

    # ── Reporting ──────────────────────────────────────────────────────────────
    path('reporting/',                            login_required(views.reporting),       name='reporting'),

    # Outils globaux (avant le slug générique)
    path('reporting/dr2-daily/',                  login_required(views.dr2_daily_report),  name='dr2_daily'),
    path('reporting/dr2-daily/export/',           login_required(views.dr2_daily_export),  name='dr2_daily_export'),
    path('reporting/generate-pptx/',              login_required(views.generate_pptx_report), name='generate_pptx'),
    path('reporting/bases-incidents/',            login_required(views.bases_incidents_view),   name='bases_incidents'),
    path('reporting/bases-incidents/export/',     login_required(views.bases_incidents_export), name='bases_incidents_export'),

    # Import global (rétro-compat)
    path('reporting/import/',                     login_required(views.reporting_import), name='reporting_import'),

    # ── Outils interactifs (avant le slug générique) ──────────────────────────
    path('reporting/igw/rapport-noc-core/',          login_required(views.igw_rapport_noc),          name='igw_rapport_noc'),
    path('reporting/igw/trafic-international/',      login_required(views.igw_trafic_international),  name='igw_trafic_international'),
    path('reporting/transmission/rapport-noc/',      login_required(views.transport_rapport_noc),     name='transport_rapport_noc'),
    path('reporting/transmission/rapport-dco-fo/',   login_required(views.transport_rapport_fo),      name='transport_rapport_fo'),
    path('reporting/fixe/rapport-ftth/',             login_required(views.fixe_rapport_ftth),         name='fixe_rapport_ftth'),

    # ── Par plateforme (slug doit venir après les chemins fixes) ───────────────
    path('reporting/<slug:platform>/import/',               login_required(views.reporting_platform_import),       name='reporting_platform_import'),
    path('reporting/<slug:platform>/pptx/',                 login_required(views.generate_pptx_platform),         name='generate_pptx_platform'),
    path('reporting/<slug:platform>/synthese/',             login_required(views.reporting_network),               name='reporting_network'),
    path('reporting/<slug:platform>/bases-incidents/',      login_required(views.platform_bases_incidents),        name='platform_bases_incidents'),
    path('reporting/<slug:platform>/bases-incidents/export/', login_required(views.platform_bases_incidents_export), name='platform_bases_incidents_export'),
    path('reporting/<slug:platform>/',                      login_required(views.reporting_platform),              name='reporting_platform'),

    # ── Autres ────────────────────────────────────────────────────────────────
    path('site-info/',                            login_required(views.site_info),          name='site_info'),
    path('site-info/search/',                     login_required(views.site_search_api),    name='site_search_api'),
    path('site-info/export/',                     login_required(views.sites_export_excel), name='sites_export_excel'),
    path('site-info/import/',                     login_required(views.sites_import_excel), name='sites_import_excel'),
    path('statistiques/export/',                  login_required(views.export_statistiques), name='export_statistiques'),
    path('api-import/',                           login_required(views.api_import_view),     name='api_import'),
    path('audit/',                                login_required(views.audit_view),          name='audit'),
]
