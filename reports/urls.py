from django.urls import path
from django.contrib.auth.decorators import login_required
from . import views

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
    path('sites-instables/',                      login_required(views.sites_instables), name='sites_instables'),
    path('reporting/',                            login_required(views.reporting),       name='reporting'),
    path('reporting/<slug:network>/',             login_required(views.reporting_network), name='reporting_network'),
    path('site-info/',                            login_required(views.site_info),       name='site_info'),
    path('site-info/search/',                     login_required(views.site_search_api), name='site_search_api'),
    path('statistiques/export/',                  login_required(views.export_statistiques), name='export_statistiques'),
    path('api-import/',                           login_required(views.api_import_view),     name='api_import'),
    path('audit/',                                login_required(views.audit_view),          name='audit'),
]