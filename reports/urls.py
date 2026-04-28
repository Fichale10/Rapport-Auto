from django.urls import path
from django.contrib.auth.decorators import login_required
from . import views

urlpatterns = [
    path('',                                    login_required(views.home),           name='home'),
    path('upload/',                             login_required(views.upload),          name='upload'),
    path('process/<uuid:pk>/',                  login_required(views.process_report),  name='process_report'),
    path('process/<uuid:pk>/status/',           login_required(views.process_status),  name='process_status'),  # ← nouveau
    path('results/<uuid:pk>/',                  login_required(views.results),         name='results'),
    path('download/<uuid:pk>/<str:file_type>/', login_required(views.download_file),   name='download_file'),
    path('history/',                            login_required(views.history),         name='history'),
    path('statistiques/',                       login_required(views.statistiques),    name='statistiques'),

path('process/<uuid:pk>/status/', login_required(views.process_status),  name='process_status'),
path('delete/<uuid:pk>/',         login_required(views.delete_report),    name='delete_report'),
path('results/<uuid:pk>/pdf/', login_required(views.export_pdf), name='export_pdf'),
path('notifications/', login_required(views.notifications), name='notifications'),
]
