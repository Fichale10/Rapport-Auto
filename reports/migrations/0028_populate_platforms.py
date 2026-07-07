from django.db import migrations

PLATFORMS = [
    {'key': 'mobile',       'label': 'Mobile',       'prefix': 'API_MOBILE_',       'icon': '📡', 'color': '#003087'},
    {'key': 'fixe',         'label': 'Fixe',         'prefix': 'API_FIXE_',         'icon': '🔌', 'color': '#0050c8'},
    {'key': 'transmission', 'label': 'Transmission', 'prefix': 'API_TRANSMISSION_', 'icon': '📶', 'color': '#6b46c1'},
    {'key': 'core',         'label': 'Core & IGW',   'prefix': 'API_CORE_',         'icon': '⚙️', 'color': '#2d3748'},
]


def create_platforms(apps, schema_editor):
    Platform = apps.get_model('reports', 'Platform')
    for p in PLATFORMS:
        Platform.objects.get_or_create(key=p['key'], defaults=p)


def delete_platforms(apps, schema_editor):
    Platform = apps.get_model('reports', 'Platform')
    Platform.objects.filter(key__in=[p['key'] for p in PLATFORMS]).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0027_platform_importcoverage'),
    ]

    operations = [
        migrations.RunPython(create_platforms, delete_platforms),
    ]
