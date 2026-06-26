from django.db import migrations


def dr2_to_mobile(apps, schema_editor):
    """DR2 est un indicateur ARCEP de fiabilité réseau mobile.
    Les incidents importés avec domain='dr2' appartiennent au domaine Mobile."""
    Incident = apps.get_model('reports', 'Incident')
    Incident.objects.filter(domain='dr2').update(domain='mobile')


def mobile_to_dr2(apps, schema_editor):
    """Reverse : ne peut pas restaurer l'état exact, on laisse en mobile."""
    pass


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0017_add_site_solaire_neteco'),
    ]

    operations = [
        migrations.RunPython(dr2_to_mobile, mobile_to_dr2),
    ]
