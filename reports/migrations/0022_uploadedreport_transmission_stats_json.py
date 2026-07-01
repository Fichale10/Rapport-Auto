from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0021_uploadedreport_fixe_stats_json'),
    ]

    operations = [
        migrations.AddField(
            model_name='uploadedreport',
            name='transmission_stats_json',
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
