from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0007_uploadedreport_top_causes_json'),
    ]

    operations = [
        migrations.AddField(
            model_name='uploadedreport',
            name='region_sites_json',
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
