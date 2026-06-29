from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0020_merge_20260629_1225'),
    ]

    operations = [
        migrations.AddField(
            model_name='uploadedreport',
            name='fixe_stats_json',
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
