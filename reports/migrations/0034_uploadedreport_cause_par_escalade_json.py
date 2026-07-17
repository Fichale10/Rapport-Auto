from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0033_uploadedreport_top_causes_count_json'),
    ]

    operations = [
        migrations.AddField(
            model_name='uploadedreport',
            name='cause_par_escalade_json',
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
