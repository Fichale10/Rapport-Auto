from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0038_dr2processeddate_dr2violationrecord'),
    ]

    operations = [
        migrations.AddField(
            model_name='dr2violationrecord',
            name='root_cause',
            field=models.TextField(blank=True, default=''),
        ),
    ]