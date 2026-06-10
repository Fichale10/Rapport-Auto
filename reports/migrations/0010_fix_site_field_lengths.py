from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0009_site'),
    ]

    operations = [
        migrations.AlterField(
            model_name='site',
            name='ge_auto',
            field=models.CharField(blank=True, default='', max_length=20),
        ),
        migrations.AlterField(
            model_name='site',
            name='site_lithium',
            field=models.CharField(blank=True, default='', max_length=20),
        ),
        migrations.AlterField(
            model_name='site',
            name='site_esm',
            field=models.CharField(blank=True, default='', max_length=20),
        ),
    ]
