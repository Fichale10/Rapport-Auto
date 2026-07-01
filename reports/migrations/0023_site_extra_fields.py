from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0022_uploadedreport_transmission_stats_json'),
    ]

    operations = [
        migrations.AddField(
            model_name='site',
            name='technologie_fo',
            field=models.CharField(blank=True, default='', max_length=50),
        ),
        migrations.AddField(
            model_name='site',
            name='chef_base_dfo',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
        migrations.AddField(
            model_name='site',
            name='srt',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
        migrations.AddField(
            model_name='site',
            name='tech_field_rx_mob',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
        migrations.AddField(
            model_name='site',
            name='tech_field_fttx',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
        migrations.AddField(
            model_name='site',
            name='resp_fo_backbone_ftth',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
        migrations.AddField(
            model_name='site',
            name='tech_fo_backbone_ftth',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
        migrations.AddField(
            model_name='site',
            name='chef_base_energie',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
        migrations.AddField(
            model_name='site',
            name='tech_energie',
            field=models.CharField(blank=True, default='', max_length=200),
        ),
    ]
