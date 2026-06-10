from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0008_uploadedreport_region_sites_json'),
    ]

    operations = [
        migrations.CreateModel(
            name='Site',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('site_name',             models.CharField(max_length=100, unique=True)),
                ('date_mes',              models.DateField(blank=True, null=True)),
                ('site_id',               models.CharField(blank=True, default='', max_length=20)),
                ('region',                models.CharField(blank=True, default='', max_length=50)),
                ('base',                  models.CharField(blank=True, default='', max_length=50)),
                ('olt',                   models.CharField(blank=True, default='', max_length=10)),
                ('longitude',             models.FloatField(blank=True, null=True)),
                ('latitude',              models.FloatField(blank=True, null=True)),
                ('config',                models.CharField(blank=True, default='', max_length=50)),
                ('techno',                models.CharField(blank=True, default='', max_length=30)),
                ('typ_trans',             models.CharField(blank=True, default='', max_length=30)),
                ('typ_energie',           models.CharField(blank=True, default='', max_length=50)),
                ('ge_auto',               models.CharField(blank=True, default='', max_length=10)),
                ('site_lithium',          models.CharField(blank=True, default='', max_length=10)),
                ('site_esm',              models.CharField(blank=True, default='', max_length=10)),
                ('config_2g',             models.CharField(blank=True, default='', max_length=50)),
                ('config_3g',             models.CharField(blank=True, default='', max_length=50)),
                ('config_4g',             models.CharField(blank=True, default='', max_length=100)),
                ('classif_tech',          models.CharField(blank=True, default='', max_length=30)),
                ('type_site',             models.CharField(blank=True, default='', max_length=30)),
                ('numero_agent',          models.CharField(blank=True, default='', max_length=30)),
                ('societe_gardiens',      models.CharField(blank=True, default='', max_length=50)),
                ('contacts_surveillants', models.CharField(blank=True, default='', max_length=100)),
            ],
            options={'ordering': ['site_name']},
        ),
    ]
