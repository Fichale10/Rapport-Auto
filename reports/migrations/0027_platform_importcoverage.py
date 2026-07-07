from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0026_site_top_cause_json'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name='Platform',
            fields=[
                ('id',     models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('key',    models.CharField(max_length=30, unique=True)),
                ('label',  models.CharField(max_length=100)),
                ('prefix', models.CharField(blank=True, max_length=50)),
                ('icon',   models.CharField(default='📡', max_length=10)),
                ('color',  models.CharField(default='#003087', max_length=20)),
            ],
            options={'ordering': ['key']},
        ),
        migrations.CreateModel(
            name='ImportCoverage',
            fields=[
                ('id',            models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('date_from',     models.DateField()),
                ('date_to',       models.DateField()),
                ('status',        models.CharField(
                    choices=[('success', 'Succès'), ('partial', 'Partiel'), ('error', 'Erreur')],
                    default='success', max_length=10,
                )),
                ('records_count', models.IntegerField(default=0)),
                ('fetched_at',    models.DateTimeField(auto_now_add=True)),
                ('source',        models.CharField(default='api', max_length=20)),
                ('notes',         models.TextField(blank=True, default='')),
                ('platform', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='coverages',
                    to='reports.platform',
                )),
                ('triggered_by', models.ForeignKey(
                    blank=True, null=True,
                    on_delete=django.db.models.deletion.SET_NULL,
                    to=settings.AUTH_USER_MODEL,
                )),
            ],
            options={'ordering': ['-fetched_at']},
        ),
        migrations.AddIndex(
            model_name='importcoverage',
            index=models.Index(fields=['platform', 'date_from', 'date_to'], name='reports_imp_platfor_idx'),
        ),
        migrations.AddIndex(
            model_name='importcoverage',
            index=models.Index(fields=['platform', 'status'], name='reports_imp_status_idx'),
        ),
    ]
