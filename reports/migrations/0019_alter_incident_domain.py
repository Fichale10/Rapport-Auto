# Migration existante en prod — stub pour résolution de conflit
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0018_reassign_dr2_to_mobile'),
    ]

    operations = [
        migrations.AlterField(
            model_name='incident',
            name='domain',
            field=models.CharField(choices=[('mobile', 'Réseau Mobile'), ('fixe', 'Réseau Fixe'), ('transport', 'Transport'), ('igw', 'IGW'), ('core', 'Core')], db_index=True, max_length=20),
        ),
    ]
