# Merge two 0023 migrations generated on the same base (0022).

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0023_site_extra_fields'),
        ('reports', '0023_alter_site_site_parent_1_alter_site_site_parent_2'),
    ]

    operations = [
    ]
