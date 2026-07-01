from django.db import models
from django.contrib.auth.models import User
from django.db.models.signals import post_save
from django.dispatch import receiver


class UserProfile(models.Model):
    ROLE_ADMIN        = 'admin'
    ROLE_GESTIONNAIRE = 'gestionnaire'
    ROLE_LECTEUR      = 'lecteur'
    ROLE_CHOICES = [
        (ROLE_ADMIN,        'Administrateur'),
        (ROLE_GESTIONNAIRE, 'Gestionnaire'),
        (ROLE_LECTEUR,      'Lecteur'),
    ]

    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default=ROLE_LECTEUR)

    def __str__(self):
        return f"{self.user.username} ({self.get_role_display()})"

    @property
    def can_write(self):
        """Gestionnaire et Administrateur peuvent importer/exporter."""
        return self.role in (self.ROLE_ADMIN, self.ROLE_GESTIONNAIRE)

    @property
    def can_manage_users(self):
        return self.role == self.ROLE_ADMIN or self.user.is_superuser

    @property
    def role_badge_color(self):
        return {
            self.ROLE_ADMIN:        'admin',
            self.ROLE_GESTIONNAIRE: 'gestionnaire',
            self.ROLE_LECTEUR:      'lecteur',
        }.get(self.role, 'lecteur')


@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    if created:
        role = UserProfile.ROLE_ADMIN if instance.is_superuser else UserProfile.ROLE_LECTEUR
        UserProfile.objects.get_or_create(user=instance, defaults={'role': role})


@receiver(post_save, sender=User)
def save_user_profile(sender, instance, **kwargs):
    try:
        instance.profile.save()
    except UserProfile.DoesNotExist:
        UserProfile.objects.create(user=instance)
