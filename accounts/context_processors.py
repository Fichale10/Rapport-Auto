def user_role(request):
    """Expose le profil et le rôle de l'utilisateur dans tous les templates."""
    if request.user.is_authenticated:
        try:
            profile = request.user.profile
        except Exception:
            from accounts.models import UserProfile
            profile, _ = UserProfile.objects.get_or_create(
                user=request.user,
                defaults={'role': UserProfile.ROLE_ADMIN if request.user.is_superuser else UserProfile.ROLE_LECTEUR}
            )
        return {
            'user_profile':   profile,
            'can_write':      profile.can_write,
            'can_manage':     profile.can_manage_users,
            'user_role_name': profile.get_role_display(),
        }
    return {
        'user_profile':   None,
        'can_write':      False,
        'can_manage':     False,
        'user_role_name': '',
    }
