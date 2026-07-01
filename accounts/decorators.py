from functools import wraps
from django.shortcuts import redirect
from django.contrib import messages


def gestionnaire_required(view_func):
    """Réserve la vue aux Gestionnaires et Administrateurs (pas les Lecteurs)."""
    @wraps(view_func)
    def wrapper(request, *args, **kwargs):
        try:
            if request.user.profile.can_write:
                return view_func(request, *args, **kwargs)
        except Exception:
            if request.user.is_superuser:
                return view_func(request, *args, **kwargs)
        messages.error(request, "Accès réservé aux gestionnaires et administrateurs.")
        return redirect('home')
    return wrapper


def admin_required(view_func):
    """Réserve la vue aux Administrateurs uniquement."""
    @wraps(view_func)
    def wrapper(request, *args, **kwargs):
        try:
            if request.user.profile.can_manage_users:
                return view_func(request, *args, **kwargs)
        except Exception:
            if request.user.is_superuser:
                return view_func(request, *args, **kwargs)
        messages.error(request, "Accès réservé aux administrateurs.")
        return redirect('home')
    return wrapper
