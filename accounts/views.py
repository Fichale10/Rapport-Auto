from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout, update_session_auth_hash
from django.contrib import messages
from django.contrib.auth.models import User
from django.contrib.auth.decorators import login_required
#from axes.helpers import is_already_locked


def login_view(request):
    if request.user.is_authenticated:
        return redirect('/')

    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')

        # Compte inactif
        try:
            user_obj = User.objects.get(username=username)
            if not user_obj.is_active:
                messages.error(request, 'Votre compte est en attente de validation par un administrateur.')
                return render(request, 'accounts/login.html')
        except User.DoesNotExist:
            pass

        # axes bloque automatiquement via authenticate() après 5 échecs
        user = authenticate(request, username=username, password=password)

        if user is not None:
            login(request, user)
            next_url = request.GET.get('next', '/')
            return redirect(next_url)
        else:
            from axes.models import AccessAttempt
            attempts = AccessAttempt.objects.filter(username=username)
            if attempts.exists() and attempts.first().failures_since_start >= 5:
                messages.error(request, 'Trop de tentatives échouées. Réessayez dans 1 heure.')
            else:
                messages.error(request, 'Identifiant ou mot de passe incorrect.')

    return render(request, 'accounts/login.html')


def logout_view(request):
    logout(request)
    return redirect('accounts:login')


def register_view(request):
    if request.user.is_authenticated:
        return redirect('/')

    if request.method == 'POST':
        username   = request.POST.get('username', '').strip()
        email      = request.POST.get('email', '').strip()
        first_name = request.POST.get('first_name', '').strip()
        last_name  = request.POST.get('last_name', '').strip()
        password1  = request.POST.get('password1', '')
        password2  = request.POST.get('password2', '')

        if not username:
            messages.error(request, "L'identifiant est obligatoire.")
        elif User.objects.filter(username__iexact=username).exists():
            messages.error(request, 'Cet identifiant est déjà utilisé.')
        elif password1 != password2:
            messages.error(request, 'Les mots de passe ne correspondent pas.')
        else:
            from django.contrib.auth.password_validation import validate_password
            from django.core.exceptions import ValidationError
            try:
                validate_password(password1)
            except ValidationError as e:
                for err in e.messages:
                    messages.error(request, err)
                return render(request, 'accounts/register.html')

            user = User.objects.create_user(
                username=username,
                email=email,
                password=password1,
                first_name=first_name,
                last_name=last_name,
            )
            user.is_active = False
            user.save()
            messages.success(request, 'Inscription envoyée ! Votre compte sera activé par un administrateur.')
            return redirect('accounts:login')

    return render(request, 'accounts/register.html')


@login_required
def profil_view(request):
    from reports.models import UploadedReport

    if request.method == 'POST':
        action = request.POST.get('action')
        if action == 'update_info':
            request.user.first_name = request.POST.get('first_name', '').strip()
            request.user.last_name  = request.POST.get('last_name', '').strip()
            request.user.email      = request.POST.get('email', '').strip()
            request.user.save()
            messages.success(request, 'Informations mises à jour avec succès.')
        elif action == 'change_password':
            old  = request.POST.get('old_password')
            new1 = request.POST.get('new_password1')
            new2 = request.POST.get('new_password2')
            if not request.user.check_password(old):
                messages.error(request, 'Mot de passe actuel incorrect.')
            elif new1 != new2:
                messages.error(request, 'Les nouveaux mots de passe ne correspondent pas.')
            elif len(new1) < 8:
                messages.error(request, 'Le mot de passe doit contenir au moins 8 caractères.')
            else:
                request.user.set_password(new1)
                request.user.save()
                update_session_auth_hash(request, request.user)
                messages.success(request, 'Mot de passe changé avec succès.')
        return redirect('accounts:profil')

    reports = UploadedReport.objects.filter(user=request.user, processed=True)
    return render(request, 'accounts/profil.html', {
        'total_reports':    reports.count(),
        'total_incidents':  sum(r.total_incidents for r in reports),
        'total_unresolved': sum(r.unresolved_count for r in reports if r.unresolved_count),
    })


@login_required
def gestion_users(request):
    from accounts.models import UserProfile
    try:
        if not request.user.profile.can_manage_users:
            return redirect('/')
    except Exception:
        if not request.user.is_superuser:
            return redirect('/')

    if request.method == 'POST':
        import secrets, string
        user_id = request.POST.get('user_id')
        action  = request.POST.get('action')
        role    = request.POST.get('role', UserProfile.ROLE_LECTEUR)
        try:
            u = User.objects.get(pk=user_id)
            if action == 'approuver':
                u.is_active = True
                u.save()
                profile, _ = UserProfile.objects.get_or_create(user=u)
                if role in dict(UserProfile.ROLE_CHOICES):
                    profile.role = role
                    profile.save()
                messages.success(request, f'Compte de {u.username} approuvé avec le rôle {profile.get_role_display()}.')
            elif action == 'changer_role':
                if role in dict(UserProfile.ROLE_CHOICES):
                    profile, _ = UserProfile.objects.get_or_create(user=u)
                    profile.role = role
                    profile.save()
                    messages.success(request, f'Rôle de {u.username} changé en {profile.get_role_display()}.')
            elif action == 'reset_password':
                alphabet = string.ascii_letters + string.digits + '!@#&'
                new_pwd  = ''.join(secrets.choice(alphabet) for _ in range(12))
                u.set_password(new_pwd)
                u.save()
                messages.success(request, f'Mot de passe de {u.username} réinitialisé → {new_pwd}')
            elif action == 'rejeter':
                if u.pk != request.user.pk:
                    u.delete()
                    messages.success(request, 'Compte supprimé.')
                else:
                    messages.error(request, 'Impossible de supprimer votre propre compte.')
        except User.DoesNotExist:
            pass
        return redirect('accounts:gestion_users')

    en_attente   = User.objects.filter(is_active=False).order_by('-date_joined')
    actifs       = User.objects.filter(is_active=True).order_by('is_superuser', 'username')
    role_choices = UserProfile.ROLE_CHOICES

    # S'assure que chaque utilisateur a un profil
    for u in actifs:
        UserProfile.objects.get_or_create(
            user=u,
            defaults={'role': UserProfile.ROLE_ADMIN if u.is_superuser else UserProfile.ROLE_LECTEUR}
        )

    return render(request, 'accounts/gestion_users.html', {
        'en_attente':    en_attente,
        'actifs':        actifs,
        'role_choices':  role_choices,
    })