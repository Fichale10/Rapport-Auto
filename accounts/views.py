from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.contrib.auth.models import User


def login_view(request):
    # Déjà connecté → redirection
    if request.user.is_authenticated:
        return redirect('/')

    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')

        user = authenticate(request, username=username, password=password)

        if user is not None:
            login(request, user)
            # Respect du paramètre ?next=
            next_url = request.GET.get('next', '/')
            return redirect(next_url)
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
        username = request.POST.get('username', '').strip()
        email = request.POST.get('email', '').strip()
        first_name = request.POST.get('first_name', '').strip()
        last_name = request.POST.get('last_name', '').strip()
        password1 = request.POST.get('password1', '')
        password2 = request.POST.get('password2', '')

        if not username:
            messages.error(request, "L'identifiant est obligatoire.")
        elif User.objects.filter(username__iexact=username).exists():
            messages.error(request, 'Cet identifiant est deja utilise.')
        elif password1 != password2:
            messages.error(request, 'Les mots de passe ne correspondent pas.')
        elif len(password1) < 8:
            messages.error(request, 'Le mot de passe doit contenir au moins 8 caracteres.')
        else:
            user = User.objects.create_user(
                username=username,
                email=email,
                password=password1,
                first_name=first_name,
                last_name=last_name,
            )
            login(request, user)
            messages.success(request, 'Compte cree avec succes.')
            return redirect('/')

    return render(request, 'accounts/register.html')

from django.contrib.auth.decorators import login_required
from django.contrib.auth import update_session_auth_hash

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
            from django.contrib import messages
            messages.success(request, 'Informations mises à jour avec succès.')
        elif action == 'change_password':
            from django.contrib import messages
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
                from django.contrib import messages
                messages.success(request, 'Mot de passe changé avec succès.')
        return redirect('accounts:profil')

    reports = UploadedReport.objects.filter(user=request.user, processed=True)
    return render(request, 'accounts/profil.html', {
        'total_reports':    reports.count(),
        'total_incidents':  sum(r.total_incidents for r in reports),
        'total_unresolved': sum(r.unresolved_count for r in reports if r.unresolved_count),
    })