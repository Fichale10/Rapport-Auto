from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages


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