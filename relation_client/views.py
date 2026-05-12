from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib.admin.views.decorators import staff_member_required
from django.contrib.auth.models import User
from django.contrib import messages
from django.http import HttpResponse
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from .models import ReponseClient

def login_view(request):
    if request.user.is_authenticated:
        return redirect('dashboard')
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = authenticate(request, username=username, password=password)
        if user:
            login(request, user)
            return redirect('dashboard')
        else:
            messages.error(request, 'Identifiant ou mot de passe incorrect.')
    return render(request, 'relation_client/login.html')

def logout_view(request):
    logout(request)
    return redirect('login')

@login_required(login_url='login')
def dashboard(request):
    reponses = ReponseClient.objects.all()
    total = reponses.count()

    def calc(qs):
        t = qs.count()
        if t == 0:
            return {'nps': 0, 'csat': 0, 'ces': 0, 'total': 0}
        promoteurs  = qs.filter(categorie_nps='Promoteur').count()
        detracteurs = qs.filter(categorie_nps='Detracteur').count()
        nps  = round((promoteurs/t*100) - (detracteurs/t*100), 1)
        csat = round(qs.filter(csat='Satisfait').count()/t*100, 1)
        eleve  = qs.filter(ces='Effort eleve').count()
        faible = qs.filter(ces='Effort faible').count()
        ces = round((eleve/t*100) - (faible/t*100), 1)
        return {'nps': nps, 'csat': csat, 'ces': ces, 'total': t}

    context = {
        'global_stats': calc(reponses),
        'vie_stats':    calc(reponses.filter(canal__icontains='VIE')),
        'iard_stats':   calc(reponses.filter(canal__icontains='IARD')),
        'total':        total,
        'joints':       reponses.filter(statut_appel='Joint').count(),
        'injoignables': reponses.filter(statut_appel='Injoignable').count(),
    }
    return render(request, 'relation_client/dashboard.html', context)

@login_required(login_url='login')
def saisie(request):
    if request.method == 'POST':
        try:
            nps_score = request.POST.get('nps_score')
            nps_score = int(nps_score) if nps_score else None
            ReponseClient.objects.create(
                cc                   = request.user,
                date_enregistrement  = request.POST.get('date_enregistrement') or date.today(),
                nom_prenoms          = request.POST.get('nom_prenoms'),
                contact              = request.POST.get('contact'),
                numero_police        = request.POST.get('numero_police'),
                objet_visite         = request.POST.get('objet_visite'),
                canal                = request.POST.get('canal'),
                date_appel           = request.POST.get('date_appel') or date.today(),
                statut_appel         = request.POST.get('statut_appel'),
                nps_score            = nps_score,
                categorie_nps        = request.POST.get('categorie_nps', ''),
                csat                 = request.POST.get('csat', ''),
                ces                  = request.POST.get('ces', ''),
                motif_insatisfaction = request.POST.get('motif_insatisfaction'),
                observations         = request.POST.get('observations'),
            )
            messages.success(request, 'Reponse enregistree avec succes ✅')
            return redirect('saisie')
        except Exception as e:
            messages.error(request, f'Erreur : {e}')
    return render(request, 'relation_client/saisie.html')

@login_required(login_url='login')
def liste(request):
    reponses = ReponseClient.objects.all().order_by('-date_appel')
    canal  = request.GET.get('canal')
    statut = request.GET.get('statut')
    nps    = request.GET.get('nps')
    if canal:  reponses = reponses.filter(canal=canal)
    if statut: reponses = reponses.filter(statut_appel=statut)
    if nps:    reponses = reponses.filter(categorie_nps=nps)
    return render(request, 'relation_client/liste.html', {'reponses': reponses})

@login_required(login_url='login')
def indicateurs(request):
    reponses = ReponseClient.objects.all()

    def calc(qs):
        t = qs.count()
        if t == 0:
            return None
        promoteurs  = qs.filter(categorie_nps='Promoteur').count()
        detracteurs = qs.filter(categorie_nps='Detracteur').count()
        passifs     = qs.filter(categorie_nps='Passif').count()
        nps  = round((promoteurs/t*100) - (detracteurs/t*100), 1)
        csat = round(qs.filter(csat='Satisfait').count()/t*100, 1)
        eleve  = qs.filter(ces='Effort eleve').count()
        faible = qs.filter(ces='Effort faible').count()
        ces = round((eleve/t*100) - (faible/t*100), 1)
        return {
            'nps': nps, 'csat': csat, 'ces': ces, 'total': t,
            'promoteurs': promoteurs, 'detracteurs': detracteurs,
            'passifs': passifs,
            'joints':      qs.filter(statut_appel='Joint').count(),
            'injoignables':qs.filter(statut_appel='Injoignable').count(),
            'pas_reponse': qs.filter(statut_appel='Pas de reponse').count(),
        }

    tableau = [
        {'perimetre': 'VIE',    'stats': calc(reponses.filter(canal__icontains='VIE'))},
        {'perimetre': 'IARD',   'stats': calc(reponses.filter(canal__icontains='IARD'))},
        {'perimetre': 'GLOBAL', 'stats': calc(reponses)},
    ]
    return render(request, 'relation_client/indicateurs.html', {'tableau': tableau})

@login_required(login_url='login')
def export(request):
    reponses = ReponseClient.objects.all().order_by('-date_appel')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reponses Clients"
    entetes = [
        'Canal', 'Date Enregistrement', 'CC',
        'Nom et Prenoms', 'Contact', 'N Police',
        'Objet Visite', 'Date Appel', 'Statut Appel',
        'Score NPS', 'Categorie NPS', 'CSAT', 'CES',
        'Motif Insatisfaction', 'Observations'
    ]
    for col, entete in enumerate(entetes, start=1):
        cell = ws.cell(row=1, column=col, value=entete)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='E30613')
        cell.alignment = Alignment(horizontal='center')
    for row, r in enumerate(reponses, start=2):
        ws.cell(row=row, column=1,  value=r.canal)
        ws.cell(row=row, column=2,  value=str(r.date_enregistrement))
        ws.cell(row=row, column=3,  value=r.cc.username)
        ws.cell(row=row, column=4,  value=r.nom_prenoms)
        ws.cell(row=row, column=5,  value=r.contact)
        ws.cell(row=row, column=6,  value=r.numero_police)
        ws.cell(row=row, column=7,  value=r.objet_visite)
        ws.cell(row=row, column=8,  value=str(r.date_appel))
        ws.cell(row=row, column=9,  value=r.statut_appel)
        ws.cell(row=row, column=10, value=r.nps_score)
        ws.cell(row=row, column=11, value=r.categorie_nps)
        ws.cell(row=row, column=12, value=r.csat)
        ws.cell(row=row, column=13, value=r.ces)
        ws.cell(row=row, column=14, value=r.motif_insatisfaction)
        ws.cell(row=row, column=15, value=r.observations)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="SUNU_Reponses.xlsx"'
    wb.save(response)
    return response

@staff_member_required(login_url='login')
def gestion_utilisateurs(request):
    if request.method == 'POST':
        username   = request.POST.get('username')
        password   = request.POST.get('password')
        first_name = request.POST.get('first_name')
        last_name  = request.POST.get('last_name')
        email      = request.POST.get('email')
        is_staff   = request.POST.get('is_staff') == 'on'
        if User.objects.filter(username=username).exists():
            messages.error(request, f'Le nom utilisateur {username} existe deja.')
        else:
            User.objects.create_user(
                username=username, password=password,
                first_name=first_name, last_name=last_name,
                email=email, is_staff=is_staff
            )
            messages.success(request, f'Compte {username} cree avec succes ✅')
    utilisateurs = User.objects.all().order_by('-date_joined')
    return render(request, 'relation_client/utilisateurs.html', {'utilisateurs': utilisateurs})

@staff_member_required(login_url='login')
def toggle_utilisateur(request, user_id):
    user = get_object_or_404(User, id=user_id)
    if user != request.user:
        user.is_active = not user.is_active
        user.save()
        statut = 'active' if user.is_active else 'desactive'
        messages.success(request, f'Compte {user.username} {statut} ✅')
    return redirect('gestion_utilisateurs')