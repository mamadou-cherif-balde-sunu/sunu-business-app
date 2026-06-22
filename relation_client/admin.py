from django.contrib import admin
from .models import ReponseClient, DemandeModification

@admin.register(ReponseClient)
class ReponseClientAdmin(admin.ModelAdmin):
    list_display  = [
        'nom_prenoms', 'cc',
        'date_enregistrement', 'statut_appel',
        'categorie_nps', 'csat', 'ces'
    ]
    list_filter   = [
        'statut_appel',
        'categorie_nps', 'csat', 'ces'
    ]
    search_fields = ['nom_prenoms', 'contact', 'numero_police']
    ordering      = ['-date_enregistrement']

@admin.register(DemandeModification)
class DemandeModificationAdmin(admin.ModelAdmin):
    list_display  = ['demandeur', 'saisie', 'champ', 'nouvelle_valeur', 'statut', 'date_demande']
    list_filter   = ['statut']
    ordering      = ['-date_demande']