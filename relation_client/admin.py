from django.contrib import admin
from .models import ReponseClient

@admin.register(ReponseClient)
class ReponseClientAdmin(admin.ModelAdmin):
    list_display = [
        'nom_prenoms', 'cc',
        'date_enregistrement', 'statut_appel',
        'categorie_nps', 'csat', 'ces'
    ]
    list_filter = [
        'statut_appel',
        'categorie_nps', 'csat', 'ces'
    ]
    search_fields = ['nom_prenoms', 'contact', 'numero_police']
    ordering = ['-date_enregistrement']