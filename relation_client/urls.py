from django.urls import path
from . import views

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('saisie/', views.saisie, name='saisie'),
    path('liste/', views.liste, name='liste'),
    path('indicateurs/', views.indicateurs, name='indicateurs'),
    path('export/', views.export, name='export'),
    path('export/pptx/', views.export_pptx, name='export_pptx'),
    path('utilisateurs/', views.gestion_utilisateurs, name='gestion_utilisateurs'),
    path('utilisateurs/toggle/<int:user_id>/', views.toggle_utilisateur, name='toggle_utilisateur'),
    path('profil/', views.profil, name='profil'),
    path('saisie/modifier/<int:saisie_id>/', views.modifier_saisie, name='modifier_saisie'),
    path('saisie/demande/<int:saisie_id>/', views.demande_modification, name='demande_modification'),
    path('demandes/', views.demandes_admin, name='demandes_admin'),
]