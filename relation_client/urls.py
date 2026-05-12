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
    path('utilisateurs/', views.gestion_utilisateurs, name='gestion_utilisateurs'),
    path('utilisateurs/toggle/<int:user_id>/', views.toggle_utilisateur, name='toggle_utilisateur'),
]