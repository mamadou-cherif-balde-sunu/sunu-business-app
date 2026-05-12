from django.contrib import admin
from django.urls import path
from relation_client import views # Import direct depuis views.py

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.dashboard, name='dashboard'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('saisie/', views.saisie, name='saisie'),
    path('liste/', views.liste, name='liste'),
    path('indicateurs/', views.indicateurs, name='indicateurs'),
    path('export/', views.export, name='export'),
]