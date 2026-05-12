from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    # C'est ici qu'on branche l'application relation_client
    path('', include('relation_client.urls')), 
]