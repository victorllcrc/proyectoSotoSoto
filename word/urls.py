from django.urls import path
from . import views



urlpatterns = [
    path('', views.home, name='home'),
    path('proyecto/', views.cargaexcel, name='cargaexcel'),
]