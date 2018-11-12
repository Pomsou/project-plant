from django.urls import path
from rango import views


urlpatterns = [
    path('', views.rango, name='rango'),
    path('about/', views.about, name='about'),
]
