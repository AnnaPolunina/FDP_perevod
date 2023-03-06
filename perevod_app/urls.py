from django.urls import path
from . import views

urlpatterns = [
    path('', views.search_word, name='search_word'),
]
