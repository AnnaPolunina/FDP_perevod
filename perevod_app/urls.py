from django.urls import path
from . import views

urlpatterns = [
    path('', views.search_word, name='search_word'),
    path('list/', views.WordsListView.as_view(), name='all'),
    path('create/', views.WordsCreateView.as_view(), name='word_create'),
    path('<int:pk>/update/', views.WordsUpdateView.as_view(), name='word_update'),
    path('<int:pk>/delete/', views.WordsDeleteView.as_view(), name='word_delete'),
]
