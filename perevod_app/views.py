from openpyxl import load_workbook
from django.shortcuts import render
from .models import Translation
from django.views import View
from django.views.generic.list import ListView
from django.views.generic.detail import DetailView
from django.views.generic.edit import CreateView, UpdateView, DeleteView
from django.urls import reverse_lazy
from .models import Translation


class WordsBaseView(View):
    model = Translation
    fields = '__all__'
    success_url = reverse_lazy('all')

class WordsListView(WordsBaseView, ListView):
    """Список всех слов"""

class WordsDetailView(WordsBaseView, DetailView):
    """Определенная пара слов"""

class WordsCreateView(WordsBaseView, CreateView):
    """Создание новой пары слов"""

class WordsUpdateView(WordsBaseView, UpdateView):
    """Редактирование пары слов"""

class WordsDeleteView(WordsBaseView, DeleteView):
    """Удаление пары слов"""

'''class SearchWord(WordsBaseView):
    def get(request):'''

def search_word(request):
    # загрузка файла Excel
    wb = load_workbook(filename='translations.xlsx')

    # выбор нужного листа в файле
    sheet = wb['Sheet1']
    add_words_from_exel()
    # получение данных из ячеек
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if cell == request.GET.get('search_word'):
                result = row[0] , row[1]  # значение из соседней колонки
                return render(request, 'search_word.html', {'result': result})


    result = 'Слово не знайдено'
    return render(request, 'perevod_app/search_word.html', {'result': result})

def add_words_from_exel():
    '''сохранение всех записей с exel в БД'''
    wb = load_workbook(filename='translations.xlsx')
    sheet = wb['Sheet1']
    for row in sheet.iter_rows(values_only=True):
        try:
            trans_bd = Translation()
            trans_bd.en=row[0]
            trans_bd.uk=row[1]
            trans_bd.save()
        except:
            continue

def add_words_from_exel_in_DB():
    '''сохранение всех записей с БД в еxel'''
