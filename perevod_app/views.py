from openpyxl import load_workbook
from django.shortcuts import render

def search_word(request):
    # загрузка файла Excel
    wb = load_workbook(filename='translations.xlsx')

    # выбор нужного листа в файле
    sheet = wb['Sheet1']

    # получение данных из ячеек
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if cell == request.GET.get('search_word'):
                result = row[0] , row[1]  # значение из соседней колонки
                return render(request, 'search_word.html', {'result': result})


    result = 'Слово не знайдено'
    return render(request, 'search_word.html', {'result': result})
