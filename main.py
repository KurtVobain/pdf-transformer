import re
import os
import pdfplumber
import codecs
import pandas as pd

# Массив с номерами дела
number = list()
# Массив с датой дела
date = list()
# Массив с номером суда дела
court_number = list()

def info_coolect(pdf_file):
    """Собираем массивы с № приказов, датами и № участков"""

    # Читаем первую страницу pdf-документа в формате str
    pdf = pdfplumber.open(pdf_file)
    page = pdf.pages[0]
    text = page.extract_text()

    # Получаем искомую подстроку и получем массив ее слов,
    # из которого затем возьмем номер и дату
    string = r'судебный приказ[ \n]\S*[ \n]\№'
    second_string = r'исполнительный лист[ \n]\S*[ \n]\№'

    # Если не находится первая строка, значит документ содержит ключевую строку второго типа
    if re.search(string, text):
        target_string = [value for value in text[re.search(string, text).span()[1]:re.search(string, text).span()[1] + 35].split(' ') if value]

    elif re.search(second_string, text):
        target_string = [value for value in text[re.search(second_string, text).span()[1]:re.search(second_string, text).span()[1] + 35].split(' ') if value]

    else:
        print('Номер и/или дата судебного приказа не найдена')
        return next

    # Обработка одного исключения, буквы и номер судебного приказа
    # были рзделены
    if len(target_string[0].replace('\n', '')) >= 3:
        # Получаем номер судебного приказа
        number.append(target_string[0].replace('\n', ''))

        # Получаем дату судебного приказа
        date.append(target_string[2].replace(',', '').replace('\n', ''))

    # Если длина эл. меньше 3 символов, значит номер приказа разделен пробелом,
    # и разбит на 2 эл. в массиве. Склеиваем их и берем датой другой эл. массива
    else:
        # Получаем номер судебного приказа
        number.append(target_string[0].replace('\n', '') + target_string[1].replace('\n', ''))

        # Получаем дату судебного приказа
        date.append(target_string[3].replace(',', ''))


    # Получаем номер судебного участка
    if text.find('Судебный участок мирового судьи №') != -1:
        court_number.append(text[text.find('Судебный участок мирового судьи №') + 32:text.find(
            'Судебный участок мирового судьи №') + 38].split(' ')[1])

    elif text.find('Судебный участок №') != -1:
        court_number.append(text[text.find('Судебный участок №') + 19:text.find(
            'Судебный участок №') + 21].split()[0])

    else:
        print('Номер участка не найден')

    # Закрываем документ
    pdf.close()


# Массив в котором будут сохранены пути из файла
paths = list()

# Считываем пути из файла
with codecs.open('пути к папкам.txt', 'r', 'utf_8_sig' ) as f:
    for path in f.readlines():
        # Убираем с каждой строки символы перехода на нову строку и  возврата каретки
        # Добавляем в массив
        paths.append(path.replace('\n', '').replace('\r', ''))

# Проходим по всем указанным путям
for path in paths:
    # Проходим по всем папкам
    for root, dirs, files in os.walk(os.path.abspath(path)):
        # Проходим по всем файлам
        for file in files:
            print(os.path.join(root, file))
            info_coolect(os.path.join(root, file))


info_collection = pd.DataFrame()

# Собираем df с необходимой информацией
info_collection['№ приказа'] = number
info_collection['Дата'] = date
info_collection['№ участка'] = court_number

# Сохраняем таблицу excel
info_collection.to_excel('как угодно.xlsx')