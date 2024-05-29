import re
import openpyxl
from openpyxl import Workbook

path = '/Users/nikolajtukmaceva/Downloads/Все_категории.xlsx'

# Категории для удаления
categories_to_remove = ['Одежда', 'подгузники детские', 'автохимия - масло моторное', 'средство для посудомоечных машин', 'Корм сухой для собак', 'Корм сухой для кошек', 'Ноутбуки и компьютеры', 'Обувь', 'Детское питание', 'мебель', 'смартфоны и планшеты', 'аптека', 'крупная бытовая техника', 'детская одежда', 'напитки', 'бакалея и кондитерские изделия']


#Метод берет любое значение ячейки и возвращает int 
def excel_cell_to_int(cell_value):
    # Проверяем, не пустая ли ячейка
    if cell_value is None:
        return 0
    
    # Преобразуем значение в строку, если оно не строка
    if not isinstance(cell_value, str):
        cell_value = str(cell_value)
    
    # Удаляем все нецифровые символы, кроме запятых и точек (для чисел с плавающей точкой)
    cell_value = re.sub(r'[^\d.,]', '', cell_value)
    cell_value = cell_value.replace(" ","")

    # Если после удаления спецсимволов ничего не осталось или осталась только точка, возвращаем 0
    if not cell_value or cell_value == '.':
        return 0
    
    # Если значение содержит запятые или точки, то это может быть число с плавающей точкой
    if ',' in cell_value or '.' in cell_value:
        # Заменяем запятые на точки, чтобы не потерять дробную часть
        cell_value = cell_value.replace(',', '.')
        # Преобразуем в float и округляем до ближайшего целого
        return int(float(cell_value))
    else:
        # Преобразуем в int
        return int(cell_value)

#Метод берет любое значение ячейки и возвращает str
def process_excel_cell(cell_value):
    # Проверяем, является ли значение пустым
    if not cell_value:
        return ""
    
    # Проверяем, является ли значение пустой строкой
    if isinstance(cell_value, str) and cell_value.strip() == "":
        return ""
    
    # Проверяем, является ли значение числовым
    if isinstance(cell_value, (int, float)):
        return str(cell_value).lower()
    
    # Если значение не пустое, не пустая строка и не число, возвращаем его в нижнем регистре
    return str(cell_value).lower()

#Метод принимает на вход путь к исходной таблице и фильтрует по заданным параметрам
def filterProducsByParametrs(path):

    # Открываем существующий файл Excel
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    
    # Создаем новый рабочий лист для результата
    wb_result = Workbook()
    ws_result = wb_result.active
    
    # Проходимся по каждой строке данных
    for row in ws.iter_rows(values_only=True):

        # Получение значения из таблицы 
        root_category = process_excel_cell(row[0])
        category_2_lvl = process_excel_cell(row[1])
        category_3_lvl = process_excel_cell(row[2])
        category_4_lvl = process_excel_cell(row[3])
        categories = [root_category, category_2_lvl, category_3_lvl, category_4_lvl]
    
        average_selling_price = excel_cell_to_int(row[11])
        average_number_of_orders_per_day = excel_cell_to_int(row[14])
        number_of_catalog_views_in_28_days= excel_cell_to_int(row[18])
        number_of_product_card_views_in_28_days= excel_cell_to_int(row[19])
    
        # Проверка категорий 
        match_found = False
        for category in categories:
            if category.lower() in [cat.lower() for cat in categories_to_remove]:
                match_found = True
                break
    
        if match_found:
            continue
    
        # Проверка среднего числа продаж в день
        if not(average_number_of_orders_per_day > 5):
            continue
    
        # Проверка стоимости товара
        if not(average_selling_price > 400 and average_selling_price < 5500):
            continue  
    
        # Проверка просмотров каталога за 28 дней
        if not(number_of_catalog_views_in_28_days > 10000):
            continue
        
        # Проверка просмотров карточки товара за 28 дней
        if not(number_of_product_card_views_in_28_days > 1000):
            continue

        # Если товар прошел все проверки, добавляем его в новый список
        ws_result.append(row)
    
        # Форматируем вывод    
        print(f"Категория: {root_category:<20} | Стоимость: {average_selling_price:>7} | Продажи в день: {average_number_of_orders_per_day:>7} | Просмотры каталога: {number_of_catalog_views_in_28_days:>7} | Просмотры карточки: {number_of_product_card_views_in_28_days:>7}")
    
    # Сохраняем результат в новый файл Excel
    output = '/Users/nikolajtukmaceva/Downloads/Фильтр_по_параметрам.xlsx'
    wb_result.save(output)
    print(output)



filterProducsByParametrs(path)