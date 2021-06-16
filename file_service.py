from gspreadsheet_actions import column_check, read_row, write_to_cell
from gdrive_actions import *
from client_special_actions import make_folder_name, make_file_name, choose_template_type, create_file_from_template, create_file_url, create_folder_url
import os
import time
import traceback


while True:
    try:
        # Задаём параметры наблюдаемого объекта и включаем наблюдение
        spreadsheet_id = os.environ.get('SPREADSHEET_ID')
        pressed_button_rows_list = column_check(spreadsheet_id, os.environ.get('SHEET_NAME'), 'AS')
        # обрабатываем нажатие
        for item in pressed_button_rows_list:
            
            # Пишем в контрольную ячейку 'AQ9' номер строки, где произошли изменения
            write_to_cell(spreadsheet_id, os.environ.get('SHEET_NAME'), str(9), 'AS', str(item))
            
            # Проверка на наличие данных для создания папки и папки
            result = read_row(spreadsheet_id, os.environ.get('SHEET_NAME'),  'A', 'AX', str(item))
            result = result['valueRanges'][0]['values'][0]
            if result[6] == '':
                print('Не хватает названия магазина')
                continue
            if  result[7] == '':
                print('Не хватает юр названия магазина')
                continue
            if  result[14] == '':
                print('Не хватает email')
                continue
            if  result[28] == '':
                print('Не хватает типа договора')
                continue
            if  result[30] == '':
                print('Не хватает ИНН для создания папки')
                continue
            if  result[31] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[32] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[33] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[34] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[35] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[36] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[37] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[38] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[39] == '':
                print('Не хватает данных для создания файла')
                continue
            if  result[40] == '':
                print('Не хватает данных для создания файла')
                continue
            
            # Задаём имя файла договора
            file_name = make_file_name(spreadsheet_id, os.environ.get('DOGOVORY_SHEET_NAME'), 'A', os.environ.get('SHEET_NAME'), str(item), 'AT')
            # вписываем его в ячейку и запрашиваем строку
            try:
                row_values_list = read_row(spreadsheet_id, os.environ.get('SHEET_NAME'),  'A', 'AX', str(item))
                print(len(row_values_list))
            except ConnectionError as e:
                if str(e) == "The spreadsheet doesn't respond":
                    print("Таблица недоступна")

            # Выбираем шаблон договора
            template_file_id = choose_template_type(spreadsheet_id, row_values_list)

            # Считываем контрольные слова для замены в шаблоне
            control_row_values_list = read_row(spreadsheet_id, os.environ.get('SHEET_NAME'),  'A', 'AX', str(8))

            # Создаём файл договора
            new_doc_id = create_file_from_template(template_file_id, control_row_values_list, row_values_list, file_name)

            # Проверяем есть ли папки с таким имемем, если есть - удаляем
            folder_name = make_folder_name(row_values_list)
            folder_id = find_folder_by_name(folder_name, os.environ.get('TARGET_FOLDER_ID'))
            if folder_id == None:
                # Создаем новую папку с указанным именем и перемещаем её в корневую папку проекта
                folder_id = create_folder(folder_name)
                print('создал новую папку {}'.format(folder_name))
            
            
            move_file(folder_id, os.environ.get('TARGET_FOLDER_ID'))
            print('переместил её в корневую папку проекта')
            # перемещаем в папку договора сам договор
            move_file(new_doc_id, folder_id)
            print('переместил в папку клиента договор')
            # Создаём ссылку на файл договора
            formula_url = create_file_url(new_doc_id, file_name)
            write_to_cell(spreadsheet_id, os.environ.get('SHEET_NAME'), str(item), 'AT', formula_url)

            # Создаем ссылку на новую папку 
            formula_url =create_folder_url(folder_id)
            # Вставляем ссылку в таблицу
            write_to_cell(spreadsheet_id, os.environ.get('SHEET_NAME'), str(item), 'AQ', formula_url)



            
            
    except Exception as e:
        print('Ошибка в общем цикле')
        # print(e)
        traceback.print_exc()
