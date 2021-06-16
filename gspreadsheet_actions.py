"""
    Модуль предназначен для работы с Google Sheets API v4. Модуль содержит фукнции чтения 
    и обработки информации из гугл таблиц.

    Для работы с модулем необходимо предварительно настроить сервис-аккаунт google для работы с 
    Google Sheets API v4. Таблица, к которой обращается функция должна предварительно быть 
    расшарена для данного сервис-аккаунта. Данные сервис-аккаунта должны быть сохранены в файле 
    credentials.json (файл скачивается в финале настройке сервис-аккаунта).

    Автор: Александр Таранов
    email: alex.al.taranov@gmail.com
    дата: 18.05.2021
"""
from init_google_api_services import init_sheets_api_sevice
import pandas as pd
import time
#
import memory_profiler


def column_check(spreadsheet_id, sheet_name, column_name, value_render_option='UNFORMATTED_VALUE'):
    """
    Краткое описание:
    Эта функция проверяет изменение состояния столбца или части столбца

    Входные параметры:
    Функция не имеет входных параметро
     - spreadsheet_id      - id гугл таблицы
     - sheet_name          - имя листа
     - column_name         - имя столбца (формат "A1 Notation")
     - value_render_option - формат считывания данных из ячейки. Это . 
                             Значение по умолчанию: 'UNFORMATTED_VALUE'. 
                             Варианты: 'FORMATTED_VALUE' - считывается значение ячейки в 
                                                           в формате ячейки, т.е. ячейка в 
                                                           фин формате скачается как $ 100.1
                                        'UNFORMATTED_VALUE' - скачивается только значение 
                                                              ячейки (без формата)
                                        'FORMULA' - скачивается формула ячейки, если ячейка
                                                    содержит формулу, или значение, если 
                                                    ячейка не содержит формулу.

    Выходные параметы:
    Возвращается список из номеров строк таблицы, в которых зафиксировано событие.

    Описание работы:
        Функция считывает значения ячеек проверяемого столбца, затем в бесконечом цикле снова читает
    значения ячеек и сравнивает с запомненым, если обнаруживается разнца, то возвращается список
    с номерами столбцов таблицы, где произошло изменение, но только если это изменение с False на True.
    (то есть если пользователь поставил ггалочку а не убрал)
        Функция делает проверку на ошибки, выводит имя и класс ошибки в терминал.Также функция 
    обрабатывает изменени количества строк в документе или ошибку ввода большего количества строк
    в параметре controlled_ranges.
        Функция использует предварительно настроенный сервис-аккаунт google для работы с 
    Google Sheets API v4. Таблица, к которой обращается функция должна предварительно расшарена для
    данного сервис-аккаунта. Данные сервис-аккаунта сохранены в файле credentials.json (файл 
    скачивается в финале настройке сервис-аккаунта)

    """
    controlled_ranges = sheet_name + "!" + column_name + ":" + column_name
    service = init_sheets_api_sevice()
    request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, 
            ranges=controlled_ranges, 
            valueRenderOption=value_render_option)
    while True:
        try:
            response = request.execute()
        except Exception as e:
            print(str(e) + str(type(e)))
            time.sleep(10)
        else:
            series = pd.Series(response['valueRanges'][0]['values'])
            break
    while True:
        time.sleep(5)
        request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, 
            ranges=controlled_ranges, 
            valueRenderOption=value_render_option)
        try:
            response = request.execute()
        except Exception as e:
            print(str(e) + str(type(e)))
            time.sleep(10)
        else:
            series_current = pd.Series(response['valueRanges'][0]['values'])

            try:
                difference = series_current.compare(series)
            except ValueError as e:
                if str(e) == "Can only compare identically-labeled Series objects":
                    print("Количество строк в документе изменилось")
            else:
                difference_list = difference.index.to_list()
                if len(difference_list) > 0:
                    # Проверяем были ли в изменениях True, если да, то реагируем
                    result_list = []
                    for item in difference_list:
                        if  series_current[item][0] == True:
                            result_list.append(item + 1)
                    if len(result_list) > 0:
                        print("Зафиксирована установка флагов в строках: " + ", ".join(str(item) for item in result_list))
                        return result_list
            finally:
                series = series_current
                print('Memory: ' + str(memory_profiler.memory_usage()) + 'MB' )


def read_row(spreadsheet_id, sheet_name, start_column_name, final_column_name, row_number):
    """
    Краткое описание:
    Эта функция считывает данные из строки (части строки) и возвращает их списком

    Входные параметры:
    Функция не имеет входных параметро
     - spreadsheet_id      - id гугл таблицы
     - start_column_name   - имя начального столбца в формате "A1 Notation" (начало 
                             считываемой строки)
     - final_column_name   - имя конечного столбца в формате "A1 Notation" (конец 
                             считываемой строки)
     - row_number          - номер считываемой строки (формат "A1 Notation", т.е 
                             реальный номер)

    Выходные параметы:
    Возвращается список из значений ячеек, входящих в строку, с сохранением типов данных

    Описание работы:
    Стандартными стредствами Sheets API v4 считывается часть строки. В теле программы
    присутствует проверка на наличие проблем со связью с таблицей. После пяти неудачных
    попыток отправить запрос функция бросает исключение
    """

    print('read_row')
    controlled_ranges = sheet_name + "!" + start_column_name + row_number + ":" + final_column_name + row_number
    value_render_option = 'UNFORMATTED_VALUE'

    service = init_sheets_api_sevice()

    request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, 
            ranges=controlled_ranges, 
            valueRenderOption=value_render_option)

    print('###')
    i = 5
    while i > 0:
        try:
            response = request.execute()
        except Exception as e:
            print(str(e) + str(type(e)))
            time.sleep(10)
        else:
            return response
        finally:
            i -= 1
    raise ConnectionError("The spreadsheet doesn't respond")

def write_to_cell(spreadsheet_id, sheet_name, row_number, column_name, value):
    service = init_sheets_api_sevice()
    values = [[value]]
    data = [{'range': sheet_name + "!" + column_name  + row_number,
        'values': values}]
    body = {'valueInputOption': 'USER_ENTERED', 'data': data}
    result = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, 
            body=body).execute()

