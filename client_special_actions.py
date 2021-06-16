"""
    Модуль предназначен для работы с Google Sheets API v4. Модуль содержит фукнции чтения 
    и обработки информации из гугл таблиц.

    Для работы с модулем необходимо предварительно настроить сервис-аккаунт google для работы с 
    Google Sheets API v4. Таблица, к которой обращается функция должна предварительно быть 
    расшарена для данного сервис-аккаунта. Данные сервис-аккаунта должны быть сохранены в файле 
    credentials.json (файл скачивается в финале настройке сервис-аккаунта).

    Автор: Александр Таранов
    email: alex.al.taranov@gmail.com
    дата: 09.05.2021
"""

from init_google_api_services import init_sheets_api_sevice
from gspreadsheet_actions import write_to_cell
from gdrive_actions import download_file, upload_doc_file
from doc_handler import replace_text
from pprint import pprint
import os

def make_folder_name(row_values_list):
    row_values_list = row_values_list['valueRanges'][0]['values'][0]
    folder_name = row_values_list[6] + '_' + row_values_list[7] + '_' + str(row_values_list[30])
    return folder_name

def make_file_name(spreadsheet_id, sheet_dogovor, column_dogovor, sheet_manager, row_manager, column_manager):
    
    controlled_ranges = sheet_dogovor + "!" + column_dogovor  + ":" + column_dogovor
    value_render_option = 'UNFORMATTED_VALUE'
    service = init_sheets_api_sevice()
    request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, 
            ranges=controlled_ranges, 
            valueRenderOption=value_render_option)
    response = request.execute()

    last_num = response['valueRanges'][0]['values'][-1][0]
    list_len = len(response['valueRanges'][0]['values'])
    dogovor_num = last_num + 1
    write_to_cell(spreadsheet_id, sheet_dogovor, str(list_len + 1), column_dogovor, dogovor_num)
    write_to_cell(spreadsheet_id, sheet_manager, row_manager, column_manager, dogovor_num)

    file_name = 'Договор №' + str(dogovor_num)
    print(file_name)
    return file_name

def choose_template_type(spreadsheet_id, row_values_list):
    url_and_template_type_ranges = 'СВОДНАЯ!AS1:AT4'
    value_render_option = 'UNFORMATTED_VALUE'
    service = init_sheets_api_sevice()
    request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, 
        ranges=url_and_template_type_ranges,
        valueRenderOption=value_render_option)
    response = request.execute()
    pprint(response)
    response = response['valueRanges'][0]['values']
    pprint(response)
    template_dict = {}
    for item in response:
        template_dict[item[1]] = item[0]
    row_values_list = row_values_list['valueRanges'][0]['values'][0]
    template_type = row_values_list[28]
    # pprint(template_dict)
    url = template_dict[template_type]
    url_arr = url.split('/')
    template_id = url_arr[-2]
    print('Тип договора: "{}"'.format(template_type))
    return template_id

def create_file_from_template(template_file_id, control_row_values_list, row_values_list, new_file_name):
    print('create_file_from_template')
    control_row_values_list = control_row_values_list['valueRanges'][0]['values'][0]
    row_values_list = row_values_list['valueRanges'][0]['values'][0]
    zip_values = list(zip(control_row_values_list, row_values_list))
    zip_values = [item for item in zip_values if item[0] != '']
    print(zip_values)
    file_name = download_file(template_file_id)
    new_name = replace_text(zip_values, file_name, new_file_name)
    print('os.remove(file_name)')
    print(file_name)
    os.remove(file_name)
    file_id = upload_doc_file(new_name)
    print('os.remove(new_name)')
    os.remove(new_name)
    return file_id

def create_folder_url(folder_id):
    # =ГИПЕРССЫЛКА("https://drive.google.com/drive/folders/1-OVIXcRQs5Ob414Lun0uX2GDoIByGWEP";"ПАПКА")
    formula_url = '=ГИПЕРССЫЛКА("https://drive.google.com/drive/folders/' + folder_id +  '";"ПАПКА")'
    print(formula_url)
    return formula_url

def create_file_url(file_id, file_name):
    # =ГИПЕРССЫЛКА("https://drive.google.com/file/d/1pjSy2msIKD_NhzohRUOk5tgysM4rfaDv/view?usp=drivesdk";"ЧЕК")
    file_name_arr = file_name.split('.')
    file_name_arr = file_name_arr[0].split('№')
    dogovor_num = file_name_arr[1]
    formula_url = '=ГИПЕРССЫЛКА("https://drive.google.com/file/d/' + file_id + '/view?usp=drivesdk";"' + dogovor_num + '")'
    print(formula_url)
    return formula_url
