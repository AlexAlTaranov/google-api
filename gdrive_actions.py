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

from init_google_api_services import init_gdrive_api_sevice, init_sheets_api_sevice
import io
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import os

from pprint import pprint


def create_folder(folder_name):
    """
    Функция создаёт папку в корневой директории

    Входные параметры:
     - folder_name      - имя папки
     
    """
    service = init_gdrive_api_sevice()

    file_metadata = {
    'name': folder_name,
    'mimeType': 'application/vnd.google-apps.folder'
    }
    file = service.files().create(body=file_metadata, fields='id').execute()
    return file.get('id')

def move_file(file_id, target_folder_id):
    """
    Функция перемещает файл (папку) из корневой директории в другую папку

    Входные параметры:
     - file_id               - id перемещаемого файла (папки)
     - target_folder_id      - id папки, куда перемещается
     
    """
    service = init_gdrive_api_sevice()

    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    file = service.files().update(fileId=file_id,
        addParents=target_folder_id,
        removeParents=previous_parents,
        fields='id, parents').execute()

# def find_folder(folder_name, drive_id):
#     service = init_gdrive_api_sevice()

#     folder_names_dict = {}
#     page_token = None
#     try:
#         while True:
#             response = service.files().list(q="mimeType='application/vnd.google-apps.folder'",
#                 driveId =drive_id,
#                 spaces='drive',
#                 fields='nextPageToken, files(id, name)',
#                 pageToken=page_token).execute()
#             for file in response.get('files', []):
#                 print('Found file: %s (%s)' % (file.get('name'), file.get('id')))
#             folder_names_dict[file.get('name')] = file.get('id')
#             page_token = response.get('nextPageToken', None)
#             if page_token is None:
#                 break
#         if folder_name in folder_names_dict.keys():
#             return folder_names_dict[folder_name]
#         else:
#             return None
#     except Exception as e:
#         print("!!!!!!!!!!!")



def print_folders_in_folder():
    """Print folders belonging to a folder.

    """
    service = init_gdrive_api_sevice()

    folder_id = "1oCRSDdK0lSxKi0buv7N_gKvJMpAs7keb"
    query = f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder'"
    response = service.files().list(q=query).execute()
    pprint(response)

        
def delete_file(file_id):
  """
  Функция удалает файл (папку) из директории исходя из id файла (папки)

  Входные параметры:
     - file_id      - id файла (папки)

  """
  service = init_gdrive_api_sevice()
  try:
    service.files().delete(fileId=file_id).execute()
  except Exception as e:
    print(e)

def find_folder_by_name(folder_name, drive_id):
    """
    Эта функция ищет папку по имени, находящуюся в другой папке.

    Входные параметры:
     - folder_name      - имя папки

    Выходные параметры:
    Если папка с указанным именем есть, то функция возвращает id этой папки.
    В случае, когда такой папки нет, возвращается None.

    """
    service = init_gdrive_api_sevice()

    
    query = f"'{drive_id}' in parents and mimeType = 'application/vnd.google-apps.folder'"
    response = service.files().list(q=query).execute()
    name_list =  [item['name'] for  item in response['files']]
    try:
        folder_inlist_index = name_list.index(folder_name)
    except ValueError as e:
        print('None')
        return None
    else:
        print(response['files'][folder_inlist_index]['id'])
        return response['files'][folder_inlist_index]['id']

def show_all_files():
    service = init_gdrive_api_sevice()
    response = service.files().list().execute()
    pprint(response)

def show_all_files_in_folder(folder_id):
    service = init_gdrive_api_sevice()
    query = f"'{folder_id}' in parents"
    response = service.files().list(q=query).execute()
    pprint(response)

def download_file(file_id):
    """
    Эта функция скачивает файлы с расширением .doc .docx

    Входные параметры:
     - file_id      - id файла

    Выходные параметы:
    Функция возвращает имя скачанного файла (параметр file_name)

    Описание работы:
    Проверяется тип указанного файла: если он относится к .doc или .docx, то
    этот файл побайтово скачивается с сохранением имени в исполняемую директорию

    """
    service = init_gdrive_api_sevice()
    response = service.files().get(fileId=file_id, fields="name, mimeType").execute()
    print('Скачивание файла ...')
    # mimeType for *.docx and *.doc files
    docx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    doc = "application/msword"
    if response['mimeType'] in [doc, docx]:
        file_name = response['name']
        request = service.files().get_media(fileId=file_id)
        fh = io.FileIO(file_name, mode='wb')
        downloader = MediaIoBaseDownload(fh, request, chunksize=1024*1024)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            if status:
                print("Скачивание файла %d%%." % int(status.progress() * 100))
            print("Файл скачан")
        return file_name

def upload_doc_file(file_name):
    """
    Эта функция загружает файлы (на данный момент .doc или .docx) на корневую 
    папку сервисного аккаунта

    Входные параметры:
     - file_name      - имя файла

    Выходные параметы:
    Функция возвращает id файла или None, если расширение ещё не поддерживается программой

    Описание работы:
    Проверяется тип указанного файла: если он относится к .doc или .docx, то
    этот файл побайтово скачивается с сохранением имени в исполняемую директорию

    """
    extentions_dict = {
    'docx' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'doc' : 'application/msword',
    'jpg' : 'image/jpeg',
    'pdf' : 'application/pdf'
    }

    docx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    doc = "application/msword"

    file_name_arr = file_name.split('.')
    service = init_gdrive_api_sevice()
    file_metadata = {'name': file_name}

    if file_name_arr[1] in extentions_dict.keys():
        media = MediaFileUpload(file_name, mimetype=extentions_dict[file_name_arr[1]])
    else:
        raise ValueError("Unknown file extention")
        return None
    file = service.files().create(body=file_metadata,
        media_body=media,
        fields='id').execute()
    # print('File ID: %s' % file.get('id'))
    print('Файл загружен на диск')
    return file.get('id')

# def download_file_new_name(file_id, sheet_name):
#     """
#     Эта функция скачивает файлы с расширением .doc .docx

#     Входные параметры:
#      - file_id      - id файла

#     Выходные параметы:
#     Функция возвращает имя скачанного файла (параметр file_name)

#     Описание работы:
#     Проверяется тип указанного файла: если он относится к .doc или .docx, то
#     этот файл побайтово скачивается с сохранением имени в исполняемую директорию

#     """
#     service = init_gdrive_api_sevice()
#     response = service.files().get(fileId=file_id, fields="name, mimeType").execute()
#     # mimeType for *.docx and *.doc files
#     docx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#     doc = "application/msword"
#     if response['mimeType'] in [doc, docx]:
#         file_name = response['name']
#         request = service.files().get_media(fileId=file_id)
#         fh = io.FileIO(file_name, mode='wb')
#         downloader = MediaIoBaseDownload(fh, request, chunksize=1024*1024)
#         done = False
#         while done is False:
#             status, done = downloader.next_chunk()
#             if status:
#                 print("Download %d%%." % int(status.progress() * 100))
#             print("Download Complete!")
#         return sheet_name + '__' + file_name

# def upload_doc_file(file_name):
#     """
#     Эта функция загружает файлы (на данный момент .doc или .docx) на корневую 
#     папку сервисного аккаунта

#     Входные параметры:
#      - file_name      - имя файла

#     Выходные параметы:
#     Функция возвращает id файла или None, если расширение ещё не поддерживается программой

#     Описание работы:
#     Проверяется тип указанного файла: если он относится к .doc или .docx, то
#     этот файл побайтово скачивается с сохранением имени в исполняемую директорию

#     """

#     docx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#     doc = "application/msword"

#     file_name_arr = file_name.split('.')
#     print(file_name_arr)
#     service = init_gdrive_api_sevice()
#     file_metadata = {'name': file_name}
#     if file_name_arr[1] == 'doc':
#         print('if')
#         print(file_name_arr[1])
#         media = MediaFileUpload(file_name, mimetype=doc)
#     elif file_name_arr[1] == 'docx':
#         print('elif')
#         print(file_name_arr[1])
#         media = MediaFileUpload(file_name, mimetype=docx)
#     else:
#         print(file_name_arr[1])
#         print('else')
#         return None
#     file = service.files().create(body=file_metadata,
#         media_body=media,
#         fields='id').execute()
#     print('File ID: %s' % file.get('id'))
#     return file.get('id')




# def upload_file(file_name):
#     service = init_gdrive_api_sevice()

#     file_metadata = {'name': file_name}
#     media = MediaFileUpload(file_name, mimetype='image/jpeg')
#     file = service.files().create(body=file_metadata,
#         media_body=media,
#         fields='id').execute()
#     print('File ID: %s' % file.get('id'))
#     return file.get('id')
