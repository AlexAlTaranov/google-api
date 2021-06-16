from googleapiclient import discovery
from google.oauth2 import service_account

def init_gdrive_api_sevice():
    scopes = ['https://www.googleapis.com/auth/drive']
    credentials_file = 'credentials.json'
    credentials = service_account.Credentials.from_service_account_file(credentials_file, 
        scopes=scopes)
    service = discovery.build('drive', 'v3', credentials=credentials)
    return service

def init_sheets_api_sevice():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    credentials_file = 'credentials.json'
    credentials = service_account.Credentials.from_service_account_file(credentials_file, 
        scopes=scopes)
    service = discovery.build('sheets', 'v4', credentials=credentials)
    return service