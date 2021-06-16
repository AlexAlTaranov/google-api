FROM python:latest

ARG builtime_spreadsheet_id='_'
ARG builtime_target_folder_id='_'
ARG builtime_sheet_name='_'
ARG builtime_dogovory_sheet_name='_'

ENV SPREADSHEET_ID=$builtime_spreadsheet_id
ENV TARGET_FOLDER_ID=$builtime_target_folder_id
ENV SHEET_NAME=$builtime_sheet_name
ENV DOGOVORY_SHEET_NAME=$builtime_dogovory_sheet_name

RUN apt-get update
RUN pip3 install python-docx
RUN pip3 install pandas
RUN pip3 install memory_profiler
RUN pip3 install google-api-python-client oauth2client google-auth-httplib2 google-auth-oauthlib

COPY client_special_actions.py /
COPY credentials.json /
COPY doc_handler.py /
COPY gdrive_actions.py /
COPY gspreadsheet_actions.py /
COPY init_google_api_services.py /
COPY file_service.py /

ENTRYPOINT ["python", "./file_service.py"]
