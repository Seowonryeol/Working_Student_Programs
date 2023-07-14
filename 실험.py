from __future__ import print_function
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials
import io

def download_files(query):
    response = service.files().list(q=query, fields='files(id, name, parents)').execute()
    files = response.get('files', [])

    for file in files:
        file_id = file['id']
        file_name = file['name']
        file_path = get_file_path(file)
        
        if file_path:
            download_file(file_id, file_name, file_path)

def get_file_path(file):
    file_parents = file.get('parents', [])
    if len(file_parents) > 0:
        parent_id = file_parents[0]
        parent = service.files().get(fileId=parent_id, fields='id, name, parents').execute()
        parent_path = get_file_path(parent)
        if parent_path:
            return os.path.join(parent_path, file['name'])
    else:
        return ''

def create_folder(folder_path):
    try:
        os.mkdir(folder_path)
        print("폴더가 성공적으로 생성되었습니다.")
    except FileExistsError:
        print("이미 폴더가 존재합니다.")
    except Exception as e:
        print("폴더 생성에 실패했습니다."),str(e)

def download_file(file_id, file_name,folder_path):
    request = service.files().get_media(fileId=file_id)
    file_path=os.path.join(folder_path, file_name)
    fh = io.FileIO(file_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(f"다운로드 진행률: {int(status.progress() * 100)}%")

    print(f"{file_name} 다운로드 완료!")

def search_files(query):
    response = service.files().list(q=query).execute()
    files = response.get('files', [])
    return files
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

"""Shows basic usage of the Drive v3 API.
Prints the names and ids of the first 10 files the user has access to.
"""
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'dnjsfuf123456@gmail.com oauth 클라이언트.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('drive', 'v3', credentials=creds)
"""
    # Call the Drive v3 API
    results = service.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])



# 파일 이름이 "example.txt"인 파일 검색
"""
month=int(input("원하는 달을 입력하세요 (ex)5월의 경우 : 5) :"))
create_folder("./%d월 산학협력프로젝트 구글 드라이브 파일 다운로드" %month)
query = "'1DNU3_cwP4aYspy_9znGMTNwHTXHG6-az' in parents and name startsWith '2023' and name contains '5월' and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'" 
result = search_files(query)
for file in result:
    print(f"파일 이름: {file['name']}, 파일 ID: {file['id']}")
    # 파일 다운로드 실행
    download_file(file['name'], file['id'],"./%d월 산학협력프로젝트 구글 드라이브 파일 다운로드" %month)

"""
    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))
"""
