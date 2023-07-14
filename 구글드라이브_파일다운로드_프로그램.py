from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials
import io
import pytz
from datetime import datetime, timedelta
# If modifying these scopes, delete the file token.pickle.

#구글 드라이브에서 파일을 다운로드하는 프로그램입니다.
#원하는 달을 입력하면 1. 2023으로 시작하고 2. 원하는 달이 들어가고 3. weekly나 Weelky가 들어가는 4. docx 파일을 다운로드합니다.

SCOPES = ['https://www.googleapis.com/auth/drive']

downloaded_files_cnt=0 #다운로드된 파일 수, 전체 파일 수, 실패한 파일 수
files_cnt=0
downloaded_failed_files_cnt=0
failed_files=[]

def download_file(service,file_name, file_id, modified_time, folder_path): #파일을 다운로드하는 함수
    global downloaded_files_cnt
    global downloaded_failed_files_cnt
    request = service.files().get_media(fileId=file_id)
    file_path=os.path.join(folder_path, file_name)
    try: #파일 다운로드 시도
        fh = io.FileIO(file_path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
    except: #실패시
        print(file_name+" 파일 다운로드에 실패했습니다.")
        downloaded_failed_files_cnt+=1
        failed_files.append(file_name)
        return
        
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        downloaded_files_cnt+=1
        process_rate=((downloaded_files_cnt/files_cnt)*100).__round__(1)
        print(f"다운로드 진행률 ({downloaded_files_cnt}/{files_cnt}%%, {process_rate}) : {int(status.progress() * 100)}%")   

    modified_time = datetime.strptime(modified_time, "%Y-%m-%dT%H:%M:%S.%fZ") #구글 드라이브에서 가져온 시간을 로컬에 맞게 변환한 후 파일에 적용
    utc_timezone = pytz.timezone('UTC')
    modified_time = utc_timezone.localize(modified_time)
    modified_time= int(modified_time.timestamp())
    os.utime(file_path, (modified_time,modified_time))
    
def get_file_path(service, file_id): #파일 경로를 가져오는 함수
    file_resource = service.files().get(fileId=file_id, fields='name, parents').execute()
    file_name = file_resource.get('name')
    parent_ids = file_resource.get('parents', [])

    if parent_ids:
        parent_folder_id = parent_ids[0]
        parent_folder_path = get_file_path(service,parent_folder_id)
        return f"{parent_folder_path}/{file_name}"
    else:
        return file_name
    
def create_folders_from_path(path): #한 파일의 모든 상위 폴더를 생성하는 함수
    # 경로를 상위 폴더들로 분할
    folders = path.split('/')
    
    # 첫 번째 항목은 루트 폴더일 수도 있으므로 빈 문자열인 경우 제거
    if folders[0] == '':
        folders = folders[1:]
    
    # 각 폴더를 순서대로 생성
    current_path = ''
    for folder in folders:
        current_path = os.path.join(current_path, folder)
        if not os.path.exists(current_path):
            try:
                os.mkdir(current_path)
            except FileExistsError as e:
                print("",end="")
            except Exception as e:
                print("", str(e))
def main():
    global files_cnt
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)

    # Call the Drive v3 API
    month=int(input("원하는 달을 입력하세요 (ex)5월의 경우 : 5) :")) #달 입력, 쿼리문
    query = 'name starts with "2023" and name contains "%d월" and (name contains "weekly" or name contains "Weekly") and mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"'%month
    results=service.files().list(q=query, pageSize=1000, fields="nextPageToken, files(id, name, modifiedTime)").execute()
    items = results.get('files', [])
    create_folders_from_path("./%d월 산학협력프로젝트 구글파일 다운로드" %month) #폴더 생성

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
           print(u'{0}'.format(item['name']))
           files_cnt+=1
        print("총 %d개의 파일이 검색되었습니다." %files_cnt) #총 파일 수 출력

    for file in items:
        file_id=file['id']
    # 파일 리소스 가져오기
        file_path=get_file_path(service, file_id)
        print(f"파일명: {file_path}")
        create_folders_from_path(os.path.join("%d월 산학협력프로젝트 구글파일 다운로드" %month,os.path.dirname(file_path))) #상위 폴더 생성
        # 파일 다운로드 실행
        download_file(service, file['name'], file['id'], file['modifiedTime'], os.path.join("%d월 산학협력프로젝트 구글파일 다운로드/" %month, os.path.dirname(file_path))) #파일 다운로드
    print("총 %d개의 파일을 다운로드에 성공했습니다. %d개의 파일을 다운로드에 실패했습니다." %(downloaded_files_cnt,downloaded_failed_files_cnt)) #다운로드 성공, 실패한 파일 수 출력
    print("다운로드에 실패한 파일 목록 : ")
    for file in failed_files:
        print(file)
    input()

if __name__ == '__main__':
    main()