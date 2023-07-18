import os
import re
import datetime
from datetime import datetime
import pandas as pd


def get_all_files(directory, month, level): #디렉토리와 하위폴더 내 전체 파일의 경로+이름을 반환하는 함수
        all_files = []

    # 디렉토리 내의 파일과 폴더 목록을 가져옴
        items = os.listdir(directory)

        for item in items:
          item_path = os.path.join(directory, item)
          if os.path.isfile(item_path) and item_path.find(str(month)+'월')!=-1 and item_path.find(str(level)+'차')!=-1:
            # 파일인 경우 파일 목록에 추가
             all_files.append(item_path)
          elif os.path.isdir(item_path):
                # 폴더인 경우 재귀적으로 폴더 안의 파일 목록을 가져옴
           subfolder_files = get_all_files(item_path,month, level)
           all_files.extend(subfolder_files)

        return all_files       

def get_file_modified_date(filename): #파일의 최종 수정날짜를 반환하는 함수
    modified_time = os.path.getmtime(filename)
    modified_date = datetime.fromtimestamp(modified_time).date()
    return modified_date

month, level =map(int,input("원하는 달과 차수를 콤마를 두고 입력하세요 (ex)5월 1차의 경우 : 5, 1) : ").split(','))
folder_path = './'
file_paths=[]
files = get_all_files(folder_path, month, level)
file_infs=[]
for file_path in files:
    #print(file_path) 파일 위치 출력
    if os.path.isfile(file_path) and os.path.splitext(file_path)[1] in ['.hwp','.doc','.docx', '.txt'] :  # 폴더가 아닌 파일만 처리
      #  print(os.path.splitext(file_path)[1]) 확장자 출력
        elements = re.split(r'\s+|_|-| |\\', file_path) #파일 이름을 요소별로 분해
        index_2023=[i for i, word in enumerate(elements) if '2023' in word][0]
        student_id=[word for word in elements if '20' in word]
        len_ele=len(elements)
        modified_date = get_file_modified_date(file_path)
        file_infs.append([modified_date,elements[index_2023+4],student_id[1],elements[index_2023+6],file_path])
       # print(f'{modified_date} {elements[index_2023+4]} {student_id[1]} {elements[index_2023+6]}')
        file_paths.append(file_path)
#print("\n경로와 전체 파일명을 포함한 파일 내역입니다")
df=pd.DataFrame(file_infs,columns=["최종수정 날짜","기업명","학번","이름", "전체경로"])
df.to_excel("./"+str(month)+"월 "+str(level)+"차 산학보고서 자료.xlsx", index=False)
# for path in file_paths:
#     print(path)