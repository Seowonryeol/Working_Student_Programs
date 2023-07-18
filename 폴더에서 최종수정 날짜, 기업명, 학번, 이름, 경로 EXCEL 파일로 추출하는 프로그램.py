import os
import re
import datetime
from datetime import datetime
import pandas as pd


def get_all_files(directory): #디렉토리와 하위폴더 내 전체 파일의 경로+이름을 반환하는 함수
  all_files = []

    # 디렉토리 내의 파일과 폴더 목록을 가져옴
  items = os.listdir(directory)

  for item in items:
    item_path = os.path.join(directory, item)
    if os.path.isfile(item_path):
            # 파일인 경우 파일 목록에 추가
      all_files.append(item_path)
    elif os.path.isdir(item_path):
                # 폴더인 경우 재귀적으로 폴더 안의 파일 목록을 가져옴
      subfolder_files = get_all_files(item_path)
      all_files.extend(subfolder_files)

  return all_files       

def get_file_modified_date(filename): #파일의 최종 수정날짜를 반환하는 함수
    modified_time = os.path.getmtime(filename)
    modified_date = datetime.fromtimestamp(modified_time).date()
    return modified_date
folders=[]
#입력 가이드 출력 및폴더명 입력
print("excel 파일로 정보를 추출하고 싶은 폴더명을 입력하세요. 정확하게 입력하셔야합니다. (ex) 2021년 5월 1차의 경우 : 2021년 5월 1차)\n" 
      "폴더 명이 너무 길 경우 다른 폴더와 구분되는 폴더명의 일부만 입력하셔도 됩니다. (ex) 2021년 5월 1차의 경우 : 2021년 5월)") 
folder_name=input("폴더명을 입력하세요 : ")

#정보를 추출하고자 하는 폴더를 찾는 과정
while(len(folders)!=1): #해당하는 폴더가 1개일 때까지 반복
  folders=[f for f in os.listdir('./') if os.path.isdir(f) and f.find(folder_name)!=-1]
  #print(folders)
  if(len(folders)==0): #폴더가 존재하지 않을 경우
    print("해당 폴더가 존재하지 않습니다. 다시 입력해주세요.\n폴더명 :", end="")
    folder_name=input()
    continue
  elif len(folders)>=2: #폴더가 2개 이상일 경우
    for i, folder_name in enumerate(folders, start=1):
      print(str(i) + ". " + folder_name)
    print("중 하나를 선택하여 번호를 입력하세요 : ",end="")
    folder_num=int(input())
    folder_name=folders[folder_num-1]
  else: #폴더가 하나일 경우
    folder_name=folders[0]
    print(folder_name+" 폴더에서 정보를 추출합니다.")
    continue

#확장자를 입력받음(* 입력 시 모든 확장자)
extensions = [ext.strip() for ext in input("해당 폴더 내에서 추출하고자 하는 파일의 확장자를 콤마를 두고 입력하세요. 모든 확장자를 원할 경우 공백을 입력하세요\n"
                                           "(ex)docx, pdf의 경우 : docx, pdf) : ").split(',')]  
#해당 폴더 내 파일들을 찾아서 파일명을 요소별로 분해하고, 요소들을 통해 정보를 추출하는 과정
files = get_all_files("./"+folder_name)
file_paths=[]
file_infs=[]
for file_path in files:
    if os.path.isfile(file_path) and (os.path.splitext(file_path)[1][1:] in extensions or extensions[0]==''):  # 폴더가 아닌 파일만 처리, 선택한 확장자의 파일만 처리
        elements = re.split(r'\s+|_|-| |\\', os.path.basename(file_path)) #파일 이름을 요소별로 분해
        index_2023=[i for i, word in enumerate(elements) if '2023' in word][0]
        student_id=[word for word in elements if '20' in word]
        len_ele=len(elements)
        modified_date = get_file_modified_date(file_path)
        file_infs.append([elements[index_2023+4],student_id[1],modified_date,elements[index_2023+6],file_path])
        file_paths.append(file_path)
df=pd.DataFrame(file_infs,columns=["기업명","학번","최종수정날짜", "이름", "전체경로"])
df.to_excel("./"+folder_name+ " 폴더 최종수정 날짜, 기업명, 학번, 이름, 경로 추출 EXCEL 파일.xlsx", index=False)
# for path in file_paths:
#     print(path)