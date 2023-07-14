import os
import re
from datetime import datetime
from docx import Document
from fpdf import FPDF
import comtypes.client
from PyPDF2 import *

#월, 차수 넣으면 각 파일 pdf파일 생성
#특정 달에 해당하는 모든 word 파일을 하나로 통합한 pdf파일 생성
#각각 해서 총 2개로, 안 되면 둘 중 하나

def merge_pdfs(input_files_1,input_files_2, output_file, month):
    input_files_2_copied=input_files_2.copy() #2차 폴더 내에만 존재하는 파일을 병합하기 위해 input_files_2의 복사본 생성
    merger = PdfMerger()
    for input_file_1 in input_files_1:
        merger.append("./"+str(month)+"월 1차 bi-weekly 보고서 리스트/" + input_file_1) #1차 폴더에 있는 pdf를 병합
        elements_1 = re.split(r'\s+|_|-| |\\', input_file_1) #1차 파일 이름을 요소별로 분해
        for input_file_2 in input_files_2:
            student_id_1=[word for word in elements_1 if '20' in word][1] #1차 파일의 학번 반환
            elements_2 = re.split(r'\s+|_|-| |\\', input_file_2) #2차 파일 이름을 요소별로 분해
            student_id_2=[word for word in elements_2 if '20' in word][1] #2차 파일의 학번 반환
            if (student_id_1==student_id_2): #2차 폴더 내의 pdf가 1차 폴더 내 pdf의 학번이 같다면
                input_files_2_copied.remove(input_file_2) #2차 폴더 내에만 존재하는 파일을 구분하기 위해 1차 폴더 내 파일에 해당하는 파일을 제외
                merger.append("./"+str(month)+"월 2차 bi-weekly 보고서 리스트/" + input_file_2) #2차 폴더에 있는 pdf를 병합
                break
    print(input_files_2_copied)
    for input_file_2_copied in input_files_2_copied: #2차 폴더 내에만 존재하는 파일을 병합
        merger.append("./"+str(month)+"월 2차 bi-weekly 보고서 리스트/"+input_file_2_copied)
    merger.write(output_file)
    merger.close()




def get_all_files(directory, month): #디렉토리와 하위폴더 내 전체 파일의 경로+이름을 반환하는 함수
        all_files = []
    # 디렉토리 내의 파일과 폴더 목록을 가져옴
        items = os.listdir(directory)

        for item in items:
          item_path = os.path.join(directory, item)
          if os.path.isfile(item_path) and item_path.find(str(month)+'월')!=-1: #조건에 맞는 파일 반환
            # 파일인 경우 파일 목록에 추가
             if "~$" not in item and "~" not in item:
                all_files.append(item_path)
          elif os.path.isdir(item_path):
                # 폴더인 경우 재귀적으로 폴더 안의 파일 목록을 가져옴
           subfolder_files = get_all_files(item_path,month)
           all_files.extend(subfolder_files)

        return all_files       

print("입력하는 월 1차 bi-weekly 보고서 리스트와 2차 bi-weekly 보고서 리스트 폴더 내의 모든 보고서 pdf 파일들을 각 사람별 1,2차를 붙여 생성합니다."
                "ex) 서원렬 5월 1차.pdf + 서원렬 5월 2차.pdf + 고선희 5월 1차.pdf 순서\n")
month=int(input("원하는 달을 입력하세요 (ex)5월의 경우 : 5) :"))
if(os.path.exists(str(month) + "월 bi-weekly 보고서 통합본.pdf")):
    os.remove(str(month) + "월 bi-weekly 보고서 통합본.pdf")
folder_path = './'
items_1=os.listdir("./"+str(month)+"월 1차 bi-weekly 보고서 리스트")
items_2=os.listdir("./"+str(month)+"월 2차 bi-weekly 보고서 리스트")
merge_pdfs(items_1, items_2, str(month) + "월 bi-weekly 보고서 통합본.pdf", month)
print("병합을 완료하였습니다")
input()
# for path in ｆ
#     print(path)