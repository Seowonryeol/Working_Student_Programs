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
        print("./"+str(month)+"월 1차 bi-weekly 보고서 리스트\\"+ input_file_1+"파일이 추가되었습니다.")
        for input_file_2 in input_files_2:
            student_id_1=[word for word in elements_1 if '20' in word][1] #1차 파일의 학번 반환
            elements_2 = re.split(r'\s+|_|-| |\\', input_file_2) #2차 파일 이름을 요소별로 분해
            student_id_2=[word for word in elements_2 if '20' in word][1] #2차 파일의 학번 반환
            if (student_id_1==student_id_2): #2차 폴더 내의 pdf가 1차 폴더 내 pdf의 학번이 같다면
                input_files_2_copied.remove(input_file_2) #2차 폴더 내에만 존재하는 파일을 구분하기 위해 1차 폴더 내 파일에 해당하는 파일을 제외
                merger.append("./"+str(month)+"월 2차 bi-weekly 보고서 리스트/" + input_file_2) #2차 폴더에 있는 pdf를 병합
                print("./"+str(month)+"월 2차 bi-weekly 보고서 리스트\\"+input_file_2+" 파일이 추가되었습니다.")
                break
    print(input_files_2_copied)
    for input_file_2_copied in input_files_2_copied: #2차 폴더 내에만 존재하는 파일을 병합
        merger.append("./"+str(month)+"월 2차 bi-weekly 보고서 리스트/"+input_file_2_copied)
        print("./"+str(month)+"월 2차 bi-weekly 보고서 리스트\\"+input_file_2_copied+" 파일이 추가되었습니다.")
    merger.write(output_file)
    merger.close()


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