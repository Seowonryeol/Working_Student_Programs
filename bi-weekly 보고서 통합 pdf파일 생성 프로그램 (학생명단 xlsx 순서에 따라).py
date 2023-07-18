import os
import re
from datetime import datetime
from docx import Document
from fpdf import FPDF
import comtypes.client
from PyPDF2 import *
import sys

#월, 차수 넣으면 각 파일 pdf파일 생성
#특정 달에 해당하는 모든 word 파일을 하나로 통합한 pdf파일 생성
#각각 해서 총 2개로, 안 되면 둘 중 하나

def merge_pdfs(input_files_1,input_files_2, output_file, month): #1차 폴더와 2차 폴더 내의 pdf를 하나하나 순서를 입력해서 합쳐주는 함수
    merger = PdfMerger()
    for (i, input_file) in enumerate(input_files_1, 1):
            print(str(i)+". " +input_file) #1차 폴더 내에 있는 모든 pdf 이름을 출력
    for i in range(1, len(input_files_1)+1):
        file_select_num=int(input("번호를 입력하세요 : "))-1 #1차 폴더 내에 있는 pdf 중 병합할 pdf를 선택  
        if file_select_num==100:
            file_name=input("파일 이름을 입력하세요 : ")+".pdf"
            merger.write(file_name)
            merger.close()
            print("병합을 완료하였습니다")
            input()
            sys.exit()
        merger.append("./"+str(month)+"월 1차 bi-weekly 보고서 리스트/" + input_files_1[file_select_num]) #1차 폴더에 있는 pdf를 병합
        elements_1 = re.split(r'\s+|_|-| |\\', input_files_1[file_select_num]) #1차 파일 이름을 요소별로 분해
        print("./"+str(month)+"월 1차 bi-weekly 보고서 리스트\\"+ input_files_1[file_select_num]+"파일이 추가되었습니다.")
        for input_file_2 in input_files_2:
                student_id_1=[word for word in elements_1 if '20' in word][1] #1차 파일의 학번 반환
                elements_2 = re.split(r'\s+|_|-| |\\', input_file_2) #2차 파일 이름을 요소별로 분해
                student_id_2=[word for word in elements_2 if '20' in word][1] #2차 파일의 학번 반환
                if (student_id_1==student_id_2): #2차 폴더 내의 pdf가 1차 폴더 내 pdf의 학번이 같다면
                    merger.append("./"+str(month)+"월 2차 bi-weekly 보고서 리스트/" + input_file_2) #2차 폴더에 있는 pdf를 병합
                    print("./"+str(month)+"월 2차 bi-weekly 보고서 리스트\\"+input_file_2+" 파일이 추가되었습니다.")
                    break
        if(i%10==0):
             file_name=input("파일 이름을 입력하세요 : ")+".pdf"
             merger.write(file_name)
             merger.close()
             print("병합을 완료하였습니다")
             input()
             sys.exit()


print("입력하는 월 1차 bi-weekly 보고서 리스트와 2차 bi-weekly 보고서 리스트 폴더 내의 모든 보고서 pdf 파일들을 각 사람별 1,2차를 붙여 생성합니다."
                " 선생님께서 주신 산학협력프로젝트 xlsx 순서에 맞게 pdf를 합쳐줍니다. 기존 파일의 순서와 다릅니다."
                "ex) 서원렬 5월 1차.pdf + 서원렬 5월 2차.pdf + 고선희 5월 1차.pdf 순서\n")
month=input("원하는 달을 입력하세요 (ex)5월의 경우 : 5) :")
folder_path = './'
items_1=os.listdir("./"+month+"월 1차 bi-weekly 보고서 리스트")
items_2=os.listdir("./"+month+"월 2차 bi-weekly 보고서 리스트")
merge_pdfs(items_1, items_2, str(month) + "월 bi-weekly 보고서 통합본.pdf", month)