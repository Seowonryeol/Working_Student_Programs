import os
import re
from datetime import datetime
from docx import Document
from fpdf import FPDF
import comtypes.client
from PyPDF2 import *
import sys
import pandas as pd
import openpyxl

#월, 차수 넣으면 각 파일 pdf파일 생성
#특정 달에 해당하는 모든 word 파일을 하나로 통합한 pdf파일 생성
#각각 해서 총 2개로, 안 되면 둘 중 하나

def merge_pdfs_according_to_xlsx_file(xlsx_file_path, input_files_1,input_files_2, output_file, month): #1차 폴더와 2차 폴더 내의 pdf를 하나하나 순서를 입력해서 합쳐주는 함수
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

def get_all_files_with_extension(directory, extension):  # 디렉토리 내의 모든 파일을 검색해 리턴함
    all_files = []

    # os.walk() 함수를 사용하여 디렉토리를 순회하면서 파일들을 찾음
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(extension):
            # 파일의 절대 경로를 생성하여 리스트에 추가
               file_path = os.path.join(root, file)
               all_files.append(file_path)
    return all_files

def get_data_from_excel_column(df, column_name):
    """
    엑셀 파일에서 column_name 키워드가 포함된 셀을 찾고, 해당 셀 아래에 있는 이름 데이터를 가져오는 함수.

    Parameters:
        df (str): 엑셀 파일에서 읽어 만들어진 df
        column name (str): 데이터를 찾고자하는 열의 헤더(이름) ex) 학번, 이름

    Returns:
        list: 엑셀 파일에서 찾은 헤더 하위의 데이터들을 담은 리스트.
    """

    # column_name을 포함한 셀 찾기
    student_number_cell = df.apply(
        lambda row: row.astype(str).str.contains(column_name, case=False)
    ).any()

    # column_name이라고 적힌 셀의 위치
    if student_number_cell.any():
        name_column_index = student_number_cell.where(student_number_cell == True).first_valid_index()
    else:
        raise ValueError(f"Couldn't find a column with the name '{column_name}'.")

    name_row = df[df[name_column_index] == column_name].index[0]

    # '이름'이라고 적힌 셀의 아래쪽 셀들에 있는 데이터 가져오기
    names = df.loc[name_row + 1 :, name_column_index].tolist()
    # 문자열 형태인 값을 반환
    return list(map(str, names))


print("/n입력하는 월 1차 bi-weekly 보고서 리스트와 2차 bi-weekly 보고서 리스트 폴더 내의 모든 보고서 pdf 파일들을 각 사람별 1,2차를 붙여 생성합니다."
                " 선생님께서 주신 산학협력프로젝트 xlsx 순서에 맞게 pdf를 합쳐줍니다. 폴더 내 기존 파일의 순서와 다릅니다."
                "ex) 서원렬 5월 1차.pdf + 서원렬 5월 2차.pdf + 고선희 5월 1차.pdf 순서\n") 
month=input("원하는 달을 입력하세요 (ex)5월의 경우 : 5) :") #프로그램 설명 및 월 입력
abs_practice_folder_path = input("프로그램이 실행되어야할 절대경로를 입력하세요. 해당 폴더에 xlsx파일도 있어야합니다. ex) C:/sunny : ") #프로그램이 실행될 절대 경로 입력

xlsx_files=[]

for file in os.listdir(abs_practice_folder_path): #폴더 내에서 xlsx 엑셀 파일 탐색
    if file.endswith(".xlsx"):
        xlsx_files.append(file)

for (i, xlsx_file) in enumerate(xlsx_files, 1): #xlsx 파일명 출력 및 참조할 xlsx 파일 입력받음
    print(str(i)+". "+ xlsx_file)
xlsx_file_num=int(input("위 파일들 중 참조하고자 하는 xlsx 파일 번호를 입력하세요 : "))
xlsx_file_path=abs_practice_folder_path+"/"+xlsx_files[xlsx_file_num-1]

#파일을 병합할 때 기준이 되는 열 입력
print("파일 병합 기준이 되는 열의 이름을 입력하세요. 그 열의 데이터 순서대로 파일이 병합됩니다. ex)이름, 학번 : ",  end="")
column_name=input()

df_students = pd.read_excel(xlsx_file_path, sheet_name=0, header=None)
students_num=get_data_from_excel_column(df_students, column_name) #엑셀 파일 참조 후 df 생성

files=get_all_files_with_extension(abs_practice_folder_path,".pdf")

print(xlsx_file_path, students_num, files)

merger = PdfMerger()
cnt = 0
for student_num in students_num:
    for file in files:
        file_basename = os.path.basename(file)
        if  student_num in file_basename and str(month) in file_basename:
            merger.append(os.path.join(abs_practice_folder_path, file))
            print(file_basename + " 병합 완료")
            cnt += 1
merger.write(os.path.join(abs_practice_folder_path, f"(산학협력프로젝트 Bi-weekly 보고서 제출 현황.xlsx에 따른) {month}월 bi-weekly 보고서 통합본.pdf"))
print(f"총 {cnt}개의 파일, (산학협력프로젝트 Bi-weekly 보고서 제출 현황.xlsx에 따른) %d월 bi-weekly 보고서 통합본.pdf 생성 완료\n")
merger.close()

input("프로그램을 종료하려면 아무 키나 누르세요")
#merge_pdfs_according_to_xlsx_file(xlsx_file_path, items_1, items_2, "(산학협력프로젝트 Bi-weekly 보고서 제출 현황.xlsx에 따른) %d월 bi-weekly 보고서 통합본.pdf" %month, month) #pdf merge 실행