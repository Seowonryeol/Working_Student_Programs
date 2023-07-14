import os
import io
from datetime import datetime
from docx import Document
from fpdf import FPDF
import comtypes.client
from PyPDF2 import *

#월, 차수 넣으면 각 파일 pdf파일 생성
#모든 word 파일을 하나로 통합한 pdf파일 생성
#각각 해서 총 2개로, 안 되면 둘 중 하나
def create_folder(folder_path):
    try:
        os.mkdir(folder_path)
        print("폴더가 성공적으로 생성되었습니다.")
    except FileExistsError:
        print("이미 폴더가 존재합니다.")
    except Exception as e:
        print("폴더 생성 중 오류가 발생했습니다:", str(e))

def convert_word_to_pdf(input_file_path, output_file_path):
    abs_input_file_path = os.path.abspath(input_file_path)
    abs_output_file_path = os.path.abspath(output_file_path)

    # Check if output_file_path already exists
    if os.path.exists(abs_output_file_path):
        # Generate a unique file name
        output_directory = os.path.dirname(abs_output_file_path)
        output_filename = os.path.basename(abs_output_file_path)
        output_filename, output_extension = os.path.splitext(output_filename)
        counter = 1
        while True:
            unique_filename = f"{output_filename}_{counter}{output_extension}"
            unique_file_path = os.path.join(output_directory, unique_filename)
            if not os.path.exists(unique_file_path):
                abs_output_file_path = unique_file_path
                break
            counter += 1

    # Perform the conversion
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(abs_input_file_path)
    doc.SaveAs(abs_output_file_path, FileFormat=17)  # 17 represents the PDF format in Word
    doc.Close()
    word.Quit()


def get_all_files_with_month_level(directory, month, level): #디렉토리와 하위폴더 내 달과 차수 조건에 맞는 파일들을 반환하는 함수
        all_files = []

    # 디렉토리 내의 파일과 폴더 목록을 가져옴
        items = os.listdir(directory)

        for item in items:
          item_path = os.path.join(directory, item)
          if os.path.isfile(item_path) and item_path.find(str(month)+'월')!=-1 and item_path.find(str(level)+'차')!=-1 and len(item_path.split("\\")[0]) in range(5,7): #월과 차수를 포함하고, 첫번째 하위 폴더명의 길이가 3~4인 파일
            # 파일인 경우 파일 목록에 

            if "~$" not in item and "~" not in item:
                all_files.append(item_path)
          elif os.path.isdir(item_path):
                # 폴더인 경우 재귀적으로 폴더 안의 파일 목록을 가져옴
           subfolder_files = get_all_files_with_month_level(item_path,month, level)
           all_files.extend(subfolder_files)

        return all_files       


month, level =map(int,input("원하는 달과 차수를 콤마를 두고 입력하세요 (ex)5월 1차의 경우 : 5, 1) : ").split(','))
folder_path = './'
create_folder(folder_path+str(month)+"월 "+str(level)+"차 bi-weekly 보고서 리스트")
files_path = get_all_files_with_month_level(folder_path, month, level)
for input_file_path in files_path:
    if os.path.isfile(input_file_path) and os.path.splitext(input_file_path)[1] in ['.doc','.docx'] :  # 워드 파일만 변환
        convert_word_to_pdf(input_file_path, "./"+str(month)+"월 "+str(level)+"차 bi-weekly 보고서 리스트/"+os.path.splitext(os.path.basename(input_file_path))[0])
        print(input_file_path+"파일이 추가되었습니다.")
input()