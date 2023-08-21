from importlib.metadata import files
import os
import io
from datetime import datetime
from docx import Document
from fpdf import FPDF
import comtypes.client
from PyPDF2 import *

failed_files=[]
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

def convert_word_to_pdf(abs_file_path, abs_output_file):

    # Check if output_file already exists
    if os.path.exists(abs_output_file):
        print(abs_file_path +" 파일은 이미 존재합니다")
        return
    
    try:
        # Perform the conversion
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(abs_file_path)
        doc.SaveAs(abs_output_file, FileFormat=17)  # 17 represents the PDF format in Word
        doc.Close()
        word.Quit()
        print(abs_output_file + "이 추가되었습니다.")
    except Exception as e:
        print(f"PDF 변환 중 오류가 발생했습니다: {str(e)}")
        failed_files.append(abs_output_file)
def get_all_files(directory,month, level):
    all_files = []

    for root, _, files in os.walk(directory):
        for file in files:
            abs_file_path = os.path.abspath(os.path.join(root, file))
            if abs_file_path.find("%d월" % month) != -1 and abs_file_path.find("%d차" % level) != -1 and abs_file_path.find("보고 파일") == -1 and os.path.splitext(abs_file_path)[1] in ['.doc','.docx']:
                all_files.append(abs_file_path)

    return all_files

abs_practice_path=input("프로그램을 실행하고자 하는 폴더의 절대경로를 입력하세요 ex)C:\sunny : ")
month, level =map(int,input("원하는 달과 차수를 콤마를 두고 입력하세요 (ex)5월 1차의 경우 : 5, 1) : ").split(','))
create_folder(abs_practice_path+"/"+str(month)+"월 "+str(level)+"차 bi-weekly 보고서 리스트")
files = get_all_files(abs_practice_path,month, level)
print(f"총 {len(files)}개의 파일이 검색되었습니다.")

for file in files: #전체 파일명단 출력
    print(file)

cnt=0
for file in files:
        filename_without_extension = os.path.splitext(os.path.basename(file))[0]
        output_file = os.path.join(abs_practice_path, str(month) + "월 " + str(level) + "차 bi-weekly 보고서 리스트", filename_without_extension + ".pdf")
        convert_word_to_pdf(file, output_file)
        cnt+=1
        print('{0}/{1}, {:.1f}% 완료'.format(cnt, len(files), cnt/len(files)*100))
for failed_file in failed_files:
    print(failed_file + "변환에 실패했습니다")
input("변환이 끝났습니다. 계속하려면 아무 키나 입력하세요")