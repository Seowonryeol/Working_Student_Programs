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

def merge_pdfs(input_files, output_file):
    merger = PdfMerger()
    print(input_files)
    for input_file in input_files:
        print(input_file+" 파일이 추가되었습니다")
        merger.append(input_file)

    merger.write(output_file)
    merger.close()


def get_all_files(directory, month): #디렉토리와 하위폴더 내 전체 파일의 경로+이름을 반환하는 함수
        all_files = []

    # 월이 맞고, 첫번째 하위 폴더의 이름 길이가 3~4인 조건에 맞는 pdf들을 반환하는 함수
        items = os.listdir(directory)

        for item in items:
          item_path = os.path.join(directory, item)
          if os.path.isfile(item_path) and item_path.find(str(month)+'월')!=-1 and os.path.splitext(item_path)[1]=='.pdf' and len(item_path.split("\\")[0]) in range(5,7): #월이 맞고, 첫번째 하위 폴더의 이름 길이가 3~4이고 pdf인 조건에 맞는 파일 반환
            # 파일인 경우 파일 목록에 추가
             if "~$" not in item and "~" not in item:
                all_files.append(item_path)
                print(item_path+" 파일이 추가되었습니다")
          elif os.path.isdir(item_path):
                # 폴더인 경우 재귀적으로 폴더 안의 파일 목록을 가져옴
           subfolder_files = get_all_files(item_path,month)
           all_files.extend(subfolder_files)

        return all_files       


month=int(input("원하는 달을 입력하세요 (ex)5월의 경우 : 5) :"))
folder_path = './'
files_path = get_all_files(folder_path, month)

merge_pdfs(files_path, str(month) + "월 bi-weekly 보고서 통합본.pdf")
input()
