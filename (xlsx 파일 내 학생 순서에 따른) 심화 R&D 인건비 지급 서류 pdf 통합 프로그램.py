import os
import pandas as pd
from PyPDF2 import PdfMerger


def get_all_files(directory):  # 디렉토리 내의 모든 파일을 검색해 리턴함
    all_files = []

    # os.walk() 함수를 사용하여 디렉토리를 순회하면서 파일들을 찾음
    for root, _, files in os.walk(directory):
        for file in files:
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


def create_folder(path):
    try:
        # 폴더가 이미 존재하지 않는 경우에만 폴더를 생성합니다.
        if not os.path.exists(path):
            os.makedirs(path)
            print(f"'{path}' 디렉토리가 성공적으로 생성되었습니다.")
        else:
            print(f"'{path}' 디렉토리는 이미 존재합니다.")
    except OSError as e:
        print(f"디렉토리 '{path}' 생성 실패. 오류 내용: {e}")


# 탐색하고 병합해야할 파일의 종류들과 출력 파일에 들어가야할 이름들
file_types_undergraduate = ["개인", "신분증", "통장"]
file_types_graduate = ["개인", "신분증", "통장", "청렴서약", "건강"]
file_types_output = ["개인정보수집이용제공 동의서", "신분증, 통장 사본", "청렴서약서", "건강보험자격득실확인서"]

print(
    f"이 프로그램은 학부생/대학원생의 심화R&D 인건비 지급서류들('신분증 사본', '통장 사본' 등)을 주어진 엑셀 파일의 학생 이름 순서에 따라"
    " 각 서류별로 하나의 pdf로 합쳐 만들어주는 프로그램입니다. (학부생/대학원생 구분) ex) 신분증, 통장사본-학부생.pdf, 개인정보동의서-학부생.pdf\n"
    "학생들의 정보가 담긴 엑셀 파일명을 절대경로로 입력하세요. 확장자 포함해야합니다. 학부생 sheet가 1번, 대학원생 sheet가 2번 시트여야합니다. ex) c:/user/심화 R&D 엑셀.xlsx\n절대 경로를 포함한 파일명 : ",
    end="",
)
# 정보를 읽어들일 엑셀파일의 이름 입력
excel_file_name = input()

#파일을 병합할 때 기준이 되는 열 입력
print("파일 병합 기준이 되는 열의 이름을 입력하세요. 그 열의 데이터 순서대로 파일이 병합됩니다. ex)이름, 학번 : ",  end="")
column_name=input()

# 학부생/대학원생의 엑셀 파일 읽기
df_undergraduate = pd.read_excel(excel_file_name, sheet_name=0, header=None)
df_graduate = pd.read_excel(excel_file_name, sheet_name=1, header=None)

# '이름'이라는 글자를 포함한 열 찾기
undergraduate_std_num = get_data_from_excel_column(df_undergraduate, column_name)
graduate_std_num = get_data_from_excel_column(df_graduate, column_name)

# 프로그램이 실행될 폴더 경로 입력
print("프로그램이 실행돼야하는 폴더의 경로를 절대경로로 입력하세요 ex) c:/user/산학R&D\n폴더 절대 경로 : ", end="")

# 특정 폴더의 모든 파일 얻기
abs_practice_folder_path = input()

print(
    "파일이 저장될 폴더의 절대 경로를 입력하세요. 프로그램 실행 위치와 동일하다면 바로 엔터키를 눌러주세요.폴더가 자동 생성되지 않으니 폴더를 생성하고 경로를 입력해주세요 ex)c:/user/산학R&D/7월 \n 폴더 절대경로 :",
    end="",
)

#pdf 파일이 저장될 절대 경로 입력
pdf_created_abs_practice_folder_path = input()
if pdf_created_abs_practice_folder_path == "":
    pdf_created_abs_practice_folder_path = abs_practice_folder_path
files = get_all_files(abs_practice_folder_path)

# 학부생 개인정보 동의서에 대해 pdf 파일 생성
merger = PdfMerger()
cnt = 0
for name in undergraduate_std_num:
    for file in files:
        file_basename = os.path.basename(file)
        if name in file_basename and "개인" in file_basename:
            merger.append(os.path.join(abs_practice_folder_path, file))
            print(file_basename + " 병합 완료")
            cnt += 1
merger.write(os.path.join(pdf_created_abs_practice_folder_path, f"개인정보이용동의서_학부생.pdf"))
print(f"총 {cnt}개의 파일, 개인정보이용동의서_학부생.pdf 파일 생성완료\n")
merger.close()

# 학부생 각 통장사본, 신분증 사본 파일 유형에 대해 pdf 파일 생성
merger = PdfMerger()
cnt = 0
for name in undergraduate_std_num:
    for file_type in file_types_undergraduate[1:]:
        for file in files:
            file_basename = os.path.basename(file)
            if name in file_basename and file_type in file_basename:
                print(file_basename + " 병합 완료")
                merger.append(os.path.join(abs_practice_folder_path, file))
                cnt += 1
merger.write(os.path.join(pdf_created_abs_practice_folder_path, f"신분증, 통장 사본_학부생.pdf"))
print(f"총 {cnt}개의 파일신분증, 통장 사본 - 학부생.pdf 파일 생성완료\n")
merger.close()

# 대학원생 통장사본, 신분증 사본을 제외한 파일 유형에 대해 pdf 파일 생성
indices = [0, 3, 4]
for index in indices:
    merger = PdfMerger()
    cnt = 0
    for name in graduate_std_num:
        for file in files:
            file_type = file_types_graduate[index]
            file_basename = os.path.basename(file)
            if name in file_basename and file_type in file_basename:
                merger.append(os.path.join(abs_practice_folder_path, file))
                print(file_basename + " 병합 완료")
                cnt += 1
    output_index = index - 1 if index > 0 else 0
    merger.write(
        os.path.join(
            pdf_created_abs_practice_folder_path, f"{file_types_output[output_index]}_대학원생.pdf"
        )
    )
    print(f"총 {cnt}개의 파일, {file_types_output[output_index]}_대학원생.pdf 파일 생성완료\n")
    merger.close()

# 대학원생 통장사본, 신분증 사본 파일 유형에 대해 pdf 파일 생성
merger = PdfMerger()
cnt = 0
for name in graduate_std_num:
    for file in files:
        for file_type in file_types_graduate[1:3]:
            file_basename = os.path.basename(file)
            if name in file_basename and file_type in file_basename:
                print(file_basename + " 병합 완료")
                merger.append(os.path.join(abs_practice_folder_path, file))
                cnt += 1
merger.write(os.path.join(pdf_created_abs_practice_folder_path, f"통장 사본, 신분증 사본_대학원생.pdf"))
merger.close()
print(f"총 {cnt}개의 파일, 통장 사본, 신분증 사본_대학원생.pdf 파일 생성완료")
# 프로그램 종료
input("프로그램을 종료하려면 아무 키나 누르세요")
