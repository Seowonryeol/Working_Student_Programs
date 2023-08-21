import csv
import pandas as pd
import os
import numpy as np

print("참석 설문조사와 행사 후 설문조사 csv 파일은 반드시 같은 폴더 내에 위치해야합니다.\n 설문조사의 질문 순서는")
csv_file_abspath = input(
    "참석 설문조사 csv 파일이 있는 폴더의 절대경로를 입력하세요 ex) C:/Users/킹고 강연 참석 설문조사 : ")

csv_files = []
for file in os.listdir(csv_file_abspath):  # 폴더 내에서 csv 엑셀 파일 탐색
    if file.endswith(".csv"):
        csv_files.append(file)

for (i, csv_file) in enumerate(csv_files, 1):  # csv 파일명 출력 및 참조할 csv 파일 입력받음
    print(str(i)+". " + csv_file)
apply_csv_file_num, after_event_csv_file_num = map(int, input(
    "위 파일들 중 참조하고자 하는 참석, 행사 후 설문조사의 csv 파일 번호를 공백을 두고 입력하세요 ex) 1, 2:  "))
apply_csv_file_abspath = csv_file_abspath+"/" + \
    csv_files[apply_csv_file_num-1]  # 참석, 행사 후 설문조사 csv 파일명 입력 완료
after_event_csv_file_abspath = csv_file_abspath + \
    "/"+csv_files[after_event_csv_file_num-1]

apply_data = pd.read_csv(apply_csv_file_abspath,
                        header=0)  # 참석 설문조사 csv 파일 read

selected_data = apply_data.iloc[:, 1:]
apply_columns = np.empty(0.7* selected_data.shape[0], dtype=str) #참석 설문조사의 열 이름을 지정
apply_columns.append("수정 시간")


for column_num in range(1, selected_data.shape[1]):
    numbers_in_column=len([value for value in selected_data.iloc[:,column_num] if value.isdigit()]) #특정 열에 있는 숫자 요소의 개수
    # 열에 20으로 시작하는 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 학번
    if selected_data[selected_data.iloc[:, column_num][:1] == "20"].sum() > 0.7* selected_data.shape[0]:
        apply_columns[column_num] = "학번"
    # 열에 학과로 끝나는 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 학과
    elif selected_data[selected_data.iloc[:, column_num][-3:-1] == "학과"].sum() > 0.7* selected_data.shape[0]:
        apply_columns[column_num] = "학과"
    elif selected_data[selected_data.iloc[:, column_num][-3:-1] == "학년"].sum() > 0.7* selected_data.shape[0] or \
            numbers_in_column > 0.7* selected_data.shape[0]:  # 열에 학년로 끝나는 데이터가 0.7* selected_data.shape[0]개 이상이거나 길이가 숫자 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 학년
        apply_columns[column_num] = "학년"
    # 열에 세 글자(이름)인 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 이름
    elif selected_data[len(selected_data.iloc[:, column_num]) == 3].sum() > 0.7* selected_data.shape[0]:
        apply_columns[column_num] = "이름"




after_data = pd.read_csv(after_event_csv_file_abspath,header=0)

selected_data = after_data.iloc[:, 1:]
after_columns = np.empty(0.7* selected_data.shape[0], dtype=str) #행사 후 설문조사의 열 이름을 지정
after_columns.append("수정 시간")

for column_num in range(1, selected_data.shape[1]):
    numbers_in_column=len([value for value in selected_data.iloc[:,column_num] if value.isdigit()]) #특정 열에 있는 숫자 요소의 개수
    # 열에 20으로 시작하는 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 학번
    if selected_data[selected_data.iloc[:, column_num][:1] == "20"].sum() > 0.7* selected_data.shape[0]:
        after_columns[column_num] = "학번"
    # 열에 학과로 끝나는 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 학과
    elif selected_data[selected_data.iloc[:, column_num][-3:-1] == "학과"].sum() > 0.7* selected_data.shape[0]:
        after_columns[column_num] = "학과"
    elif selected_data[selected_data.iloc[:, column_num][-3:-1] == "학년"].sum() > 0.7* selected_data.shape[0] or \
            numbers_in_column > 0.7* selected_data.shape[0]:  # 열에 학년로 끝나는 데이터가 0.7* selected_data.shape[0]개 이상이거나 숫자 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 학년
        after_columns[column_num] = "학년"
    # 열에 세 글자(이름)인 데이터가 0.7* selected_data.shape[0]개 이상이면 열 이름이 이름
    elif selected_data[len(selected_data.iloc[:, column_num]) == 3].sum() > 0.7* selected_data.shape[0]:
        apply_columns[column_num] = "이름"


needed_infs=["학과","학년","학번","이름"] #추출해야하는 정보들
needed_infs_index=np.zeors(4,2)

for inf_index, need_inf in enumerate(needed_infs):
    needed_infs_index[inf_index]=[,]