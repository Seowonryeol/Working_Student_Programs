import os
import shutil


# 0이 제일 하위경로, 숫자가 커질수록 앞경로 0,0 넣으면 파일명 반환 1,1 최하위 폴더 반환
def get_sub_paths(path, start_level, end_level):
    """
    주어진 경로에서 start_level부터 end_level까지의 중간 경로 요소를 반환합니다.
    """
    all_folders = []
    level = 0
    while level < start_level:
        path, _ = os.path.split(path)
        level += 1

    while level <= end_level:
        path, folder = os.path.split(path)
        if folder != "":
            all_folders.append(folder)
        else:
            break
        level += 1
        all_folders.reverse()

    return os.path.join(*all_folders)


def get_folder_depth(path):
    """
    주어진 경로의 폴더 깊이를 반환합니다. 현재 경로에서 root폴더면 0, 그 이상부터 +1
    """
    depth = 0
    while True:
        path, folder = os.path.split(path)
        if folder != "":
            depth += 1
        else:
            break
    return depth-1


def create_folder(folder_path):
    try:
        os.mkdir(folder_path)
    # except FileExistsError as e:
    # print("폴더가 이미 존재합니다.")
    except Exception as e:
        print("", str(e))


# 디렉토리와 하위폴더 내 전체 파일의 경로+이름을 반환하는 함수
def get_all_files(directory, month, level):
    all_files = []
    # 디렉토리 내의 파일과 폴더 목록을 가져옴
    items = os.listdir(directory)

    for item in items:  # 디렉토리 내의 모든 파일, 폴더에 대해
        item_path = os.path.join(directory, item)
        if (os.path.isfile(item_path) and item_path.find(str(month) + "월") != -1 and item_path.find(str(level) + "차") != -1
        ):  # 조건에 맞는 파일 반환
            # 파일인 경우 파일 목록에 추가
            if "~$" not in item and "~" not in item:
                all_files.append(item_path)
        elif os.path.isdir(item_path):
            # 폴더인 경우 재귀적으로 폴더 안의 파일 목록을 가져옴
            # print(item_path+" 폴더가 성공적으로 등록되었습니다.")
            subfolder_files = get_all_files(item_path, month, level)
            all_files.extend(subfolder_files)
    return all_files


month, level = map(
    int, input("원하는 달과 차수를 콤마를 두고 입력하세요 (ex)5월 1차의 경우 : 5, 1) : ").split(",")
)
bi_weekly_folder="./bi-weekly 보고서 " + str(month) + "월 " + str(level) + "차/"
files_path = get_all_files("./", month, level)
create_folder(bi_weekly_folder)
for file_path in files_path: 
    if(get_folder_depth(file_path)==5): #교수님명, 기업명 폴더 생성
        if(len(get_sub_paths(file_path,4,4))>=3 and len(get_sub_paths(file_path,4,4))<=4):
            create_folder(bi_weekly_folder+get_sub_paths(file_path,4,4))
            create_folder(bi_weekly_folder+get_sub_paths(file_path,4,4)+'/'+get_sub_paths(file_path,2,2))
for file_path in files_path:
    if(get_folder_depth(file_path)==5):
        if(os.path.isfile(file_path) and len(get_sub_paths(file_path,4,4))>=3 and len(get_sub_paths(file_path,4,4))<=4):
            shutil.copy(file_path,os.path.join(bi_weekly_folder,get_sub_paths(file_path,4,4),get_sub_paths(file_path,2,2)))
