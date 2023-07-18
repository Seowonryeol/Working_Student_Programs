import os
import shutil


# 0이 제일 하위경로, 숫자가 커질수록 뒷경로 0,0 넣으면 root폴더 반환 1,1 그 하위 폴더, 0,2 넣으면 최상위폴더/하위폴더/하위폴더 반환
def get_sub_paths(path, start_level, end_level):
    
    all_folders = path.split('\\')
    if start_level==0 and end_level==0:
        return "./"
    elif start_level==0 and end_level==1:
        return '\\'.join(all_folders[0:1])[2:]
    elif start_level==1 and end_level==1:
        return '\\'.join(all_folders[0:1])[2:]
    return '\\'.join(all_folders[start_level-1:end_level])


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
    except FileExistsError as e:
        print("",end="")
    except Exception as e:
        print("", str(e))


# 디렉토리와 하위폴더 내 전체 파일의 경로+이름을 반환하는 함수
def get_all_files(directory):
    all_files = []
    # 디렉토리 내의 파일과 폴더 목록을 가져옴
    items = os.listdir(directory)

    for item in items:  # 디렉토리 내의 모든 파일, 폴더에 대해
        item_path = os.path.join(directory, item)
        if (os.path.isfile(item_path)):
            # 파일인 경우 파일 목록에 추가
            if "~$" not in item and "~" not in item:
                all_files.append(item_path)
        elif os.path.isdir(item_path):
            # 폴더인 경우 재귀적으로 폴더 안의 파일 목록을 가져옴
            # print(item_path+" 폴더가 성공적으로 등록되었습니다.")
            subfolder_files = get_all_files(item_path)
            all_files.extend(subfolder_files)
    return all_files


files_path=get_all_files("./") #전체 파일 스캔
deepest_folder_depth=0 #가장 하위 디렉토리에 위치한 폴더의 깊이
first_deepest_file_path="" #첫번째로 발견된 가장 하위 디렉토리에 위치한 파일의 경로
for file_path in files_path:
    if(get_folder_depth(file_path)>deepest_folder_depth and len(get_sub_paths(file_path,1,1))>=2 and len(get_sub_paths(file_path,1,1))<=6): #폴더명 길이가 사람이름의 길이인 2~6자 이내 
        deepest_folder_depth=get_folder_depth(file_path)
        first_deepest_file_path=file_path
#포함돼야할 하위 디렉토리들과 그 순서를 입력받음
print(deepest_folder_depth)
arrange_order=map(int, input(first_deepest_file_path + "의 형식으로 파일들이 존재할 때, 새로 만들고자하는 폴더에 포함돼야할 하위 디렉토리'만'의 순서를 입력해주세요.\n./"
                           +get_sub_paths(first_deepest_file_path,2,2) +"/" +
                            get_sub_paths(first_deepest_file_path,3,3) +"/"+get_sub_paths(first_deepest_file_path,1,1)+"/"
                            +"의 형태로 재구성하고자 한다면 2, 3, 1을 입력해주세요)").split(","))

arrange_order=list(arrange_order)
arrange_order_count=len(arrange_order) #포함돼야할 하위 디렉토리 종류들의 개수
bi_weekly_folder="./bi-weekly 보고서 with directories as "
for order in arrange_order: #최상위 폴더명 생성
    bi_weekly_folder+=get_sub_paths(first_deepest_file_path,order,order)+"#"

create_folder(bi_weekly_folder) #최상위 폴더 생성
count=1
while arrange_order_count-count>=0: #하위 폴더 생성
    for file_path in files_path:
        if(get_folder_depth(file_path)==deepest_folder_depth and os.path.isfile(file_path) and len(get_sub_paths(file_path,1,1))>=2 and len(get_sub_paths(file_path,1,1))<=6): #디렉토리상 가장 하위에 위치한 파일이라면
            sub_folder_name=bi_weekly_folder
            for sub_order in range(0,count): #포함돼야할 하위 디렉토리들의 개수만큼
                sub_folder_name+="/"+get_sub_paths(file_path,arrange_order[sub_order],arrange_order[sub_order]) #하위 디렉토리명 생성
                
            create_folder(sub_folder_name) #하위 폴더들 생성
    count+=1
count=0

for file_path in files_path:
    if(os.path.isfile(file_path) and os.path.splitext(file_path)[1] in ['.hwp','.doc','.docx'] and len(get_sub_paths(file_path,1,1))>=2 and len(get_sub_paths(file_path,1,1))<=6):
        sub_folder_name=bi_weekly_folder
        for sub_folder_order in arrange_order:
            sub_folder_name+="\\"+get_sub_paths(file_path,sub_folder_order,sub_folder_order)
        # 파일이 이미 존재하는지 확인
        new_file_path = sub_folder_name+"\\"+os.path.basename(file_path)
        if os.path.exists(new_file_path):
            new_file_path = sub_folder_name+"\\"+os.path.splitext(os.path.basename(file_path))[0]+"_new"+ os.path.splitext(os.path.basename(file_path))[1]
        try:
            shutil.copy(file_path, new_file_path)
            print("파일을 복사했습니다.")
        except FileExistsError as e:
            print("파일이 이미 존재합니다:", e.filename)
        except Exception as e:
            print("", str(e))


