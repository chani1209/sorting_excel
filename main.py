# 엑셀파일 읽어낼 pandas 라이브러리 설치 필요 -> 구동하려면 openpyxl, xlrd 라이브러리 설치 필요
import pandas as pd
import os

# 제외할 확장자를 리스트로 변환
def make_except_extension_list(except_extension):
    except_extension_list = except_extension.split(' ')
    except_extension_list = [x.strip() for x in except_extension_list]
    return except_extension_list

# 파일 확장자가 제외할 확장자인지 확인
def is_file_extension_allowed(file_path, allowed_extensions):
    _, file_extension = os.path.splitext(file_path)
    return file_extension.lower() not in allowed_extensions

# 파일명에 제외할 단어가 있는지 확인
def is_file_name_allowed(file_path, allowed_file_name):
    file_name = os.path.basename(file_path)
    for name in allowed_file_name:
        if name in file_name:
            return True
    return False

 
# 파일 삭제
def delete_files_with_name_in_directory(directory, file_names, except_extension_list):
    files_to_delete = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if is_file_name_allowed(file_path, file_names) and is_file_extension_allowed(file_path, except_extension_list):
                files_to_delete.append(file_path)

    if not files_to_delete:
        print("삭제할 파일이 없습니다.")
        return

    print("삭제 대상 파일 리스트:")
    for file_path in files_to_delete:
        print(file_path)

    confirm = input("위의 파일들을 삭제하시겠습니까? (y/n): ")
    if confirm.lower() == 'y':
        for file_path in files_to_delete:
            os.remove(file_path)
            print(f"{file_path}을(를) 삭제했습니다.")
    else:
        print("파일 삭제를 취소했습니다.")


# main 함수
def main():
    file_name = input("목록이 있는 파일명을 입력하세요: ")
    if not file_name.endswith('.xlsx'):
        file_name += '.xlsx'

    if not os.path.exists(file_name):
        print("파일이 존재하지 않습니다.")
        return
    
    directory_path = input("경로를 입력하세요: /Users/LGD")
    if not os.path.exists(directory_path):
        print("경로가 존재하지 않습니다.")
        return

    input_extension_str = input("제외할 확장자를 입력하세요: .png .xlsx .txt 형식으로 소문자로 입력하세요. (공백시 전체) : ")

    except_extension_list = make_except_extension_list(input_extension_str)
    delete_keyword = []

    df = pd.read_excel(file_name, header=None)
    # 엑셀파일의 첫번째 열을 리스트로 변환
    delete_keyword = df.iloc[:,0].tolist()
    delete_files_with_name_in_directory(directory_path, delete_keyword, except_extension_list)


if __name__ == '__main__':
    main()