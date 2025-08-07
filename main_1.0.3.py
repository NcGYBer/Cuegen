import sys
import pandas as pd
import os
from mutagen.mp3 import MP3 as mp3
from mutagen.id3 import ID3NoHeaderError as id3

def main():

    print("큐시트생성기 1.0.3 for Windows\n")

    current_dir = make_dir()
    all_metadata = []

    list_excel_path = os.path.join(current_dir, '_list.xlsx')
    empty_excel = pd.DataFrame()
    empty_excel.to_excel(list_excel_path, index=False)

    print(f"[Info] 바탕화면에 새로운 작업 폴더 '{os.path.basename(current_dir)}'와 빈 '_list.xlsx' 파일이 생성되었습니다.")
    print(f"[Info] '{current_dir}' 폴더가 파일 탐색기로 곧 열립니다. 잠시 기다려 주세요.")

    try:
        os.startfile(current_dir)
    except AttributeError:
        print("[Warning] os.startfile은 Windows에서만 지원됩니다. 해당 폴더를 수동으로 열어주세요.")
        print(f"폴더 경로: {current_dir}")
    except Exception as e:
        print(f"[Error] 폴더 열기 실패: {e}")

    print("[Info] 열린 폴더에서 '_list.xlsx' 파일을 열어 오디오 파일명(경로와 확장자 포함)을 A열에 작성한 후 저장하세요.")
    print("[Info] 그리고 메타데이터를 추출할 음원 파일들을 **이 폴더 안으로 모두 옮겨주세요.**")
    input("\n파일 준비가 완료되면 Enter 키를 누르세요...\n")

    song_file_list = get_song_list(current_dir) 
    
    if not song_file_list:
        print("[Error] '_list.xlsx' 파일에서 MP3 파일 목록을 찾을 수 없습니다. 프로그램을 종료합니다.")
        input("Enter 키를 누르세요...\n")
        sys.exit(0)

    c = 0
    for path in song_file_list:
        song_file_name = os.path.basename(path)
        data = get_mp3_metadata(current_dir, song_file_name)
        data['file_name'] = song_file_name
        all_metadata.append(data)
        c += 1
    
    df = pd.DataFrame(all_metadata)
    df.to_excel(os.path.join(current_dir, '_song_data.xlsx'))

    print(f"[Info] Success! {c}곡의 메타데이터 추출이 완료되어 '_song_data.xlsx' 파일로 저장되었습니다.")
    input("\n프로그램을 종료하려면 Enter 키를 누르세요...\n")


def get_desktop_path_env_var():
    try:
        desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        
        if os.path.isdir(desktop_path):
            return desktop_path
        else:
            print(f"[Error] 예상된 데스크탑 경로 '{desktop_path}'가 존재하지 않습니다. 프로그램을 종료합니다.")
            input("Enter 키를 누르세요...\n")
            sys.exit(0)

    except KeyError:
        print("[Error] 'USERPROFILE' 환경 변수를 찾을 수 없습니다. Windows 환경에서 프로그램을 실행해 주세요.")
        input("\n프로그램을 종료하려면 Enter 키를 누르세요...\n")
        sys.exit(0)

    except Exception as e:
        print(f"[Error] 데스크탑 경로를 가져오는 중 오류 발생: {e}")
        input("\n프로그램을 종료하려면 Enter 키를 누르세요...\n")
        sys.exit(0)


def make_dir():
    desktop_base_path = get_desktop_path_env_var()
    if desktop_base_path is None:
        return None
    
    folder_base_name = "cue"
    new_dir_path = os.path.join(desktop_base_path, folder_base_name)
    counter = 0

    while os.path.exists(new_dir_path):
        counter += 1
        new_dir_path = os.path.join(desktop_base_path, f"{folder_base_name}_{counter}")
    
    try:
        os.makedirs(new_dir_path)
        return new_dir_path
    except OSError as e:
        print(f"[Error] 디렉토리 생성 실패 - '{new_dir_path}': {e}")
        input("\n프로그램을 종료하려면 Enter 키를 누르세요...\n")
        sys.exit(0)


def get_song_list(dir):
    file_path = os.path.join(dir, '_list.xlsx')

    try:
        df = pd.read_excel(file_path, header=None)

        if not df.empty and 0 in df.columns:
            return [str(item).strip() for item in df[0].tolist() if pd.notna(item) and str(item).strip() != '']
        else:
            print(f"[Warning] '{file_path}' 파일이 비어 있거나 첫 번째 컬럼(A열)이 없습니다.")
            return []

    except FileNotFoundError:
        print(f"[Error]: 파일 '{file_path}'를 찾을 수 없습니다.")
        return []

    except Exception as e:
        print(f"[Error] 엑셀 파일을 불러오는 중 오류가 발생했습니다: {e}")
        return []

def decode_safely(text, file_name, tag_name):
    if not text:
        return ""

    try:
        byte_data = text.encode('latin-1')
        
        encodings_to_try = ['utf-8', 'cp949', 'euc-kr']
        for encoding in encodings_to_try:
            try:
                return byte_data.decode(encoding, errors='strict').strip()
            
            except UnicodeDecodeError:
                continue
    
    except UnicodeEncodeError:
        print(f"[Warning] '{file_name}'의 {tag_name} 태그가 심하게 손상되어 복구 불가. 원본 값 사용.")
        return str(text).strip()

    print(f"[Warning] '{file_name}'의 {tag_name} 태그 디코딩 실패. 원본 값 사용.")
    return str(text).strip()


def get_mp3_metadata(dir, file_name):
    metadata = {'title': "", 'Artist': "", 'Album': "", 'Publisher': "", 'Composer': "", 'Lyricist': ""}
    file_path = os.path.join(dir, file_name)
    root, ext = os.path.splitext(file_name)

    if ext.lower() == '.mp3':  
        try:
            audio = mp3(file_path)
            
            if 'TIT2' in audio.tags:
                title = decode_safely(audio.tags['TIT2'].text[0], file_name, 'title')
                metadata['title'] = title if title else root
            else:
                metadata['title'] = root

            if 'TPE1' in audio.tags:
                metadata['Artist'] = decode_safely(audio.tags['TPE1'].text[0], file_name, 'Artist')
            
            if 'TALB' in audio.tags:
                metadata['Album'] = decode_safely(audio.tags['TALB'].text[0], file_name, 'Album')

            if 'TPUB' in audio.tags:
                metadata['Publisher'] = decode_safely(audio.tags['TPUB'].text[0], file_name, 'Publisher')

            if 'TCOM' in audio.tags:
                metadata['Composer'] = decode_safely(audio.tags['TCOM'].text[0], file_name, 'Composer')

            if 'TEXT' in audio.tags:
                metadata['Lyricist'] = decode_safely(audio.tags['TEXT'].text[0], file_name, 'Lyricist')

            print(f'[Info] {file_name} 메타데이터 추출 성공!')
            return metadata
        
        except FileNotFoundError:
            metadata['title'] = "file not found"
            metadata['Artist'] = "file not found"
            metadata['Album'] = "file not found"
            metadata['Publisher'] = "file not found"
            metadata['Composer'] = "file not found"
            metadata['Lyricist'] = "file not found"
            print(f"[Warning] '{file_name}' 파일을 찾을 수 없습니다. (경로: {file_path})")
            return metadata
        
        except id3:
            metadata['title'] = root
            metadata['Artist'] = "-"
            metadata['Album'] = "-"
            metadata['Publisher'] = "-"
            metadata['Composer'] = "-"
            metadata['Lyricist'] = "-"
            print(f"[Warning] '{file_name}' 파일에 메타데이터가 없습니다. 'title'은 파일명으로 대체합니다.")
            return metadata
        
        except Exception as e:
            metadata['title'] = root
            metadata['Artist'] = "error"
            metadata['Album'] = "error"
            metadata['Publisher'] = "error"
            metadata['Composer'] = "error"
            metadata['Lyricist'] = "error"
            print(f"[Error] '{file_name}' 메타데이터 추출 중 예외 발생: {e}. 'title'은 파일명으로 대체합니다.")
            return metadata

    else:
        metadata['title'] = root
        metadata['Artist'] = "-"
        metadata['Album'] = "-"
        metadata['Publisher'] = "-"
        metadata['Composer'] = "-"
        metadata['Lyricist'] = "-"
        print(f"[Warning] '{file_name}' 파일은 MP3 파일이 아닙니다. 'title'은 파일명으로 대체합니다.")
        return metadata


if __name__ == "__main__":
    main()