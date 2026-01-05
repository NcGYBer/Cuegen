import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime
import os
from mutagen.mp3 import MP3 as mp3

class IntegratedStudioApp:
    # 앱 초기화 및 기본 변수 설정
    def __init__(self, root):
        self.root = root
        self.root.title("방정리 및 큐시트 생성 통합 도구 (2.0.1v)")
        self.root.geometry("900x850")
        
        self.df = None
        self.sessions = [] 
        self.project_dir = "" 
        
        self.create_widgets()
        self.prepare_workspace()

    # GUI 위젯(버튼, 테이블, 로그창 등) 생성 및 배치
    def create_widgets(self):
        # 상단: 파일 선택 영역 프레임
        file_frame = ttk.Frame(self.root)
        file_frame.pack(fill='x', padx=10, pady=5)
        
        # 엑셀 파일 로드 버튼
        ttk.Button(file_frame, text="엑셀 파일 선택", command=self.load_file).pack(side='left')
        # 선택된 파일명을 표시하는 레이블
        self.file_label = ttk.Label(file_frame, text="파일이 선택되지 않았습니다")
        self.file_label.pack(side='left', padx=10)
        
        # 중간: 데이터 분석 결과 미리보기 영역 프레임
        table_frame = ttk.LabelFrame(self.root, text="병합 및 회차 분리 결과 미리보기 (확인용)")
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 데이터 표시용 트리뷰(테이블) 위젯 설정
        columns = ('Session', 'Start', 'End', 'Length', 'Active_Take_Name', 'File_Path')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        # 테이블용 가로/세로 스크롤바
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)

        # 테이블 각 컬럼의 헤더 제목 설정
        self.tree.heading('Session', text='회차')
        self.tree.heading('Start', text='시작 시간')
        self.tree.heading('End', text='종료 시간')
        self.tree.heading('Length', text='길이')
        self.tree.heading('Active_Take_Name', text='트랙명')
        self.tree.heading('File_Path', text='파일 경로')

        # 테이블 각 컬럼의 너비 및 정렬 설정
        self.tree.column('Session', width=60, anchor='center', stretch=False)
        self.tree.column('Start', width=100, anchor='center', stretch=False)
        self.tree.column('End', width=100, anchor='center', stretch=False)
        self.tree.column('Length', width=100, anchor='center', stretch=False)
        self.tree.column('Active_Take_Name', width=200, anchor='w', stretch=False)
        self.tree.column('File_Path', width=450, anchor='w', stretch=False)
        
        # 스크롤바 및 테이블 배치
        scrollbar_x.pack(side='bottom', fill='x')
        scrollbar_y.pack(side='right', fill='y')
        self.tree.pack(side='left', fill='both', expand=True)
        
        # 하단: 시스템 로그 출력 영역 프레임
        log_frame = ttk.LabelFrame(self.root, text="처리 로그 및 안내")
        log_frame.pack(fill='x', expand=False, padx=10, pady=5)
        
        # 스크롤 가능한 로그 출력 텍스트 영역
        self.log_widget = scrolledtext.ScrolledText(log_frame, height=15, state='disabled', bg="#f8f9fa", font=("맑은 고딕", 9))
        self.log_widget.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 하단: 진행률 표시 영역 프레임
        progress_frame = ttk.Frame(self.root)
        progress_frame.pack(fill='x', padx=15, pady=(5, 0))
        
        # 작업 진행도 표시 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', side='top')
        
        # 현재 진행 상태(%) 표시 레이블
        self.progress_label = ttk.Label(progress_frame, text="대기 중...")
        self.progress_label.pack(side='top', pady=2)

        # 최하단: 실행 버튼 영역 프레임
        process_frame = ttk.Frame(self.root)
        process_frame.pack(fill='x', padx=10, pady=10)
        
        # 추출 및 엑셀 생성 시작 버튼
        self.extract_btn = ttk.Button(process_frame, text="모든 회차 데이터 추출 및 회차별 엑셀 생성 시작", 
                                     command=self.start_extraction, state='disabled')
        self.extract_btn.pack(fill='x', ipady=10)

    # 바탕화면에 오늘 날짜의 작업 폴더 생성 및 안내 출력
    def prepare_workspace(self):
        try:
            desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
            today = datetime.now().strftime("%Y%m%d")
            folder_base_name = f"cue_작업_{today}"
            self.project_dir = os.path.join(desktop_path, folder_base_name)
            counter = 0
            while os.path.exists(self.project_dir):
                counter += 1
                self.project_dir = os.path.join(desktop_path, f"{folder_base_name}_{counter}")
            os.makedirs(self.project_dir)
            
            self.log("="*70)
            self.log(f" [시스템] 새 작업 환경 준비 완료")
            self.log(f" [경로] {self.project_dir}")
            self.log("-" * 70)
            self.log(" ★ 작업 순서:")
            self.log("  1. MP3 파일을 위 작업 폴더에 넣어주세요.")
            self.log("  2. 원본 엑셀 파일을 선택하세요.")
            self.log("  3. 회차별로 자동 분리된 표 데이터를 확인하세요.")
            self.log("  4. 버튼을 클릭하면 회차별로 엑셀이 생성됩니다.")
            self.log("="*70)
            
            os.startfile(self.project_dir)
        except Exception as e:
            self.log(f" [에러] 작업 환경 생성 실패: {str(e)}")

    # 사용자가 선택한 엑셀 파일을 읽고 유효한 데이터 필터링
    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택", initialdir=self.project_dir,
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                self.log(f"\n [알림] 파일 로드 시작: {os.path.basename(file_path)}")
                # 지정된 컬럼 로드 및 이름 설정
                raw_df = pd.read_excel(file_path, usecols=[1, 2, 3, 5, 9, 12], header=None)
                raw_df.columns = ['Start', 'End', 'Length', 'Active_Take_Name', 'File_Path', 'Mute']
                
                # 타임코드 형식이 없는 행 제거
                initial_count = len(raw_df)
                clean_df = raw_df[raw_df['Start'].astype(str).str.contains(r'\d[:;.]\d', na=False)].copy()
                removed_text = initial_count - len(clean_df)
                if removed_text > 0: self.log(f" [알림] 타임코드 형식이 없는 행 {removed_text}개 제외")

                self.process_multi_sessions(clean_df)
                
                self.file_label.config(text=f"현재 로드된 파일: {os.path.basename(file_path)}")
                self.extract_btn.config(state='normal')
                self.progress_label.config(text="추출 준비 완료")
            except Exception as e:
                self.log(f" [에러] 파일 분석 중 문제 발생: {str(e)}")

    # 타임코드 역전 현상을 감지하여 데이터를 회차(Session)별로 분리
    def process_multi_sessions(self, df):
        try:
            df = df.reset_index(drop=True)
            df['time_tuple'] = df.apply(lambda x: self.time_to_tuple(x['Start'], x['Active_Take_Name']), axis=1)
            
            self.sessions = []
            current_session_raw = []
            prev_tuple = (-1, -1, -1, -1)
            
            # 행별로 순회하며 시간 리셋 지점 확인
            for idx, row in df.iterrows():
                curr_tuple = row['time_tuple']
                if curr_tuple < prev_tuple:
                    if current_session_raw:
                        self.sessions.append(self.finalize_session(current_session_raw))
                        current_session_raw = []
                        self.log(f" [알림] 회차 변경 감지 (시간 리셋)")
                current_session_raw.append(row)
                prev_tuple = curr_tuple
            
            # 마지막 회차 추가 및 표 업데이트
            if current_session_raw:
                self.sessions.append(self.finalize_session(current_session_raw))
            self.update_table()
            self.log(f" [성공] 분석 완료: 총 {len(self.sessions)}개 회차")
        except Exception as e:
            self.log(f" [에러] 분석 중단: {str(e)}")

    # 한 회차 내에서 Mute된 행을 제외하고 연속된 동일 트랙을 병합
    def finalize_session(self, row_list):
        session_df = pd.DataFrame(row_list)
        # Mute(X 표시) 데이터 로그 기록 및 필터링
        mute_mask = session_df['Mute'].astype(str).str.upper() == 'X'
        for _, m_row in session_df[mute_mask].iterrows():
            self.log(f" [제외] Mute된 세그먼트: {m_row['Active_Take_Name']}")
        
        filtered_df = session_df[~mute_mask].copy().reset_index(drop=True)
        if filtered_df.empty:
            return filtered_df
        
        merged_rows = []
        skip_idx = set()
        # 동일한 이름의 연속된 트랙 병합 처리
        for i in range(len(filtered_df)):
            if i in skip_idx: continue
            name = filtered_df.loc[i, 'Active_Take_Name']
            group = [filtered_df.loc[i]]
            for j in range(i + 1, len(filtered_df)):
                if filtered_df.loc[j, 'Active_Take_Name'] == name:
                    group.append(filtered_df.loc[j])
                    skip_idx.add(j)
                else: break
            
            if len(group) > 1:
                merged_rows.append(self.merge_group(group))
                self.log(f" [성공] 트랙 병합: '{name}' ({len(group)}개)")
            else:
                merged_rows.append(filtered_df.loc[i])
        return pd.DataFrame(merged_rows)

    # 연속된 트랙 그룹의 시작 시간과 종료 시간을 계산하여 하나로 합침
    def merge_group(self, group):
        first, last = group[0], group[-1]
        start_sec = self.time_to_seconds(first['Start'])
        end_sec = self.time_to_seconds(last['End'])
        merged = first.copy()
        merged['End'] = last['End']
        merged['Length'] = self.seconds_to_time(max(0, end_sec - start_sec))
        return merged

    # 분석된 회차별 데이터를 GUI의 트리뷰 테이블에 출력
    def update_table(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for i, session_df in enumerate(self.sessions, 1):
            for _, row in session_df.iterrows():
                self.tree.insert('', 'end', values=(
                    f"{i}회차", row['Start'], row['End'], row['Length'], 
                    row['Active_Take_Name'], row['File_Path']
                ))

    # 실제 MP3 파일에서 정보를 추출하고 회차별 결과 엑셀 생성
    def start_extraction(self):
        if not self.sessions: return
        self.extract_btn.config(state='disabled')
        try:
            total_sessions = len(self.sessions)
            for s_idx, session_df in enumerate(self.sessions, 1):
                self.log(f"\n▶ [{s_idx:02d}회차 작업]")
                final_data_list = []
                total_rows = len(session_df)
                # 회차 내 각 행별로 메타데이터 추출 실행
                for i, (idx, row) in enumerate(session_df.iterrows()):
                    self.progress_var.set((i / total_rows) * 100)
                    self.progress_label.config(text=f"{s_idx}/{total_sessions}회차 진행 중...")
                    self.root.update()

                    # 파일 경로 확인 및 메타데이터 로드
                    file_path = str(row['File_Path']).strip()
                    orig_name = os.path.basename(file_path)
                    if not os.path.exists(file_path):
                        file_path = os.path.join(self.project_dir, orig_name)
                    
                    if os.path.exists(file_path):
                        meta = self.get_mp3_metadata(file_path)
                        self.log(f" [성공] 데이터 추출: '{orig_name}'")
                    else:
                        self.log(f" [실패] 파일 미발견: '{orig_name}'")
                        meta = self.get_default_meta(orig_name)

                    final_data_list.append({
                        '시작시간': row['Start'], '종료시간': row['End'], '길이': row['Length'],
                        '곡명(Title)': meta['title'], '아티스트(Artist)': meta['Artist'],
                        '앨범(Album)': meta['Album'], '제작사(Publisher)': meta['Publisher'],
                        '작곡가(Composer)': meta['Composer'], '작사가(Lyricist)': meta['Lyricist'],
                        '파일명(File_Name)': orig_name
                    })

                # 추출된 데이터를 회차별 엑셀 파일로 저장
                if final_data_list:
                    output_path = os.path.join(self.project_dir, f'_song_data_{s_idx:02d}.xlsx')
                    pd.DataFrame(final_data_list).to_excel(output_path, index=False)

            # 모든 작업 완료 처리 및 알림
            self.progress_var.set(100)            
            self.progress_label.config(text="모든 회차 작업 완료!")
            self.log("\n" + "="*70)
            self.log(f" [알림] 모든 작업 종료")
            self.log("="*70)            
            messagebox.showinfo("완료", "모든 처리가 완료되었습니다.")
            
        except Exception as e:
            self.log(f" [에러] 추출 프로세스 오류: {str(e)}")
        finally:
            self.extract_btn.config(state='normal')

    # MP3 파일의 ID3 태그를 읽어 곡명, 가수 등 정보를 반환
    def get_mp3_metadata(self, full_path):
        metadata = self.get_default_meta(os.path.basename(full_path))
        if not os.path.exists(full_path): return metadata
        _, ext = os.path.splitext(full_path)
        # MP3 확장자인 경우에만 태그 매핑 실행
        if ext.lower() == '.mp3':
            try:
                audio = mp3(full_path)
                tags = audio.tags if audio.tags else {}
                mapping = {'TIT2': 'title', 'TPE1': 'Artist', 'TALB': 'Album', 
                           'TPUB': 'Publisher', 'TCOM': 'Composer', 'TEXT': 'Lyricist'}
                for tag, key in mapping.items():
                    if tag in tags:
                        metadata[key] = self.decode_safely(tags[tag].text[0])
            except: pass
        return metadata

    # 메타데이터를 찾지 못했을 때 파일명을 제목으로 하는 기본값 생성
    def get_default_meta(self, filename):
        name, _ = os.path.splitext(filename)
        return {'title': name, 'Artist': "", 'Album': "", 'Publisher': "", 'Composer': "", 'Lyricist': ""}

    # 깨진 인코딩(latin-1 등)을 한국어 환경에 맞게 안전하게 복구
    def decode_safely(self, text):
        if not text: return ""
        try:
            byte_data = text.encode('latin-1')
            for enc in ['utf-8', 'cp949', 'euc-kr']:
                try: return byte_data.decode(enc, errors='strict').strip()
                except: continue
        except: pass
        return str(text).strip()

    # '시:분:초.프레임' 문자열을 계산 가능한 초(Seconds) 단위로 변환
    def time_to_seconds(self, time_str):
        try:
            t = str(time_str).strip().replace(';', ':').replace('.', ':')
            parts = t.split(':')
            if len(parts) < 3: return 0
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
            f = int(parts[3]) if len(parts) > 3 else 0
            # 프레임 값에 따른 FPS 추정 및 계산
            fps = 30.0
            if f >= 60: fps = 100.0 
            elif f >= 30: fps = 60.0
            return h * 3600 + m * 60 + s + (f / fps)
        except: return 0

    # 초(Seconds) 단위 수치를 다시 '시:분:초.프레임' 문자열로 변환
    def seconds_to_time(self, total_seconds):
        if total_seconds < 0: total_seconds = 0
        h, m = int(total_seconds // 3600), int((total_seconds % 3600) // 60)
        s = int(total_seconds % 60)
        f = int(round((total_seconds % 1) * 30))
        # 단위 올림 처리
        if f >= 30:
            f, s = 0, s + 1
            if s >= 60: s, m = 0, m + 1
            if m >= 60: m, h = 0, h + 1
        return f"{h:02d}:{m:02d}:{s:02d}.{f:02d}"

    # 타임코드 문자열을 대소 비교가 가능한 튜플 형태로 변환
    def time_to_tuple(self, time_str, track_name="알 수 없음"):
        try:
            t = str(time_str).strip().replace(';', ':').replace('.', ':')
            parts = [int(p) for p in t.split(':') if p.strip().isdigit()]
            if len(parts) >= 4: return tuple(parts[:4])
            elif len(parts) == 3: return (parts[0], parts[1], parts[2], 0)
            return (0, 0, 0, 0)
        except Exception as e:
            self.log(f" [에러] 타임코드 분석 실패 (곡명: {track_name} / 값: {time_str}): {str(e)}")
            return (0, 0, 0, 0)

    # 로그 위젯에 메시지를 추가하고 화면을 강제로 갱신
    def log(self, message):
        self.log_widget.configure(state='normal')
        self.log_widget.insert(tk.END, f"{message}\n")
        self.log_widget.see(tk.END)
        self.log_widget.configure(state='disabled')
        self.root.update_idletasks()

# 프로그램 실행 진입점
if __name__ == "__main__":
    root = tk.Tk()
    app = IntegratedStudioApp(root)
    root.mainloop()