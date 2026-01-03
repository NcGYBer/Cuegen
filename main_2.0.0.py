import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime
import os
from mutagen.mp3 import MP3 as mp3

class IntegratedStudioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("방정리 및 큐시트 생성 통합 도구")
        self.root.geometry("900x850")
        
        self.df = None
        self.sessions = [] 
        self.project_dir = "" 
        
        self.create_widgets()
        self.prepare_workspace()
    
    def create_widgets(self):
        file_frame = ttk.Frame(self.root)
        file_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(file_frame, text="엑셀 파일 선택", command=self.load_file).pack(side='left')
        self.file_label = ttk.Label(file_frame, text="파일이 선택되지 않았습니다")
        self.file_label.pack(side='left', padx=10)
        
        table_frame = ttk.LabelFrame(self.root, text="병합 및 회차 분리 결과 미리보기 (확인용)")
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        columns = ('Session', 'Start', 'End', 'Length', 'Active_Take_Name', 'File_Path')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        # 가로 스크롤바 생성 및 연결
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)

        self.tree.heading('Session', text='회차')
        self.tree.heading('Start', text='시작 시간')
        self.tree.heading('End', text='종료 시간')
        self.tree.heading('Length', text='길이')
        self.tree.heading('Active_Take_Name', text='트랙명')
        self.tree.heading('File_Path', text='파일 경로')

        # 사용자 지정 컬럼 너비 설정 (유지)
        self.tree.column('Session', width=60, anchor='center', stretch=False)
        self.tree.column('Start', width=100, anchor='center', stretch=False)
        self.tree.column('End', width=100, anchor='center', stretch=False)
        self.tree.column('Length', width=100, anchor='center', stretch=False)
        self.tree.column('Active_Take_Name', width=200, anchor='w', stretch=False)
        self.tree.column('File_Path', width=450, anchor='w', stretch=False) # 경로가 길 수 있어 조금 더 늘렸습니다
        
        # 위젯 배치
        scrollbar_x.pack(side='bottom', fill='x')
        scrollbar_y.pack(side='right', fill='y')
        self.tree.pack(side='left', fill='both', expand=True)
        
        log_frame = ttk.LabelFrame(self.root, text="처리 로그 및 안내")
        log_frame.pack(fill='x', expand=False, padx=10, pady=5)
        
        self.log_widget = scrolledtext.ScrolledText(log_frame, height=15, state='disabled', bg="#f8f9fa", font=("맑은 고딕", 9))
        self.log_widget.pack(fill='both', expand=True, padx=5, pady=5)
        
        progress_frame = ttk.Frame(self.root)
        progress_frame.pack(fill='x', padx=15, pady=(5, 0))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', side='top')
        
        self.progress_label = ttk.Label(progress_frame, text="대기 중...")
        self.progress_label.pack(side='top', pady=2)

        process_frame = ttk.Frame(self.root)
        process_frame.pack(fill='x', padx=10, pady=10)
        
        self.extract_btn = ttk.Button(process_frame, text="모든 회차 데이터 추출 및 회차별 엑셀 생성 시작", 
                                     command=self.start_extraction, state='disabled')
        self.extract_btn.pack(fill='x', ipady=10)

    def log(self, message):
        self.log_widget.configure(state='normal')
        self.log_widget.insert(tk.END, f"{message}\n")
        self.log_widget.see(tk.END)
        self.log_widget.configure(state='disabled')
        self.root.update_idletasks()

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

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            initialdir=self.project_dir,
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                self.log(f"\n [알림] 파일 로드 시작: {os.path.basename(file_path)}")
                # 엑셀 기준 M칼럼(index 12)을 포함하여 로드
                raw_df = pd.read_excel(file_path, usecols=[1, 2, 3, 5, 9, 12], header=None)
                raw_df.columns = ['Start', 'End', 'Length', 'Active_Take_Name', 'File_Path', 'Mute']
                
                initial_count = len(raw_df)
                # 1. 타임코드 형식이 없는 텍스트 열 제거
                raw_df = raw_df[raw_df['Start'].astype(str).str.contains(':', na=False)].copy()
                
                # 2. Mute(X) 데이터 필터링 및 지정된 양식으로 로그 출력
                mute_df = raw_df[raw_df['Mute'].astype(str).str.upper() == 'X']
                for _, m_row in mute_df.iterrows():
                    self.log(f" [제외] Mute된 세그먼트: {m_row['Active_Take_Name']}")
                
                mute_count = len(mute_df)
                raw_df = raw_df[raw_df['Mute'].astype(str).str.upper() != 'X'].copy()
                
                removed_text_rows = initial_count - len(raw_df) - mute_count
                if removed_text_rows > 0:
                    self.log(f" [알림] 불필요한 헤더/텍스트 {removed_text_rows}개를 제외했습니다.")
                
                self.log(f" [성공] 데이터 로드 완료 (유효 데이터: {len(raw_df)}개, 제외: {mute_count}개)")
                
                self.log(" [알림] 병합 및 회차 자동 분석 수행 중...")
                self.process_multi_sessions(raw_df)
                
                self.file_label.config(text=f"현재 로드된 파일: {os.path.basename(file_path)}")
                self.extract_btn.config(state='normal')
                self.progress_label.config(text="추출 준비 완료")
            except Exception as e:
                self.log(f" [에러] 파일 분석 중 문제 발생: {str(e)}")

    def process_multi_sessions(self, df):
        try:
            merged_rows = []
            processed_indices = set()
            df = df.reset_index(drop=True)
            
            for idx, row in df.iterrows():
                if idx in processed_indices: continue
                track_name = row['Active_Take_Name']
                current_group = [idx]
                for i in range(idx + 1, len(df)):
                    if df.loc[i, 'Active_Take_Name'] == track_name:
                        current_group.append(i)
                    else:
                        break
                
                if len(current_group) > 1:
                    merged_rows.append(self.merge_group([df.loc[i] for i in current_group]))
                    processed_indices.update(current_group)
                    self.log(f" [성공] 트랙 병합: '{track_name}' ({len(current_group)}개 세그먼트)")
                else:
                    merged_rows.append(row)
                    processed_indices.add(idx)

            all_merged_df = pd.DataFrame(merged_rows)
            all_merged_df['start_sec'] = all_merged_df['Start'].apply(self.time_to_seconds)
            
            self.sessions = []
            current_session_rows = []
            prev_time = -1
            
            for _, row in all_merged_df.iterrows():
                curr_time = row['start_sec']
                if curr_time < prev_time:
                    self.sessions.append(pd.DataFrame(current_session_rows))
                    current_session_rows = []
                    self.log(f" [알림] 회차 변경 감지: {self.seconds_to_time(prev_time)} -> {self.seconds_to_time(curr_time)}")
                current_session_rows.append(row)
                prev_time = curr_time
                
            if current_session_rows:
                self.sessions.append(pd.DataFrame(current_session_rows))

            self.update_table()
            self.log(f" [성공] 분석 완료: 총 {len(self.sessions)}개의 회차가 정리되었습니다.")
            self.log("-" * 70)
        except Exception as e:
            self.log(f" [에러] 분석 공정 중단: {str(e)}")

    def update_table(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for i, session_df in enumerate(self.sessions, 1):
            for _, row in session_df.iterrows():
                self.tree.insert('', 'end', values=(
                    f"{i}회차", row['Start'], row['End'], row['Length'], 
                    row['Active_Take_Name'], row['File_Path']
                ))

    def start_extraction(self):
        if not self.sessions: return
        self.extract_btn.config(state='disabled')
        self.log("\n" + "="*70)
        self.log(f" [알림] 추출 시작: 총 {len(self.sessions)}개 회차 처리 예정")
        self.log("="*70)
        
        try:
            total_sessions = len(self.sessions)
            for s_idx, session_df in enumerate(self.sessions, 1):
                self.log(f"\n▶ [{s_idx:02d}회차 작업]")
                final_data_list = []
                
                total_rows = len(session_df)
                for i, (idx, row) in enumerate(session_df.iterrows()):
                    progress = (i / total_rows) * 100
                    self.progress_var.set(progress)
                    self.progress_label.config(text=f"{s_idx}/{total_sessions}회차 진행 중... ({i+1}/{total_rows})")
                    self.root.update()

                    file_path = str(row['File_Path']).strip()
                    orig_name = os.path.basename(file_path)
                    
                    if not os.path.exists(file_path):
                        file_path = os.path.join(self.project_dir, orig_name)
                    
                    if os.path.exists(file_path):
                        try:
                            meta = self.get_mp3_metadata(file_path)
                            self.log(f" [성공] 데이터 추출: '{orig_name}'")
                        except Exception as e:
                            self.log(f" [주의] 태그 파싱 실패: '{orig_name}' ({str(e)})")
                            meta = self.get_default_meta(orig_name)
                    else:
                        self.log(f" [실패] 파일 미발견: '{orig_name}'")
                        meta = self.get_default_meta(orig_name)

                    combined_row = {
                        '시작시간': row['Start'], '종료시간': row['End'], '길이': row['Length'],
                        '곡명(Title)': meta['title'], '아티스트(Artist)': meta['Artist'],
                        '앨범(Album)': meta['Album'], '제작사(Publisher)': meta['Publisher'],
                        '작곡가(Composer)': meta['Composer'], '작사가(Lyricist)': meta['Lyricist'],
                        '파일명(File_Name)': orig_name
                    }
                    final_data_list.append(combined_row)

                if final_data_list:
                    try:
                        output_filename = f'_song_data_{s_idx:02d}.xlsx'
                        output_path = os.path.join(self.project_dir, output_filename)
                        pd.DataFrame(final_data_list).to_excel(output_path, index=False)
                        self.log(f" [성공] 파일 생성 완료: {output_filename}")
                    except PermissionError:
                        self.log(f" [에러] 파일 저장 실패: {output_filename}이 열려있습니다.")
                    except Exception as se:
                        self.log(f" [에러] 엑셀 쓰기 오류: {str(se)}")
                
                self.log("-" * 50)

            self.progress_var.set(100)
            self.progress_label.config(text="모든 회차 작업 완료!")
            self.log("\n" + "="*70)
            self.log(f" [알림] 모든 작업 종료")
            self.log("="*70)
            messagebox.showinfo("작업 완료", "모든 회차 처리가 완료되었습니다.")
                
        except Exception as e:
            self.log(f"\n [에러] 추출 프로세스 오류: {str(e)}")
        finally:
            self.extract_btn.config(state='normal')

    def get_default_meta(self, filename):
        name, _ = os.path.splitext(filename)
        return {'title': name, 'Artist': "", 'Album': "", 'Publisher': "", 'Composer': "", 'Lyricist': ""}

    def decode_safely(self, text):
        if not text: return ""
        try:
            byte_data = text.encode('latin-1')
            for enc in ['utf-8', 'cp949', 'euc-kr']:
                try: return byte_data.decode(enc, errors='strict').strip()
                except: continue
        except: pass
        return str(text).strip()

    def get_mp3_metadata(self, full_path):
        metadata = self.get_default_meta(os.path.basename(full_path))
        if not os.path.exists(full_path): return metadata
        
        _, ext = os.path.splitext(full_path)
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

    def time_to_seconds(self, time_str):
        try:
            parts = str(time_str).split(':')
            h, m = int(parts[0]), int(parts[1])
            s_f = parts[2].split('.')
            s = int(s_f[0])
            f = int(s_f[1]) if len(s_f) > 1 else 0
            return h * 3600 + m * 60 + s + f / 30.0
        except: return 0

    def seconds_to_time(self, total_seconds):
        h = int(total_seconds // 3600)
        m = int((total_seconds % 3600) // 60)
        s = int(total_seconds % 60)
        f = int((total_seconds % 1) * 30)
        return f"{h:02d}:{m:02d}:{s:02d}.{f:02d}"

    def merge_group(self, group):
        first, last = group[0], group[-1]
        start_sec = self.time_to_seconds(first['Start'])
        end_sec = self.time_to_seconds(last['End'])
        new_len = self.seconds_to_time(end_sec - start_sec)
        merged = first.copy()
        merged['End'] = last['End']
        merged['Length'] = new_len
        return merged

if __name__ == "__main__":
    root = tk.Tk()
    app = IntegratedStudioApp(root)
    root.mainloop()