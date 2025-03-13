import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import subprocess
import time
import winreg
import psutil
import datetime
import shutil
import socket
import random




import DR_module
import os_module

class DaVinciResolveApp:
    def __init__(self, root):
        def find_drp_file(target_filename="start_proj.drp"):
            """
            sys.executable의 하위 폴더를 모두 검색하여 지정된 .drp 파일을 찾는 함수

            Args:
                target_filename (str): 찾을 파일 이름 (기본값: "start_proj.drp")

            Returns:
                str: 찾은 파일의 전체 경로. 파일을 찾지 못한 경우 None 반환
            """
            # Python 실행 파일 경로 확인
            sys_root = os.path.dirname(sys.executable)
            print(f"검색 시작 경로: {sys_root}")

            # 찾은 파일 경로를 저장할 변수
            found_file = None

            # 하위 폴더 탐색
            for root, dirs, files in os.walk(sys_root):
                if target_filename in files:
                    found_file = os.path.join(root, target_filename)
                    print(f"파일 발견: {found_file}")
                    return found_file

            # 메시지 박스 표시를 위한 임시 루트 윈도우 생성
            root = tk.Tk()
            root.withdraw()  # 루트 윈도우 숨기기

            # 메시지 박스 표시
            messagebox.showerror("파일 없음", f"'{target_filename}' 같이 들어있는 기본 프로젝트파일을 찾을 수 없습니다. 배포한 그대로 사용하세요")

            # 프로그램 종료
            sys.exit(1)

        DR_module.turn_off_the_davinci()
        time.sleep(5)

        self.root = root
        self.root.title("DaVinci Resolve 작업 자동화")
        self.root.geometry("600x400")
        self.root.resizable(True, True)

        # 변수 선언

        self.preset = tk.StringVar()
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.project_path = tk.StringVar()
        self.project_path = find_drp_file(target_filename="start_proj.drp")
        self.resolve_path = tk.StringVar()  # 다빈치 리졸브 경로 추가

        self.able_to_start = None
        self.processing = False
        self.process_thread = None

        # 기본값 설정


        # UI 구성
        self.create_widgets()



    def create_widgets(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 레이아웃 설정
        main_frame.columnconfigure(1, weight=1)
        # DaVinci Resolve 경로
        ttk.Label(main_frame, text="DaVinci Resolve 경로:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.resolve_entry = ttk.Entry(main_frame, textvariable=self.resolve_path)
        self.resolve_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.resolve_button = ttk.Button(main_frame, text="찾기", command=self.browse_resolve_path)
        self.resolve_button.grid(row=0, column=2, pady=5)

        # 입력 폴더
        ttk.Label(main_frame, text="입력 폴더:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.input_entry = ttk.Entry(main_frame, textvariable=self.input_folder)
        self.input_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.input_button = ttk.Button(main_frame, text="찾기", command=self.browse_input_folder)
        self.input_button.grid(row=2, column=2, pady=5)

        # 출력 폴더
        ttk.Label(main_frame, text="출력 폴더:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_entry = ttk.Entry(main_frame, textvariable=self.output_folder)
        self.output_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.output_button = ttk.Button(main_frame, text="찾기", command=self.browse_output_folder)
        self.output_button.grid(row=3, column=2, pady=5)

        # 프리셋 XML 파일 선택 (드롭다운 리스트로 변경)
        ttk.Label(main_frame, text="프리셋 선택:").grid(row=4, column=0, sticky=tk.W, pady=5)
        preset_options = ["다빈치 경로 지정 후 생성"]  # 선택 가능한 옵션들
        self.default_preset_option = preset_options[0]
        self.preset_combobox = ttk.Combobox(main_frame, textvariable=self.preset, values=preset_options)
        self.preset_combobox.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.preset_combobox.current(0)  # 기본값으로 첫 번째 항목 선택

        # # DaVinci Resolve 경로
        # ttk.Label(main_frame, text="DaVinci Resolve 경로:").grid(row=0, column=0, sticky=tk.W, pady=5)
        # ttk.Entry(main_frame, textvariable=self.resolve_path).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        # ttk.Button(main_frame, text="찾기", command=self.browse_resolve_path).grid(row=0, column=2, pady=5)
        #
        #
        # # 입력 폴더
        # ttk.Label(main_frame, text="입력 폴더:").grid(row=2, column=0, sticky=tk.W, pady=5)
        # ttk.Entry(main_frame, textvariable=self.input_folder).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        # ttk.Button(main_frame, text="찾기", command=self.browse_input_folder).grid(row=2, column=2, pady=5)
        #
        # # 출력 폴더
        # ttk.Label(main_frame, text="출력 폴더:").grid(row=3, column=0, sticky=tk.W, pady=5)
        # ttk.Entry(main_frame, textvariable=self.output_folder).grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5,
        #                                                             pady=5)
        # ttk.Button(main_frame, text="찾기", command=self.browse_output_folder).grid(row=3, column=2, pady=5)
        #
        # # 프리셋 XML 파일 선택 (드롭다운 리스트로 변경)
        # ttk.Label(main_frame, text="프리셋 선택:").grid(row=4, column=0, sticky=tk.W, pady=5)
        # preset_options = ["다빈치 경로 지정 후 생성"]  # 선택 가능한 옵션들
        # self.default_preset_option = preset_options[0]
        # self.preset_combobox = ttk.Combobox(main_frame, textvariable=self.preset, values=preset_options)
        # self.preset_combobox.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        # self.preset_combobox.current(0)  # 기본값으로 첫 번째 항목 선택

        # 상태 표시 프레임
        status_frame = ttk.LabelFrame(main_frame, text="작업 상태")
        status_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=10)
        status_frame.columnconfigure(0, weight=1)

        # 진행 상태 바
        self.progress = ttk.Progressbar(status_frame, orient=tk.HORIZONTAL, length=100, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)

        # 상태 메시지
        self.status_var = tk.StringVar(value="준비")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var)
        self.status_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)

        # 로그 출력 영역
        ttk.Label(main_frame, text="작업 로그:").grid(row=6, column=0, sticky=tk.W, pady=2)

        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        # 스크롤바 추가
        log_scrollbar = ttk.Scrollbar(log_frame)
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # 로그 텍스트 영역
        self.log_text = tk.Text(log_frame, height=6, width=60, wrap=tk.WORD,
                                yscrollcommand=log_scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.config(command=self.log_text.yview)

        # 시작/중지 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=3, sticky=tk.E, pady=10)

        self.start_button = ttk.Button(button_frame, text="작업 시작", command=self.start_processing)
        self.start_button.pack(side=tk.RIGHT, padx=5)

        self.stop_button = ttk.Button(button_frame, text="작업 중지", command=self.stop_processing, state=tk.DISABLED)
        self.stop_button.pack(side=tk.RIGHT, padx=5)

        # 전체적인 리사이징 설정
        for i in range(9):
            main_frame.rowconfigure(i, weight=1 if i == 7 else 0)

        # 패딩 설정
        for child in main_frame.winfo_children():
            child.grid_configure(padx=5, pady=2)

    def browse_resolve_path(self):
        """DaVinci Resolve 실행 파일 찾기"""
        path = filedialog.askopenfilename(
            title="DaVinci Resolve 실행 파일 선택",
            filetypes=[("실행 파일", "*.exe"), ("모든 파일", "*.*")]
        )
        if path:
            self.resolve_path.set(path)

            # 루트 윈도우를 얻는 방법 (일반적인 tkinter 애플리케이션 구조에 따라 선택)
            if hasattr(self, 'master'):
                root = self.master
            elif hasattr(self, 'root'):
                root = self.root
            elif hasattr(self, 'parent'):
                root = self.parent
            else:
                # 어떤 루트 윈도우 참조도 찾을 수 없을 경우 새 Tk 인스턴스 생성
                root = tk.Tk()
                root.withdraw()  # 메인 윈도우는 숨김

            # 정보 메시지 창 생성
            info_window = tk.Toplevel(root)
            info_window.title("작업 진행 중")
            info_window.geometry("400x150")
            info_window.resizable(False, False)

            # 창을 화면 중앙에 배치
            screen_width = info_window.winfo_screenwidth()
            screen_height = info_window.winfo_screenheight()
            x = (screen_width - 400) // 2
            y = (screen_height - 150) // 2
            info_window.geometry(f"400x150+{x}+{y}")

            # 메시지 표시
            message_label = tk.Label(
                info_window,
                text="DaVinci Resolve 작동 여부를 확인합니다.\n이 과정은 최대 3분까지 소요될 수 있습니다.\n기본 버전이 아닌 <스튜디오> 버전만 지원합니다",
                font=("Helvetica", 10),
                padx=20,
                pady=20
            )
            message_label.pack(expand=True)

            # 작업 진행 상태 표시
            status_var = tk.StringVar(value="작업 시작 중...")
            status_label = tk.Label(
                info_window,
                textvariable=status_var,
                font=("Helvetica", 9),
                fg="blue"
            )
            status_label.pack(pady=(0, 15))

            # 창을 최상위로 유지
            info_window.attributes('-topmost', True)
            info_window.update()

            # 백그라운드에서 실행할 작업 함수
            def run_tasks():
                try:
                    # 작업 1
                    status_var.set("DaVinci Resolve 연결 중...")
                    info_window.update()
                    DR_module.launch_resolve_and_connect(path)

                    # 작업 2
                    status_var.set("프로젝트 정리 중...")
                    info_window.update()
                    DR_module.delete_all_projects()

                    # 작업 3
                    status_var.set("DaVinci Resolve 종료 중...")
                    info_window.update()
                    DR_module.turn_off_the_davinci()

                    # 대기
                    status_var.set("잠시 대기 중...")
                    info_window.update()
                    time.sleep(3)

                    # 작업 4
                    status_var.set("DaVinci Resolve 재시작 중...")
                    info_window.update()
                    DR_module.relaunch_resolve_with_project(path, self.project_path)
                    time.sleep(5)

                    # 작업 5

                    # 콤보박스의 값 목록만 변경
                    list_of_preset = list(reversed(DR_module.get_available_presets()))
                    self.preset_combobox['values'] = list_of_preset[0:10]
                    self.preset_combobox.current(0)
                    DR_module.turn_off_the_davinci()

                    # 작업 완료
                    status_var.set("작업이 완료되었습니다!")
                    info_window.update()
                    self.able_to_start = True

                    # 2초 후 창 닫기
                    info_window.after(2000, info_window.destroy)

                except Exception as e:
                    # 오류 발생 시
                    status_var.set(f"오류 발생: {str(e)} 작업을 진행할 수 없는 pc 같습니다.")
                    info_window.update()

                    # 5초 후 창 닫기
                    info_window.after(5000, info_window.destroy)

            # 메시지 창을 표시한 후 100ms 후에 작업 시작
            info_window.after(100, run_tasks)

    def browse_input_folder(self):
        """입력 폴더 선택"""
        folder = filedialog.askdirectory(title="입력 폴더 선택")
        if folder:
            self.input_folder.set(folder)

    def browse_output_folder(self):
        """출력 폴더 선택"""
        folder = filedialog.askdirectory(title="출력 폴더 선택")
        if folder:
            self.output_folder.set(folder)

    def browse_preset_file(self):
        """프리셋 XML 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="프리셋 XML 파일 선택",
            filetypes=[("XML 파일", "*.xml"), ("모든 파일", "*.*")]
        )
        if file_path:
            self.preset.set(file_path)

    def validate_inputs(self):
        """입력값 유효성 검사"""
        # DaVinci Resolve 경로 확인
        resolve_path = self.resolve_path.get().strip()
        if not resolve_path:
            messagebox.showerror("오류", "DaVinci Resolve 경로를 입력해주세요.")
            return False

        # 바로가기 파일인 경우 실제 대상 파일 경로 가져오기
        if resolve_path.lower().endswith('.lnk'):
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(resolve_path)
                resolve_path = shortcut.Targetpath
                self.resolve_path.set(resolve_path)
            except Exception as e:
                self.log(f"바로가기 파일 처리 중 오류: {e}")
                messagebox.showerror("오류", f"바로가기 파일을 처리할 수 없습니다: {e}")
                return False

        # 경로 존재 확인
        if not os.path.exists(resolve_path):
            messagebox.showerror("오류", "DaVinci Resolve 경로가 유효하지 않습니다.")
            return False

        # 입력 폴더 확인
        input_folder = self.input_folder.get().strip()
        if not input_folder or not os.path.exists(input_folder):
            messagebox.showerror("오류", "유효한 입력 폴더를 선택해주세요.")
            return False

        # 출력 폴더 확인
        output_folder = self.output_folder.get().strip()
        if not output_folder:
            messagebox.showerror("오류", "출력 폴더를 선택해주세요.")
            return False

        # 출력 폴더가 없으면 생성 확인
        if not os.path.exists(output_folder):
            result = messagebox.askyesno("확인", "출력 폴더가 존재하지 않습니다. 생성하시겠습니까?")
            if result:
                try:
                    os.makedirs(output_folder)
                except Exception as e:
                    self.log(f"폴더 생성 오류: {e}")
                    messagebox.showerror("오류", f"출력 폴더를 생성할 수 없습니다: {e}")
                    return False
            else:
                return False

        return True

    def start_processing(self):

        if self.preset == self.default_preset_option:
            import tkinter.messagebox as messagebox
            messagebox.showerror(
                "다빈치 리졸브 오류",
                "다빈치 리졸브의 경로에서 작동 가능여부를 체크받지 못했습니다.\n경로를 입력하거나 <스튜디오버전>을 사용해 주세요.\n(베이직 버전은 프로그래밍 안됨)"
            )
            return


        """작업 시작"""
        if self.processing:
            return

        # 입력 유효성 검사
        if not self.validate_inputs():
            return

        # UI 상태 업데이트
        self.processing = True
        self.resolve_button.config(state=tk.DISABLED)
        self.preset_combobox.config(state=tk.DISABLED)
        self.input_button.config(state=tk.DISABLED)
        self.output_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)


        self.progress.start(10)
        self.status_var.set("작업 준비 중...")

        # 작업 내용 로그에 기록
        self.log_text.delete(1.0, tk.END)
        self.log(f"=== 작업 시작: {time.strftime('%Y-%m-%d %H:%M:%S')} ===")
        self.log(f"DaVinci Resolve 경로: {self.resolve_path.get()}")
        self.log(f"입력 폴더: {self.input_folder.get()}")
        self.log(f"출력 폴더: {self.output_folder.get()}")
        self.log(f"선택된 프리셋: {self.preset.get()}")

        # 작업 스레드 시작
        self.process_thread = threading.Thread(target=self.process_task)
        self.process_thread.daemon = True
        self.process_thread.start()

    def stop_processing(self):
        """작업 중지"""
        if not self.processing:
            return

        if messagebox.askyesno("확인", "현재 진행 중인 작업을 중지하시겠습니까?"):
            self.processing = False
            self.status_var.set("작업 중지 중...")
            self.log("사용자에 의해 작업이 중지되었습니다.")
            sys.exit()

    def process_task(self):

        def add_timestamp_to_filename(file_path):
            # 파일 경로에서 디렉토리와 파일명 분리
            directory, filename = os.path.split(file_path)

            # 파일명에서 이름과 확장자 분리
            name, extension = os.path.splitext(filename)

            # 현재 시각 가져오기 (원하는 형식으로 포맷팅)
            current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

            # 새 파일명 생성 (이름 + 시각 + 확장자)
            new_filename = f"{name}_{current_time}{extension}"

            def get_filename_only(file_path):
                # 파일 경로에서 파일명만 추출 (디렉토리 제외)
                _, filename = os.path.split(file_path)

                # 파일명에서 확장자 제외
                name, _ = os.path.splitext(filename)

                return name

            new_filename =get_filename_only(new_filename)
            print(new_filename)
            # 새 파일 경로 반환
            return os.path.join(directory, new_filename)


        """실제 작업을 수행하는 메소드 (별도 스레드에서 실행)"""
        while True:
            print(f"{self.default_preset_option} 디폴트 옵션 체크 ")
            print("루프를 돌며 작업을 시작합니다.")
            DR_module.turn_off_the_davinci()

            try:

                # 작업 설정값 가져오기
                resolve_path = self.resolve_path.get()
                resolve_path = os.path.normpath(resolve_path)
                project_path = self.project_path
                project_path = os.path.normpath(project_path)
                input_folder = self.input_folder.get()
                input_folder = os.path.normpath(input_folder)
                output_folder = self.output_folder.get()
                output_folder = os.path.normpath(output_folder)
                preset = str(self.preset.get())

                # 현재 머신의 IP 주소 가져오기
                ip = socket.gethostbyname(socket.gethostname())

                # temp+ip 폴더 경로 생성
                temp_folder_name = f"temp_{ip}"
                temp_folder_path = os.path.join(input_folder, temp_folder_name)
                os_module.move_contents_to_parent(temp_folder_path)
                ##### temp 폴더안에 뭐가 있음 빼야함;; 인
                result_output_folder = os_module.make_folder(output_folder)
                os_module.move_contents_to_parent(result_output_folder)
                ##### temp 폴더안에 뭐가 있음 빼야함;; 아웃


                list_of_file = os_module.get_video_files_list(input_folder)
                random.shuffle(list_of_file)

                for file in list_of_file:
                    # if os_module.is_file_ready(file, check_duration=23, check_count=3):
                    if os_module.is_file_ready(file, check_duration=23, check_count=3):
                        print("인아웃 파일 및 폴더 경로 바꾸고 싶으면 작성 -- 시작")
                        original_file_path = file
                        file = os_module.move_file(file)
                        print(f"파일은{file}")
                        # result_output_folder = os_module.make_folder(output_folder) 위에 작성
                        print("인아웃 파일 및 폴더 경로 바꾸고 싶으면 작성  -- 끝")
                        DR_module.launch_resolve_and_connect(resolve_path)
                        time.sleep(1)
                        DR_module.delete_all_projects()
                        time.sleep(1)
                        DR_module.relaunch_resolve_with_project(resolve_path, project_path)
                        print("여기부터 특별히 넣을 함수나 작업이 있으면 작성 -- 시작")
                        subwork()
                        print("여기부터 특별히 넣을 함수나 작업이 있으면 작성 -- 끝")
                        DR_module.import_media_to_current_project(file)
                        DR_module.import_media_to_timeline_with_all_tracks(file)
                        time.sleep(3)
                        file_name =add_timestamp_to_filename(file)
                        file_name = os.path.basename(file_name)
                        DR_module.render_with_preset(file_name,result_output_folder,str(preset))
                        time.sleep(2)
                        os_module.move_contents_to_parent(result_output_folder)
                        time.sleep(2)
                        os_module.move_finish_file(file)
                        ## os.remove(file) 추후에는 remove나 move 필요할 듯

            except Exception as e:
                # 오류 발생 시 처리
                # error_msg = f"작업 중 오류 발생: {str(e)}"
                # self.log(error_msg)
                # self.update_status("오류 발생")
                # messagebox.showerror("오류", error_msg)
                print(e)
                pass
            finally:
                # UI 상태 복원
                self.processing = False
                self.progress.stop()
                self.start_button.config(state=tk.NORMAL)
                self.stop_button.config(state=tk.DISABLED)

            print('300초 쉬고 다시 진행 합니다.')
            time.sleep(300)


    def update_status(self, message):
        """스레드에서 안전하게 상태 메시지 업데이트"""
        self.root.after(0, lambda: self.status_var.set(message))

    def log(self, message):
        """작업 로그에 메시지 추가"""

        def _log():
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)  # 스크롤을 자동으로 맨 아래로

        self.root.after(0, _log)

def subwork():
    pass




def main():
    root = tk.Tk()
    app = DaVinciResolveApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()