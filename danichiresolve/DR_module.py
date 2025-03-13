#!/usr/bin/env python

import sys
import time
import subprocess
# sys.path.append('C:\\Program Files\\Blackmagic Design\\DaVinci Resolve\\Support\\Developer\\Scripting\\Modules\\')
import psutil
import DaVinciResolveScript as dvr_script





def turn_off_the_davinci():
    try:
        # psutil 모듈 임포트 필요
        import psutil


        # 현재 실행 중인 Resolve 프로세스 종료
        print("실행 중인 DaVinci Resolve 프로세스를 확인합니다...")
        for proc in psutil.process_iter(['pid', 'name']):
            # 프로세스 이름 확인 (운영체제별로 다를 수 있음)
            if "Resolve" in proc.info['name'] or "resolve" in proc.info['name'].lower():
                print(f"DaVinci Resolve 프로세스를 종료합니다. (PID: {proc.info['pid']})")
                try:
                    # 프로세스 종료
                    process = psutil.Process(proc.info['pid'])
                    process.terminate()
                    # 3초 대기
                    process.wait(3)
                except psutil.NoSuchProcess:
                    pass
                except psutil.TimeoutExpired:
                    print("정상 종료 실패, 강제 종료를 시도합니다...")
                    try:
                        process.kill()
                    except:
                        pass
        print("DaVinci Resolve 프로세스 종료 작업 완료")
        return True
    except Exception as e:
        print(f"DaVinci Resolve 종료 중 오류 발생: {e}")
        return False

def launch_resolve_and_connect(resolve_path, max_attempts=20, check_interval=5):
    """
    DaVinci Resolve를 실행하고 API에 연결하는 함수

    Args:
        resolve_path (str): DaVinci Resolve 실행 파일 경로
        max_attempts (int): 연결 시도 최대 횟수 (기본값: 20회)
        check_interval (int): 각 시도 사이의 대기 시간(초) (기본값: 5초)

    Returns:
        tuple: (process, resolve_instance)
    """
    try:
        # DaVinci Resolve 실행
        process = subprocess.Popen(
            [resolve_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if not process.pid:
            print("DaVinci Resolve를 실행하지 못했습니다.")
            return None, None

        print(f"DaVinci Resolve가 시작되었습니다. PID: {process.pid}")
        print(f"DaVinci Resolve API 연결 시도를 시작합니다. (간격: {check_interval}초, 최대 {max_attempts}회)")

        # 정해진 횟수만큼 연결 시도
        for attempt in range(1, max_attempts + 1):
            print(f"DaVinci Resolve API 연결 시도 중... ({attempt}/{max_attempts})")

            # API 연결 시도
            resolve = get_resolve_instance()

            if resolve:
                print(f"DaVinci Resolve API 연결 성공! (시도 횟수: {attempt})")
                return process, resolve

            # 마지막 시도가 아니면 대기 후 다시 시도
            if attempt < max_attempts:
                print(f"{check_interval}초 후 다시 시도합니다...")
                time.sleep(check_interval)
            else:
                print(f"최대 시도 횟수({max_attempts}회)를 초과했습니다. DaVinci Resolve API 연결 실패.")

        turn_off_the_davinci() #혹시모르니 끄는게 있어야 할 듯
        return process, None

    except Exception as e:
        print(f"DaVinci Resolve 실행 및 연결 중 오류 발생: {e}")
        turn_off_the_davinci()  # 혹시모르니 끄는게 있어야 할 듯
        return None, None

def get_resolve_instance():
    """DaVinci Resolve 인스턴스 가져오기"""
    resolve = dvr_script.scriptapp("Resolve")
    if not resolve:
        print("DaVinci Resolve API에 연결할 수 없습니다.")
        print(resolve)
    return resolve


def delete_all_projects():
    """현재 데이터베이스의 모든 프로젝트 삭제"""
    # Resolve 인스턴스 가져오기
    resolve = get_resolve_instance()

    # 프로젝트 관리자 접근
    project_manager = resolve.GetProjectManager()
    if not project_manager:
        print("프로젝트 관리자에 접근할 수 없습니다.")
        return False

    # 현재 데이터베이스의 모든 프로젝트 가져오기
    project_list = project_manager.GetProjectListInCurrentFolder()
    if not project_list:
        print("현재 폴더에 프로젝트가 없습니다.")
        return True

    print(f"총 {len(project_list)}개 프로젝트를 찾았습니다.")

    # 각 프로젝트 삭제
    deleted_count = 0
    for project_name in project_list:
        print(f"프로젝트 삭제 중: {project_name}")
        success = project_manager.DeleteProject(project_name)
        if success:
            deleted_count += 1
            print(f"- 삭제 완료: {project_name}")
        else:
            print(f"- 삭제 실패: {project_name}")
        time.sleep(0.5)  # 약간의 지연 추가

    print(f"총 {deleted_count}/{len(project_list)}개 프로젝트가 삭제되었습니다.")
    return True


def relaunch_resolve_with_project(resolve_path, project_path, check_interval=5, max_attempts=20):
    """
    기존 DaVinci Resolve를 종료하고 새로 실행한 후 지정된 프로젝트를 여는 함수

    Args:
        resolve_path (str/StringVar): DaVinci Resolve 실행 파일 경로
        project_path (str/StringVar): 열고자 하는 프로젝트 파일 경로 (.drp)
        check_interval (int/StringVar): API 연결 시도 간격(초) (기본값: 5초)
        max_attempts (int/StringVar): 최대 시도 횟수 (기본값: 20회)

    Returns:
        tuple: (bool, project_object 또는 None) - 성공 여부와 프로젝트 객체
    """
    import os

    import subprocess
    import psutil
    import DaVinciResolveScript as dvr_script

    time.sleep(3)
    # StringVar 타입 처리
    if hasattr(resolve_path, 'get'):
        resolve_path = resolve_path.get().strip()

    if hasattr(project_path, 'get'):
        project_path = project_path.get().strip()

    if hasattr(check_interval, 'get'):
        try:
            check_interval = int(check_interval.get())
        except ValueError:
            check_interval = 5  # 기본값으로 설정
    else:
        try:
            check_interval = int(check_interval)
        except ValueError:
            check_interval = 5  # 기본값으로 설정

    if hasattr(max_attempts, 'get'):
        try:
            max_attempts = int(max_attempts.get())
        except ValueError:
            max_attempts = 20  # 기본값으로 설정
    else:
        try:
            max_attempts = int(max_attempts)
        except ValueError:
            max_attempts = 20  # 기본값으로 설정

    # 경로 유효성 확인
    if not os.path.exists(resolve_path):
        print(f"오류: DaVinci Resolve 경로가 유효하지 않습니다: {resolve_path}")
        return False, None

    if not os.path.exists(project_path):
        print(f"오류: 프로젝트 파일 경로가 유효하지 않습니다: {project_path}")
        return False, None

    try:
        # 현재 실행 중인 Resolve 프로세스 종료
        print("실행 중인 DaVinci Resolve 프로세스를 확인합니다...")
        for proc in psutil.process_iter(['pid', 'name']):
            # 프로세스 이름 확인 (운영체제별로 다를 수 있음)
            proc_name = proc.info['name'] if 'name' in proc.info else ""
            if "Resolve" in proc_name or (proc_name and "resolve" in proc_name.lower()):
                print(f"DaVinci Resolve 프로세스를 종료합니다. (PID: {proc.info['pid']})")
                try:
                    # 프로세스 종료
                    process = psutil.Process(proc.info['pid'])
                    process.terminate()
                    # 3초 대기
                    process.wait(3)
                except psutil.NoSuchProcess:
                    pass
                except psutil.TimeoutExpired:
                    print("정상 종료 실패, 강제 종료를 시도합니다...")
                    try:
                        process.kill()
                    except:
                        pass

        print("모든 DaVinci Resolve 프로세스가 종료되었습니다.")
        time.sleep(1)  # 완전히 종료될 때까지 잠시 대기

        # 경로 정규화 및 따옴표 처리 (경로에 공백이 있는 경우)
        resolve_path = os.path.normpath(resolve_path)
        project_path = os.path.normpath(project_path)

        # 프로젝트 경로에서 파일 이름만 추출
        project_name = os.path.basename(project_path)
        if project_name.lower().endswith('.drp'):
            project_name = project_name[:-4]
        project_directory = os.path.dirname(project_path)

        # DaVinci Resolve를 프로젝트와 함께 실행
        cmd = [resolve_path, project_path]
        print(f"DaVinci Resolve를 실행합니다: {' '.join(cmd)}")

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if not process.pid:
            print("DaVinci Resolve를 실행하지 못했습니다.")
            return False, None

        print(f"DaVinci Resolve를 시작했습니다. PID: {process.pid}")
        print(f"{check_interval}초 간격으로 최대 {max_attempts}회 API 연결을 시도합니다...")

        # 여러 번 API 연결 시도
        for attempt in range(1, max_attempts + 1):
            print(f"API 연결 시도 {attempt}/{max_attempts}...")

            try:
                resolve = dvr_script.scriptapp("Resolve")
                if resolve:
                    print("DaVinci Resolve API 연결 성공!")

                    # 프로젝트 확인
                    project_manager = resolve.GetProjectManager()
                    current_project = project_manager.GetCurrentProject()

                    if current_project:
                        project_name = current_project.GetName()
                        print(f"현재 프로젝트: '{project_name}'")
                        return True, current_project
                    else:
                        print("프로젝트를 확인할 수 없습니다. 추가 대기 후 재시도...")
                        # 프로젝트가 아직 로드 중일 수 있으므로 계속 시도
            except Exception as e:
                print(f"API 연결 오류: {str(e)}")

            # 다음 시도 전에 대기
            time.sleep(check_interval)

        print(f"최대 시도 횟수({max_attempts}회)를 초과했습니다. DaVinci Resolve API 연결 실패.")
        turn_off_the_davinci()
        return False, None

    except Exception as e:
        print(f"DaVinci Resolve 실행 및 연결 중 오류 발생: {str(e)}")
        turn_off_the_davinci()
        return False, None


def import_media_to_current_project(file_paths, folder_name=None):
    """
    현재 열려있는 DaVinci Resolve 프로젝트의 미디어 풀에 파일을 임포트하는 함수

    Args:
        file_paths (str 또는 list): 임포트할 파일 경로 (단일 문자열 또는 경로 리스트)
        folder_name (str, optional): 미디어 풀에 생성할 폴더 이름 (없으면 루트에 추가)

    Returns:
        tuple: (bool, str) - 성공 여부와 결과 메시지 또는 오류 메시지
    """
    try:
        # DaVinci Resolve API 연결
        import DaVinciResolveScript as dvr_script
        resolve = dvr_script.scriptapp("Resolve")

        if not resolve:
            return False, "DaVinci Resolve에 연결할 수 없습니다. 프로그램이 실행 중인지 확인하세요."

        # 현재 프로젝트 가져오기
        project_manager = resolve.GetProjectManager()
        current_project = project_manager.GetCurrentProject()

        if not current_project:
            return False, "현재 열려있는 프로젝트가 없습니다."

        # 미디어 풀 객체 가져오기
        media_pool = current_project.GetMediaPool()
        if not media_pool:
            return False, "미디어 풀을 가져올 수 없습니다."

        # 파일 경로가 단일 문자열인 경우 리스트로 변환
        if isinstance(file_paths, str):
            file_paths = [file_paths]

        # 지정된 폴더가 있으면 해당 폴더 생성 또는 선택
        target_folder = None
        if folder_name:
            # 루트 폴더 가져오기
            root_folder = media_pool.GetRootFolder()

            # 폴더 찾기 또는 생성하기
            folders = root_folder.GetSubFolderList()
            for folder in folders:
                if folder.GetName() == folder_name:
                    target_folder = folder
                    break

            # 폴더가 없으면 새로 생성
            if not target_folder:
                target_folder = media_pool.AddSubFolder(root_folder, folder_name)

            # 대상 폴더 설정
            media_pool.SetCurrentFolder(target_folder)

        # 파일 임포트
        imported_clips = media_pool.ImportMedia(file_paths)

        # 임포트 결과 확인
        if imported_clips and len(imported_clips) > 0:
            num_files = len(file_paths)
            num_clips = len(imported_clips)

            if num_clips >= num_files:
                return True, f"모든 파일({num_clips}개)을 성공적으로 임포트했습니다."
            else:
                return True, f"일부 파일만 임포트했습니다. (요청: {num_files}개, 성공: {num_clips}개)"
        else:
            return False, "파일 임포트에 실패했습니다."

    except Exception as e:
        return False, f"파일 임포트 중 오류가 발생했습니다: {e}"


def import_media_to_timeline_with_all_tracks(file_paths):
    """
    미디어 파일을 모든 트랙(비디오 및 오디오)을 유지한 채로 타임라인에 임포트하는 함수

    Args:
        file_paths (str 또는 list): 임포트할 파일 경로 (단일 문자열 또는 경로 리스트)

    Returns:
        tuple: (bool, str) - 성공 여부와 결과 메시지 또는 오류 메시지
    """
    try:
        # DaVinci Resolve API 연결
        import DaVinciResolveScript as dvr_script
        resolve = dvr_script.scriptapp("Resolve")

        if not resolve:
            return False, "DaVinci Resolve에 연결할 수 없습니다. 프로그램이 실행 중인지 확인하세요."
        time.sleep(1)
        # 현재 프로젝트 가져오기
        project_manager = resolve.GetProjectManager()
        current_project = project_manager.GetCurrentProject()
        time.sleep(1)
        if not current_project:
            return False, "현재 열려있는 프로젝트가 없습니다."
        time.sleep(1)
        # 미디어 풀 객체 가져오기
        media_pool = current_project.GetMediaPool()
        if not media_pool:
            return False, "미디어 풀을 가져올 수 없습니다."
        time.sleep(1)
        # 현재 타임라인 가져오기
        timeline = current_project.GetCurrentTimeline()
        if not timeline:
            # 타임라인이 없으면 새로 생성
            timeline_name = "New Timeline"
            timeline = media_pool.CreateEmptyTimeline(timeline_name)
            if not timeline:
                return False, "타임라인을 생성할 수 없습니다."
            print(f"새 타임라인 '{timeline_name}'을 생성했습니다.")
        time.sleep(1)
        # 파일 경로가 단일 문자열인 경우 리스트로 변환
        if isinstance(file_paths, str):
            file_paths = [file_paths]
        time.sleep(1)
        # 미디어 풀에 먼저 파일을 임포트
        root_folder = media_pool.GetRootFolder()
        media_pool.SetCurrentFolder(root_folder)
        time.sleep(1)
        imported_clips = media_pool.ImportMedia(file_paths)
        if not imported_clips or len(imported_clips) == 0:
            return False, "미디어 풀에 파일을 임포트하지 못했습니다."
        time.sleep(1)
        # 현재 타임라인의 끝 위치 찾기
        timeline_end_frame = timeline.GetEndFrame()
        time.sleep(1)
        # 타임라인에 클립 추가 (모든 트랙 유지)
        successfully_added = 0
        for clip in imported_clips:
            try:
                # 모든 트랙을 유지한 채로 타임라인에 추가
                # AppendToTimeline 대신 InsertClipIntoTimeline 사용 (전체 트랙 유지)
                append_result = media_pool.AppendToTimeline([clip])

                if append_result:
                    successfully_added += 1
            except Exception as clip_error:
                print(f"클립 추가 중 오류 발생: {clip_error}")
        time.sleep(1)
        # 결과 반환
        if successfully_added > 0:
            return True, f"{successfully_added}개의 클립을 모든 트랙을 유지한 채로 타임라인에 추가했습니다."
        else:
            return False, "타임라인에 클립을 추가하지 못했습니다."

    except Exception as e:
        return False, f"파일 임포트 중 오류가 발생했습니다: {e}"


def get_available_presets():
    """
    DaVinci Resolve에서 사용 가능한 렌더 프리셋 목록을 가져오는 함수

    Returns:
        list: 사용 가능한 프리셋 이름 목록. 연결 실패 시 빈 리스트 반환
    """
    try:


        # DaVinci Resolve API 연결
        resolve = dvr_script.scriptapp("Resolve")

        if not resolve:
            print("DaVinci Resolve에 연결할 수 없습니다.")
            return []

        # 현재 프로젝트 가져오기
        project_manager = resolve.GetProjectManager()
        current_project = project_manager.GetCurrentProject()

        if not current_project:
            print("현재 열려있는 프로젝트가 없습니다.")
            return []

        # Deliver 페이지로 전환
        current_page = resolve.GetCurrentPage()
        if current_page != "deliver":
            resolve.OpenPage("deliver")
            print("Deliver 페이지로 전환했습니다.")
            time.sleep(1)  # 페이지 전환 대기

        # 사용 가능한 렌더 프리셋 목록 가져오기
        presets = []

        try:
            # API 버전에 따라 다른 메서드 이름을 사용할 수 있음
            if hasattr(current_project, 'GetRenderPresetList'):
                presets = current_project.GetRenderPresetList()
                print(f"{len(presets)}개의 렌더 프리셋을 찾았습니다.")
            elif hasattr(current_project, 'GetRenderPresets'):
                presets = current_project.GetRenderPresets()
                print(f"{len(presets)}개의 렌더 프리셋을 찾았습니다.")
            else:
                print("프리셋 목록을 가져올 수 있는 API 메서드를 찾을 수 없습니다.")
        except Exception as e:
            print(f"프리셋 목록 가져오기 실패: {e}")
            import traceback
            traceback.print_exc()

        return presets

    except Exception as e:
        print(f"예상치 못한 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return []


def render_with_preset(file_name,output_folder, preset):
    """
    DaVinci Resolve의 Deliver 페이지로 이동하여
    지정된 프리셋을 사용해 현재 타임라인을 렌더링하는 함수

    Args:
        preset_path (str): 사용할 프리셋 파일(.drx) 경로
        output_folder (str): 렌더링된 파일을 저장할 출력 폴더 경로
        preset (str): 사용할 프리셋 이름

    Returns:
        tuple: (bool, str) - 성공 여부와 결과 메시지 또는 오류 메시지
    """
    try:
        # DaVinci Resolve API 연결
        import DaVinciResolveScript as dvr_script
        resolve = dvr_script.scriptapp("Resolve")

        if not resolve:
            return False, "DaVinci Resolve에 연결할 수 없습니다. 프로그램이 실행 중인지 확인하세요."

        # 현재 프로젝트 가져오기
        project_manager = resolve.GetProjectManager()
        current_project = project_manager.GetCurrentProject()

        if not current_project:
            return False, "현재 열려있는 프로젝트가 없습니다."

        # 현재 타임라인 확인
        timeline = current_project.GetCurrentTimeline()
        if not timeline:
            return False, "현재 타임라인이 없습니다."

        # Deliver 페이지로 전환
        page_type = resolve.GetCurrentPage()
        if page_type != "deliver":
            resolve.OpenPage("deliver")
            print("Deliver 페이지로 전환했습니다.")

        # Deliver 페이지 객체 가져오기
        deliver_page = resolve.GetCurrentPage()

        # 기존 렌더 설정 모두 제거
        deliver = current_project.GetRenderJobList()
        if deliver and len(deliver) > 0:
            current_project.DeleteAllRenderJobs()
            print("기존 렌더 작업을 모두 제거했습니다.")

        # 지정된 이름(preset)의 프리셋 로드
        preset_loaded = current_project.LoadRenderPreset(preset)
        if not preset_loaded:
            return False, f"프리셋 '{preset}'을 로드하지 못했습니다."

        print(f"프리셋 '{preset}'을 성공적으로 로드했습니다.")

        # 출력 경로 설정
        import os
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"출력 폴더 '{output_folder}'를 생성했습니다.")


        # 출력 파일 이름 설정
        output_filename = file_name
        print(f"{output_filename} 파일명 입니다.")
        # 렌더 설정
        render_settings = {
            "TargetDir": output_folder,
            "CustomName": output_filename
        }

        # 렌더 작업 추가
        current_project.SetRenderSettings(render_settings)
        add_result = current_project.AddRenderJob()

        time.sleep(1)
        if not add_result:
            return False, "렌더 작업을 추가하지 못했습니다."

        print("렌더 작업이 추가되었습니다.")
        # 렌더링 시작
        render_result = current_project.StartRendering()

        if not render_result:
            return False, "렌더링을 시작하지 못했습니다."

        print("렌더링이 시작되었습니다...")

        while current_project.IsRenderingInProgress():
            job_status = current_project.GetRenderJobStatus()
            if job_status:
                for job in job_status:
                    percentage_complete = job.get("PercentageComplete", "알 수 없음")
                    print(f"렌더링 진행 중: {percentage_complete}% 완료")
            time.sleep(3)

        print("렌더링이 완료되었습니다!")

        # 렌더링된 파일 경로 확인
        render_job_list = current_project.GetRenderJobList()
        rendered_filepath = ""

        if render_job_list and len(render_job_list) > 0:
            last_job = render_job_list[-1]
            if "OutputFilename" in last_job:
                rendered_filename = last_job["OutputFilename"]
                rendered_filepath = os.path.join(output_folder, rendered_filename)

        # 렌더링 완료 후 모든 렌더 작업 제거
        current_project.DeleteAllRenderJobs()
        print("렌더링 완료 후 모든 렌더 작업을 제거했습니다.")
        time.sleep(2)
        turn_off_the_davinci()
        time.sleep(2)
        if rendered_filepath:
            return True, f"렌더링이 성공적으로 완료되었습니다: {rendered_filepath}"
        else:
            return False, f"렌더링이 실패했습니다. 대상 폴더: {output_folder}"

    except Exception as e:
        # 오류 발생 시에도 렌더 작업 정리 시도
        try:
            current_project.DeleteAllRenderJobs()
            print("오류 발생 후 렌더 작업을 정리했습니다.")
        except:
            pass

        turn_off_the_davinci()
        time.sleep(2)
        return False, f"렌더링 중 오류가 발생했습니다: {e}"








