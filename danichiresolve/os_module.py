# -*- coding: utf-8 -*-
import os
import shutil
import time
from pymediainfo import MediaInfo
import platform
import random
import sys
import socket

import os
import shutil
import time
import socket
import os
import shutil
import time
from datetime import datetime


def move_finish_file(file_path):
    """
    파일을 같은 디렉토리 내의 'finish' 폴더로 이동시킵니다.
    동일한 파일명이 이미 존재할 경우 파일명 뒤에 현재 시간을 추가합니다.

    Args:
        file_path (str): 이동할 파일 경로

    Returns:
        str: 이동된 파일의 새 경로 또는 오류 시 None
    """
    try:
        # 파일이 존재하는지 확인
        if not os.path.exists(file_path):
            print(f"오류: {file_path} 파일이 존재하지 않습니다.")
            return None

        # 디렉토리 경로와 파일명 분리
        dir_path = os.path.dirname(file_path)
        parent_dir = os.path.dirname(dir_path)
        file_name = os.path.basename(file_path)
        name, ext = os.path.splitext(file_name)

        # finish 폴더 경로 생성
        finish_folder = os.path.join(parent_dir, "finish")

        # finish 폴더가 없으면 생성
        if not os.path.exists(finish_folder):
            os.makedirs(finish_folder)
            print(f"'finish' 폴더를 생성했습니다: {finish_folder}")

        # 이동할 대상 경로
        destination = os.path.join(finish_folder, file_name)

        # 같은 이름의 파일이 이미 존재하는지 확인
        if os.path.exists(destination):
            # 현재 시간을 파일명에 추가
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            new_file_name = f"{name}_{timestamp}{ext}"
            destination = os.path.join(finish_folder, new_file_name)
            print(f"같은 이름의 파일이 존재하여 {new_file_name}으로 이름을 변경합니다.")

        # 파일 이동
        shutil.move(file_path, destination)
        print(f"파일을 {destination}으로 이동했습니다.")

        return destination

    except Exception as e:
        print(f"파일 이동 중 오류 발생: {e}")
        return None

def move_contents_to_parent(directory_path):
    """
    지정된 디렉토리 내의 모든 파일과 폴더를 상위 디렉토리로 이동시킵니다.

    Args:
        directory_path (str): 콘텐츠를 이동시킬 디렉토리 경로

    Returns:
        bool: 성공 시 True, 실패 시 False
    """
    try:
        # 상위 디렉토리 경로 가져오기
        parent_dir = os.path.dirname(directory_path)

        if not os.path.exists(parent_dir):
            print(f"상위 디렉토리 {parent_dir}가 존재하지 않습니다.")
            return False

        # 디렉토리 내의 모든 항목 확인
        items = os.listdir(directory_path)

        for item in items:
            # 원본 경로
            src_path = os.path.join(directory_path, item)
            # 대상 경로
            dst_path = os.path.join(parent_dir, item)

            # 이미 상위 폴더에 같은 이름의 파일이 있는지 확인
            if os.path.exists(dst_path):
                print(f"경고: {dst_path}가 이미 존재합니다. {item} 건너뜁니다.")
                continue

            # 파일/폴더 이동
            shutil.move(src_path, parent_dir)
            print(f"{src_path}를 {parent_dir}로 이동했습니다.")

        print(f"{directory_path} 내의 모든 항목을 {parent_dir}로 이동 완료했습니다.")
        return True

    except Exception as e:
        print(f"이동 중 오류 발생: {e}")
        return False

def make_folder(dir_path):
    try:
        # 현재 머신의 IP 주소 가져오기
        ip = socket.gethostbyname(socket.gethostname())

        # temp+ip 폴더 경로 생성
        temp_folder_name = f"temp_{ip}"
        temp_folder_path = os.path.join(dir_path, temp_folder_name)

        # 폴더가 없으면 생성
        if not os.path.exists(temp_folder_path):
            os.makedirs(temp_folder_path)
            print(f"{temp_folder_path} 폴더를 생성했습니다.")
        else:
            print(f"{temp_folder_path} 폴더가 이미 존재합니다.")

        move_contents_to_parent(temp_folder_path)
        return temp_folder_path
    except Exception as e:
        print(f"폴더 생성 중 오류 발생: {e}")
        return None

def move_file(file_path):

    try:
        # 현재 머신의 IP 주소 가져오기
        ip = socket.gethostbyname(socket.gethostname())

        # file_path의 디렉토리 경로 가져오기
        dir_path = os.path.dirname(file_path)

        # temp+ip 폴더 경로 생성
        temp_folder_name = f"temp_{ip}"
        temp_folder_path = os.path.join(dir_path, temp_folder_name)

        # 폴더가 없으면 생성
        if not os.path.exists(temp_folder_path):
            os.makedirs(temp_folder_path)
            print(f"{temp_folder_path} 폴더를 생성했습니다.")

        else:
            print(f"{temp_folder_path}내부 파일을 빼겠습니다.")

            move_contents_to_parent(temp_folder_path)

        with open(file_path, 'r+') as lock:
            print(f"파일 {file_path} 잠금에 성공했습니다.")
            time.sleep(3)
            # 이동할 경로 계산
            destination = os.path.join(temp_folder_path, os.path.basename(file_path))

        # 파일 이동 (파일이 닫힌 후에)
        shutil.move(file_path, destination)
        print(f"파일을 {destination}로 이동했습니다.")


        return destination
    except Exception as e:
        print(f"파일 이동 중 오류 발생: {e}")
        return None




def get_video_files_list(infolder_path):
    """
    지정된 폴더 내의 비디오 파일 목록을 반환하는 함수

    Args:
        infolder_path (str): 검색할 폴더 경로

    Returns:
        list: 폴더 내의 비디오 파일 경로 목록
    """
    video_extensions = ('.mxf', '.mp4', '.mov', '.mts', '.m2t', '.mkv')
    video_files = []

    try:
        # 폴더 내 모든 파일 검색
        for file in os.listdir(infolder_path):
            # 파일 경로 생성
            file_path = os.path.join(infolder_path, file)

            # 파일인지 확인 (하위 폴더는 제외)
            if os.path.isfile(file_path):
                # 확장자 확인 (대소문자 무시)
                if file.lower().endswith(video_extensions):
                    video_files.append(file_path)

        return video_files

    except Exception as e:
        print(f"폴더 내 비디오 파일 검색 중 오류 발생: {e}")
        return []


def is_file_ready(file_path, check_duration=23, check_count=3):
    """
    파일이 더 이상 변경되지 않고 사용 준비가 되었는지 확인하는 함수

    Args:
        file_path (str): 확인할 파일 경로
        check_duration (int): 각 확인 간격(초) (기본값: 23초)
        check_count (int): 확인 횟수 (기본값: 3회)

    Returns:
        bool: 파일이 준비되었으면 True, 아니면 False
    """
    try:
        # 파일이 존재하는지 확인
        if not os.path.isfile(file_path):
            print(f"파일이 존재하지 않습니다: {file_path}")
            return False

        # MediaInfo를 통한 미디어 파일 검사 (비디오 파일인 경우)
        if any(file_path.lower().endswith(ext) for ext in ['.mxf', '.mov', '.mp4', '.mts', '.m2t', '.mkv']):
            try:
                from pymediainfo import MediaInfo
                media_info = MediaInfo.parse(file_path)

                # 비디오 트랙에 duration이 있는지 확인
                duration_exists = False
                for track in media_info.tracks:
                    if track.track_type == 'Video':
                        duration = getattr(track, 'duration', None)
                        if duration is not None:
                            duration_exists = True
                            break

                if not duration_exists:
                    print(f"파일 {file_path}의 비디오 트랙에 duration 정보가 없습니다. 아직 완성되지 않은 파일일 수 있습니다.")
                    return False
            except Exception as e:
                print(f"MediaInfo 처리 중 오류 발생: {e}")
                # MediaInfo 처리 실패 시에도 안정성 검사는 진행

        # 파일 크기와 수정 시간 안정성 검사
        initial_size = os.path.getsize(file_path)
        initial_mtime = os.path.getmtime(file_path)

        for i in range(check_count):
            print(f"파일 안정성 검사 중... ({i + 1}/{check_count})")
            time.sleep(check_duration)

            current_size = os.path.getsize(file_path)
            current_mtime = os.path.getmtime(file_path)

            if current_size != initial_size or current_mtime != initial_mtime:
                print(
                    f"파일이 여전히 변경 중입니다. (크기 변화: {initial_size} -> {current_size}, 수정 시간 변화: {initial_mtime} -> {current_mtime})")
                return False

        # 모든 검사 통과
        print(f"파일 {file_path}이(가) 준비되었습니다.")
        return True

    except Exception as e:
        print(f"파일 준비 상태 확인 중 오류 발생: {e}")
        return False

