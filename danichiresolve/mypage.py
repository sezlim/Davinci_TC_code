import os
import sys
import ctypes


def find_fusionscript_dll():
    # 현재 실행 중인 Python 인터프리터의 경로 가져오기
    python_path = sys.executable

    # 인터프리터 경로에서 한 폴더 위로 올라가기
    parent_dir = os.path.dirname(os.path.dirname(os.path.dirname(python_path)))

    print(f"시작 검색 경로: {parent_dir}")
    found_paths = []

    # 하위 폴더를 모두 검색
    for root, dirs, files in os.walk(parent_dir):
        for file in files:
            if file.lower() == "fusionscript.dll":
                full_path = os.path.join(root, file)
                found_paths.append(full_path)
                print(f"찾음: {full_path}")

    if not found_paths:
        print("fusionscript.dll 파일을 찾을 수 없습니다.")
        # DaVinci Resolve의 일반적인 설치 경로 시도
        default_path = "C:\\Program Files\\Blackmagic Design\\DaVinci Resolve\\fusionscript.dll"
        if os.path.exists(default_path):
            print(f"기본 경로에서 찾음: {default_path}")
            return default_path
        return None

    return found_paths[0]


def get_resolve_direct(path):
    if not path or not os.path.exists(path):
        print(f"유효하지 않은 DLL 경로: {path}")
        return None

    # DLL 파일 경로 지정
    dll_path = os.path.normpath(path)
    print(f"DLL 로드 시도: {dll_path}")

    try:
        # 환경 변수 설정 (DLL 로드에 도움이 될 수 있음)
        os.environ["RESOLVE_SCRIPT_API"] = dll_path
        os.environ["RESOLVE_SCRIPT_LIB"] = dll_path

        # DLL 파일 로드
        fusion_lib = ctypes.CDLL(dll_path)
        print("DLL 로드 성공")


        try:
            # scriptapp 함수 가져오기 시도
            get_resolve_func = getattr(fusion_lib, "scriptapp")
            get_resolve_func.restype = ctypes.c_void_p
            get_resolve_func.argtypes = [ctypes.c_char_p]

            # Resolve 객체 얻기
            resolve_ptr = get_resolve_func(b"Resolve")
            print("Resolve 객체 포인터 얻음:", resolve_ptr)

            if resolve_ptr:
                print("DaVinci Resolve API 연결 성공")
                return resolve_ptr
            else:
                print("Resolve 객체를 가져올 수 없습니다")
                return None

        except AttributeError:
            print("'scriptapp' 함수를 찾을 수 없습니다. DLL 인터페이스가 예상과 다를 수 있습니다.")
            # 다른 함수명 시도 가능
            return None

    except Exception as e:
        print(f"DLL 로드 중 오류 발생: {e}")
        return None


# 함수 실행
dll_path = find_fusionscript_dll()
if dll_path:
    resolve = get_resolve_direct(dll_path)
    if resolve:
        print("DaVinci Resolve에 연결됨")
        # 여기서부터 Resolve API 사용 가능
        # 하지만 C 포인터를 Python 객체로 어떻게 다룰지는 별도 처리 필요
    else:
        print("DaVinci Resolve 연결 실패")
else:
    print("fusionscript.dll을 찾을 수 없습니다")