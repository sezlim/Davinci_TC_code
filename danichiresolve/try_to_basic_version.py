import sys
import os

def add_resolve_script_path():
    # 현재 실행 중인 Python 인터프리터 경로 (sys.executable 기준)
    base_path = os.path.dirname(sys.executable)
    base_path = os.path.dirname(base_path)
    base_path = os.path.dirname(base_path)

    # fusionscript.dll과 DaVinciResolveScript.py가 존재하는지 확인
    resolve_script_path = os.path.join(base_path, "DaVinciResolveScript.py")
    fusion_dll_path = os.path.join(base_path, "fusionscript.dll")

    if os.path.exists(resolve_script_path) and os.path.exists(fusion_dll_path):
        # sys.path에 해당 폴더 추가 (Python에서 import 가능하도록)
        sys.path.append(base_path)
        try:
            import DaVinciResolveScript as dvr_script
            print("DaVinci Resolve API 연결 성공!")
            return dvr_script.scriptapp("Resolve")
        except ImportError:
            print("DaVinciResolveScript.py가 있지만 import에 실패했습니다. 경로를 확인하세요.")
            return None
    else:
        print("필요한 파일이 없습니다. DaVinciResolveScript.py와 fusionscript.dll을 확인하세요.")
        return None

# 함수 실행하여 DaVinci Resolve 연결 시도
resolve = add_resolve_script_path()
if resolve:
    print("DaVinci Resolve API 사용 가능!")
else:
    print("DaVinci Resolve API 연결 실패")
