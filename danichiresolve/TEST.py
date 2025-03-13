import xml.etree.ElementTree as ET
import re


def update_xml_values(original_xml_path, new_xml_path, output_path=None):
    """
    원본 XML의 구조를 유지하면서 새 XML에서 동일한 경로의 태그 값만 업데이트하는 함수
    또한 <DbKey>abs</DbKey> 태그의 값을 새 XML 파일 이름(확장자 제외)으로 변경함

    Args:
        original_xml_path (str): 원본 XML 파일 경로
        new_xml_path (str): 새 값을 가져올 XML 파일 경로
        output_path (str, optional): 결과를 저장할 경로. None이면 화면에 출력만 함

    Returns:
        str: 업데이트된 XML 문자열
    """
    # 원본 XML 문자열 그대로 로드 (들여쓰기, 주석 등 보존)
    with open(original_xml_path, 'r', encoding='utf-8') as f:
        original_xml_str = f.read()

    # 파싱을 위해 ElementTree 로드
    original_tree = ET.parse(original_xml_path)
    original_root = original_tree.getroot()

    new_tree = ET.parse(new_xml_path)
    new_root = new_tree.getroot()

    # 새 XML에서 모든 요소와 텍스트 값 수집
    new_values = {}

    def collect_values(element, path=""):
        current_path = path + "/" + element.tag if path else element.tag

        # 텍스트 값이 있고 비어있지 않으면 저장
        if element.text and element.text.strip():
            new_values[current_path] = element.text

        # 속성도 수집 (필요시)
        for attr_name, attr_value in element.attrib.items():
            attr_path = f"{current_path}[@{attr_name}]"
            new_values[attr_path] = attr_value

        # 재귀적으로 자식 요소 처리
        for child in element:
            collect_values(child, current_path)

    collect_values(new_root)

    # 원본 XML에서 같은 경로의 요소 값 업데이트
    def update_element_values(element, path=""):
        current_path = path + "/" + element.tag if path else element.tag

        # 이 경로에 대한 새 값이 있으면 업데이트
        if current_path in new_values:
            # 정규식으로 원본 XML 문자열에서 해당 태그의 텍스트 값만 업데이트
            tag_pattern = re.compile(f"(<{element.tag}[^>]*>)(.*?)(</{element.tag}>)")

            # 현재 경로의 모든 요소를 찾아서 텍스트 값만 업데이트
            for match in tag_pattern.finditer(original_xml_str):
                open_tag, old_text, close_tag = match.groups()
                # 태그 내용만 교체 (들여쓰기 포함된 경우 고려)
                old_text_stripped = old_text.strip()
                if old_text_stripped and not old_text_stripped.startswith("<"):
                    # 내부에 자식 태그가 없는 경우에만 교체
                    indent = old_text[:old_text.find(old_text_stripped)]
                    replacement = f"{indent}{new_values[current_path]}{indent if indent and '\n' in indent else ''}"
                    original_xml_str = original_xml_str.replace(
                        f"{open_tag}{old_text}{close_tag}",
                        f"{open_tag}{replacement}{close_tag}"
                    )

        # 속성 업데이트 (필요시)
        for attr_name in element.attrib:
            attr_path = f"{current_path}[@{attr_name}]"
            if attr_path in new_values:
                # 속성 값 업데이트 로직 추가 (필요시)
                pass

        # 재귀적으로 자식 요소 처리
        for child in element:
            update_element_values(child, current_path)

    # 이 접근 방식은 간단한 태그에는 작동하지만 복잡한 XML 구조에서는 한계가 있음
    # 더 복잡한 구조에서는 lxml과 같은 라이브러리를 사용하는 것이 좋음

    # 원본 XML 파일에서 더 정확한 접근이 필요하다면 직접 텍스트 처리 방식 구현
    def update_xml_directly():
        updated_xml = original_xml_str

        # 새 XML에서 수집한 모든 값에 대해
        for xpath, new_value in new_values.items():
            # 마지막 태그 이름 추출
            tag_name = xpath.split('/')[-1]

            # 속성이 포함된 경우 태그 이름만 추출
            if '[' in tag_name:
                tag_name = tag_name.split('[')[0]

            # 해당 태그의 닫는/여는 태그 패턴 생성
            pattern = re.compile(f"<{tag_name}[^>]*>(.*?)</{tag_name}>", re.DOTALL)

            # 원본 XML에서 해당 패턴 찾기
            matches = list(pattern.finditer(updated_xml))

            # XPath에 해당하는 인덱스 계산 (간단한 구현, 실제로는 더 복잡할 수 있음)
            path_parts = xpath.split('/')
            indices = []

            for i, part in enumerate(path_parts):
                if i > 0:  # 루트 요소 제외
                    # 같은 태그 이름의 몇 번째 요소인지 계산 (간단한 구현)
                    tag_only = part.split('[')[0] if '[' in part else part
                    count = 0

                    # 이전 경로까지의 모든 같은 이름 태그 수 세기
                    for j in range(len(matches)):
                        if f"<{tag_only}" in matches[j].group():
                            count += 1

                    indices.append(count - 1)  # 0-기반 인덱스

            # 올바른 매치 찾기 (간단한 구현)
            if matches:
                match_index = 0  # 간단한 경우 첫 번째 매치 사용

                if len(indices) > 0:
                    # 인덱스 사용하여 올바른 매치 찾기 (복잡한 경우)
                    # 이 부분은 실제 구현에서 더 정교해야 함
                    if indices[-1] < len(matches):
                        match_index = indices[-1]

                match = matches[match_index]

                # 태그 내용만 추출
                full_match = match.group(0)
                inner_content = match.group(1)

                # 내부에 하위 태그가 없는 경우에만 교체
                if not re.search(r"<[^>]+>", inner_content):
                    # 들여쓰기 유지
                    indent = ""
                    if inner_content and inner_content[0].isspace():
                        indent_match = re.match(r"(\s+)", inner_content)
                        if indent_match:
                            indent = indent_match.group(1)

                    # 새 값으로 교체하되 들여쓰기 유지
                    new_content = f"{indent}{new_value}{indent if indent and '\n' in indent else ''}"
                    new_full_match = full_match.replace(inner_content, new_content)

                    # 원본 XML 업데이트
                    updated_xml = updated_xml.replace(full_match, new_full_match)

        return updated_xml

    # 직접 텍스트 처리 방식 사용
    updated_xml = update_xml_directly()

    # <PresetList> 안에 있는 첫 번째 <DbKey> 태그의 값을 새 XML 파일 이름(확장자 제외)으로 변경
    import os
    new_xml_filename = os.path.splitext(os.path.basename(new_xml_path))[0]

    # 정확한 위치의 <DbKey> 태그를 찾기 위한 정규식
    # <PresetList> 안에 <Element> 안에 <DbKey> 패턴을 찾음
    preset_pattern = re.compile(r'<PresetList>(.*?)</PresetList>', re.DOTALL)
    preset_match = preset_pattern.search(updated_xml)

    if preset_match:
        preset_content = preset_match.group(1)
        element_pattern = re.compile(r'<Element>(.*?)</Element>', re.DOTALL)
        element_match = element_pattern.search(preset_content)

        if element_match:
            element_content = element_match.group(1)
            dbkey_pattern = re.compile(r'(<DbKey>)(.*?)(</DbKey>)', re.DOTALL)
            dbkey_match = dbkey_pattern.search(element_content)

            if dbkey_match:
                # 전체 문자열에서의 위치 계산
                full_start = preset_match.start() + preset_content.find(element_content) + element_content.find(
                    dbkey_match.group(0))
                full_end = full_start + len(dbkey_match.group(0))

                # <DbKey> 태그 변경
                updated_xml = updated_xml[:full_start] + f"<DbKey>{new_xml_filename}</DbKey>" + updated_xml[full_end:]

    # 결과 저장 또는 반환
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(updated_xml)

    return updated_xml


# 사용 예:
# updated_xml = update_xml_values('original.xml', 'new.xml', 'output.xml')
# print(updated_xml)

updated_xml = update_xml_values('DeliverPresetList.xml', 'xdcam_50dd.xml', 'output.xml')
print(updated_xml)