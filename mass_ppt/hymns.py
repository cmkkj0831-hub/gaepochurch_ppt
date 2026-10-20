"""
성가 이미지를 PPT 슬라이드에 삽입하는 모듈.
"""
import io
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Callable, Optional, Tuple
from lxml import etree


_NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def _find_hymn_slides(slide_elem) -> Optional[int]:
    """
    슬라이드가 성가 번호 슬라이드인지 확인하고, 번호를 반환.
    조건: 숫자만 적힌 텍스트박스 1개 + 그림 1개

    Returns:
        성가 번호 (int) 또는 None
    """
    namespaces = _NAMESPACES

    # 숫자만 있는 텍스트 찾기
    for shape in slide_elem.findall('.//p:sp', namespaces):
        text_body = shape.find('.//p:txBody', namespaces)
        if text_body is None:
            continue

        text = ''.join(t.text or '' for t in text_body.findall('.//a:t', namespaces))
        if text.strip().isdigit():
            return int(text.strip())

    return None


def _find_hymn_slide_groups(slides_xml: list) -> dict:
    """
    연속된 성가 슬라이드들을 그룹화.

    Returns:
        {
            'entrance': [(start_idx, count), ...],
            'offertory': [(start_idx, count), ...],
            'communion': [(start_idx, count), ...],
            'recessional': [(start_idx, count), ...]
        }
    """
    namespaces = _NAMESPACES
    hymn_groups = {
        'entrance': [],
        'offertory': [],
        'communion': [],
        'recessional': []
    }

    # 앵커 텍스트로 전례 구간 찾기
    entrance_start = None
    offertory_start = None
    communion_start = None
    recessional_start = None

    for idx, slide_bytes in enumerate(slides_xml):
        tree = etree.fromstring(slide_bytes)

        # 앵커 텍스트 검색
        all_text = ''.join(
            t.text or ''
            for t in tree.findall('.//a:t', namespaces)
        )

        if '예물 준비' in all_text:
            offertory_start = idx
        elif '영 성 체' in all_text and '기도' not in all_text:  # 영성체송 (기도문 제외)
            communion_start = idx

    # 성가 슬라이드 탐색
    i = 0
    while i < len(slides_xml):
        slide_bytes = slides_xml[i]
        tree = etree.fromstring(slide_bytes)
        hymn_num = _find_hymn_slides(tree)

        if hymn_num is not None:
            # 연속된 성가 슬라이드 묶음 세기
            start_idx = i
            count = 1
            j = i + 1
            while j < len(slides_xml):
                tree_next = etree.fromstring(slides_xml[j])
                hymn_next = _find_hymn_slides(tree_next)
                if hymn_next == hymn_num:
                    count += 1
                    j += 1
                else:
                    break

            # 위치 결정
            if entrance_start is None or start_idx < (offertory_start or float('inf')):
                hymn_groups['entrance'].append((start_idx, count, hymn_num))
                entrance_start = start_idx
            elif offertory_start and start_idx >= offertory_start and start_idx < (communion_start or float('inf')):
                hymn_groups['offertory'].append((start_idx, count, hymn_num))
                offertory_start = start_idx
            elif communion_start and start_idx >= communion_start:
                # 성체성가와 파견 구분 (영성체송 뒤로 2개 묶음)
                # 첫 번째: 성체, 두 번째: 파견
                if not hymn_groups['communion']:
                    hymn_groups['communion'].append((start_idx, count, hymn_num))
                else:
                    hymn_groups['recessional'].append((start_idx, count, hymn_num))

            i = j
        else:
            i += 1

    return hymn_groups


def _get_image_dimensions(image_bytes: bytes) -> Tuple[int, int]:
    """이미지 크기 (px)를 구함."""
    from PIL import Image
    img = Image.open(io.BytesIO(image_bytes))
    return img.size


def _calculate_image_position(image_bytes: bytes, content_width: int, content_height: int) -> Tuple[int, int, int, int]:
    """
    이미지를 content 영역에 맞게 배치.
    content_width: 오른쪽 여유 공간 (px)
    content_height: 아래쪽 여유 공간 (px)

    Returns:
        (x_offset, y_offset, width, height) - EMU 단위
    """
    img_w, img_h = _get_image_dimensions(image_bytes)

    # 비율 유지하며 최대 크기로 스케일링
    aspect_ratio = img_w / img_h
    if content_width / content_height > aspect_ratio:
        # 높이 기준으로 맞춤
        new_h = content_height
        new_w = int(content_height * aspect_ratio)
    else:
        # 너비 기준으로 맞춤
        new_w = content_width
        new_h = int(content_width / aspect_ratio)

    # 가운데 정렬
    x_offset = int((content_width - new_w) / 2)
    y_offset = int((content_height - new_h) / 2)

    # EMU로 변환 (DPI 96 기준: 914400 EMU/inch, 1 inch = 96px)
    # 72 DPI 기준: 914400 EMU/inch, 1 inch = 72px
    emu_per_px = 914400 / 96  # = 9525

    return (
        int(1567543 + x_offset * emu_per_px),  # x는 텍스트박스 끝 위치부터
        int(y_offset * emu_per_px),
        int(new_w * emu_per_px),
        int(new_h * emu_per_px)
    )


def replace_hymn_slides(
    ppt_bytes: bytes,
    hymn_numbers: dict,
    get_hymn_image: Callable[[int], Optional[bytes]]
) -> Tuple[bytes, dict]:
    """
    성가 슬라이드를 교체합니다.

    Args:
        ppt_bytes: 템플릿 PPT
        hymn_numbers: {'entrance': 58, 'offertory': 106, 'communion': [128, 129], 'recessional': 61}
        get_hymn_image: callable(hymn_number: int) -> bytes or None

    Returns:
        (result_bytes, report)
        report: {
            'inserted_hymns': [...],
            'warnings': [...]
        }
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # PPT 압축 해제
        extract_dir = tmpdir / 'pptx'
        with zipfile.ZipFile(io.BytesIO(ppt_bytes), 'r') as z:
            z.extractall(extract_dir)

        slides_dir = extract_dir / 'ppt' / 'slides'

        # 모든 슬라이드 XML 로드
        slide_files = sorted(slides_dir.glob('slide*.xml'))
        slides_xml = {}
        for slide_file in slide_files:
            slide_idx = int(slide_file.stem[5:])
            with open(slide_file, 'rb') as f:
                slides_xml[slide_idx] = f.read()

        # 성가 그룹 탐지
        slide_list = [slides_xml[i] for i in sorted(slides_xml.keys())]
        hymn_groups = _find_hymn_slide_groups(slide_list)

        # 성가 삽입
        report = {
            'inserted_hymns': [],
            'warnings': []
        }

        # 임시로 슬라이드 인덱스 관리
        # (너무 복잡하므로 단순화: 각 위치에서 기존 슬라이드 삭제 후 새로 삽입)

        result_bytes = ppt_bytes  # 일단 원본 반환 (2단계에서 구현)

        return result_bytes, report
