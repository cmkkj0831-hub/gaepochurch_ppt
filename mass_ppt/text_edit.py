"""
PPT 슬라이드의 텍스트를 수정하는 모듈.
서식(formatting) 보존이 최우선.
"""
import io
import re
import tempfile
import zipfile
from datetime import datetime, timedelta
from typing import Optional
from pathlib import Path

from lxml import etree


# XML 네임스페이스
_NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def _format_date_for_title(gomonth_str: Optional[str]) -> str:
    """
    URL의 gomonth (YYYY-MM-DD 형식의 주일)에서 토요일 날짜를 "YYYY.MM.DD" 형식으로 변환.
    """
    if not gomonth_str:
        return ""

    try:
        date_obj = datetime.strptime(gomonth_str, '%Y-%m-%d')
        # 주일에서 토요일로 (하루 전)
        saturday = date_obj - timedelta(days=1)
        return saturday.strftime('%Y.%m.%d')
    except (ValueError, TypeError):
        return ""


def _find_slide_by_anchor(slide_elem, anchor_keywords):
    """
    슬라이드 내에서 앵커 텍스트를 포함한 첫 번째 텍스트 프레임을 찾음.

    Args:
        slide_elem: slide XML 요소
        anchor_keywords: 검색할 텍스트 리스트 (OR 조건)

    Returns:
        (shape_elem, text_frame_elem) 튜플, 없으면 (None, None)
    """
    namespaces = _NAMESPACES

    # 슬라이드의 모든 도형 탐색
    for shape in slide_elem.findall('.//p:sp', namespaces):
        text_body = shape.find('.//p:txBody', namespaces)
        if text_body is None:
            continue

        # 이 도형의 모든 텍스트 수집
        all_text = ''.join(
            t.text or ''
            for t in text_body.findall('.//a:t', namespaces)
        )

        # 앵커 키워드 매칭
        for keyword in anchor_keywords:
            if keyword in all_text:
                return shape, text_body

    return None, None


def _update_run_text(run_elem, new_text):
    """
    run 요소의 텍스트를 수정 (서식 보존).

    Args:
        run_elem: <a:r> 요소
        new_text: 새로운 텍스트
    """
    namespaces = _NAMESPACES
    t_elem = run_elem.find('a:t', namespaces)

    if t_elem is not None:
        t_elem.text = new_text
    else:
        # <a:t> 요소가 없으면 생성
        t_elem = etree.SubElement(run_elem, '{%s}t' % namespaces['a'])
        t_elem.text = new_text


def _update_lecture_source(text_body, intro_text, ref_text):
    """
    독서/복음의 출처와 장절을 수정.

    독서 형식:
    - 단락 0: "▥ 이사야서의 말씀입니다." + "25,6-10ㄱ" (두 개의 run)

    Args:
        text_body: <p:txBody> 요소
        intro_text: 출처 소개 (예: "▥ 이사야서의 말씀입니다.")
        ref_text: 장절 (예: "25,6-10ㄱ")
    """
    namespaces = _NAMESPACES
    paragraphs = text_body.findall('.//a:p', namespaces)

    if not paragraphs:
        return

    first_para = paragraphs[0]
    runs = first_para.findall('.//a:r', namespaces)

    if not runs:
        return

    # 첫 번째 run: 소개 텍스트
    _update_run_text(runs[0], intro_text)

    # 두 번째 run: 장절
    if len(runs) > 1:
        _update_run_text(runs[1], ref_text)
    else:
        # 두 번째 run이 없으면 생성
        new_run = etree.Element('{%s}r' % namespaces['a'])
        # 첫 번째 run의 서식 복사
        rpr = runs[0].find('a:rPr', namespaces)
        if rpr is not None:
            import copy
            rpr_copy = copy.deepcopy(rpr)
            new_run.append(rpr_copy)

        t = etree.SubElement(new_run, '{%s}t' % namespaces['a'])
        t.text = ref_text
        first_para.append(new_run)


def _update_single_paragraph(text_body, new_text, para_index=0):
    """
    특정 단락의 텍스트를 수정 (모든 run의 텍스트를 하나로 통합).

    Args:
        text_body: <p:txBody> 요소
        new_text: 새로운 텍스트
        para_index: 수정할 단락 인덱스
    """
    namespaces = _NAMESPACES
    paragraphs = text_body.findall('.//a:p', namespaces)

    if para_index >= len(paragraphs):
        return

    para = paragraphs[para_index]
    runs = para.findall('.//a:r', namespaces)

    if not runs:
        return

    # 첫 번째 run에 텍스트 할당
    _update_run_text(runs[0], new_text)

    # 나머지 run은 빈 텍스트로 설정
    for run in runs[1:]:
        _update_run_text(run, '')


def update_ppt(ppt_bytes: bytes, data: dict) -> bytes:
    """
    PPT를 bytes로 받아 수정 후 bytes로 반환.

    Args:
        ppt_bytes: PPT 파일의 bytes
        data: 파싱된 전례 정보 (parser.parse_mass_url의 반환값)

    Returns:
        수정된 PPT의 bytes
    """
    # 임시 폴더에서 작업
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        ppt_path = tmpdir / 'temp.pptx'

        # bytes를 파일로 저장
        with open(ppt_path, 'wb') as f:
            f.write(ppt_bytes)

        # PPTX는 ZIP 형식이므로 압축 해제
        extract_dir = tmpdir / 'extracted'
        with zipfile.ZipFile(ppt_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # slide1.xml 수정 (표지)
        slide1_path = extract_dir / 'ppt' / 'slides' / 'slide1.xml'
        if slide1_path.exists():
            _update_title_slide(slide1_path, data)

        # 나머지 슬라이드들 수정
        slides_dir = extract_dir / 'ppt' / 'slides'
        for slide_file in sorted(slides_dir.glob('slide*.xml')):
            if slide_file.name == 'slide1.xml':
                continue
            _update_content_slide(slide_file, data)

        # 수정된 파일을 다시 ZIP으로 압축
        output_ppt = tmpdir / 'output.pptx'
        with zipfile.ZipFile(output_ppt, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in extract_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(extract_dir)
                    zipf.write(file_path, arcname)

        # bytes로 읽기
        with open(output_ppt, 'rb') as f:
            result_bytes = f.read()

        return result_bytes


def _update_title_slide(slide1_path: Path, data: dict):
    """
    슬라이드 1 (표지) 수정.
    - 제목: 전례일 제목 (숫자 앞뒤 공백 추가)
    - 날짜: YYYY.MM.DD 형식
    """
    tree = etree.parse(str(slide1_path))
    root = tree.getroot()
    namespaces = _NAMESPACES

    # 전례일 제목 포맷팅
    liturgy_title = data.get('liturgy_title')
    if liturgy_title:
        # 숫자 앞뒤 공백 추가 (예: "연중 제29주일" → "연중 제 29 주일")
        formatted_title = re.sub(
            r'(\D)(\d+)(\D)',
            r'\1 \2 \3',
            liturgy_title
        )

        # 괄호 부분이 있으면 "연중 제 29 주일(전교 주일)" 형식으로
        # 기존 형식 유지하되 숫자 주변 공백만 추가
        if '(' in formatted_title:
            parts = formatted_title.split('(', 1)
            main_part = parts[0].strip()
            main_part = re.sub(r'제\s*(\d+)\s*주일', r'제 \1 주일', main_part)
            formatted_title = f"{main_part}({parts[1]}"

        shape, text_body = _find_slide_by_anchor(
            root,
            ['연중 제', '주일']
        )
        if shape is not None and text_body is not None:
            _update_single_paragraph(text_body, formatted_title)

    # 날짜 수정
    date_str = _format_date_for_title(data.get('date'))
    if date_str:
        shape, text_body = _find_slide_by_anchor(root, [date_str.split('.')[2]])  # 일자로 검색
        if shape is not None and text_body is not None:
            # 모든 단락을 검색해서 날짜 부분 찾기
            for para_idx in range(len(text_body.findall('.//a:p', namespaces))):
                para = text_body.findall('.//a:p', namespaces)[para_idx]
                para_text = ''.join(
                    t.text or ''
                    for t in para.findall('.//a:t', namespaces)
                )
                if '202' in para_text or '20' in para_text:  # 날짜처럼 보이는 텍스트
                    _update_single_paragraph(text_body, date_str, para_idx)
                    break

    tree.write(str(slide1_path), encoding='utf-8', xml_declaration=True)


def _update_content_slide(slide_path: Path, data: dict):
    """
    슬라이드 2-62 수정 (제1독서, 화답송, 제2독서, 복음, 영성체송 등).
    """
    tree = etree.parse(str(slide_path))
    root = tree.getroot()
    namespaces = _NAMESPACES

    # 제1독서 (슬라이드 27)
    if data.get('first_reading_intro') and data.get('first_reading_ref'):
        shape, text_body = _find_slide_by_anchor(root, ['제1독서', '제 1 독서'])
        if shape is not None and text_body is not None:
            _update_lecture_source(
                text_body,
                data['first_reading_intro'],
                data['first_reading_ref']
            )

    # 화답송 (슬라이드 28)
    if data.get('responsorial_refrain'):
        shape, text_body = _find_slide_by_anchor(root, ['화답송'])
        if shape is not None and text_body is not None:
            # 화답송은 특정 단락 (후렴 위치) 수정
            _update_single_paragraph(text_body, data['responsorial_refrain'])

    # 제2독서 (슬라이드 29)
    if data.get('second_reading_intro') and data.get('second_reading_ref'):
        shape, text_body = _find_slide_by_anchor(root, ['제2독서', '제 2 독서'])
        if shape is not None and text_body is not None:
            _update_lecture_source(
                text_body,
                data['second_reading_intro'],
                data['second_reading_ref']
            )

    # 복음 환호송 (슬라이드 30)
    if data.get('gospel_acclamation'):
        shape, text_body = _find_slide_by_anchor(root, ['복 음 환 호 송', '복음환호송'])
        if shape is not None and text_body is not None:
            # 첫 번째 ◎ 찾아서 수정
            _update_single_paragraph(text_body, '◎ ' + data['gospel_acclamation'])

    # 복음 (슬라이드 31)
    if data.get('gospel_intro') and data.get('gospel_ref'):
        shape, text_body = _find_slide_by_anchor(root, ['복음'])
        if shape is not None and text_body is not None:
            # 복음은 첫 단락에 출처와 장절
            _update_lecture_source(
                text_body,
                data['gospel_intro'],
                data['gospel_ref']
            )

    # 영성체송 (슬라이드 56)
    if data.get('communion_hymn'):
        shape, text_body = _find_slide_by_anchor(root, ['영 성 체 송', '영성체송'])
        if shape is not None and text_body is not None:
            _update_single_paragraph(text_body, data['communion_hymn'])

    tree.write(str(slide_path), encoding='utf-8', xml_declaration=True)
