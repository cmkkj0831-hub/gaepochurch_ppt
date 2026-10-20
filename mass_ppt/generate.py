"""화면에서 확인·수정한 값과 성가 번호로 PPT를 만든다."""
from typing import Callable, List, Tuple

from .hymns import replace_hymns
from .text_edit import apply_text_fields


def generate(
    base_ppt_bytes: bytes,
    fields: dict,
    hymn_numbers: dict,
    get_hymn_images: Callable[[int], List[bytes]],
) -> Tuple[bytes, dict]:
    """
    fields: parser.FIELD_KEYS 형식 (사용자가 수정한 값, None 허용)
    hymn_numbers: {'entrance': 58, 'offertory': 106, 'communion': [128, 129], 'recessional': 61}
    get_hymn_images: 성가 번호 -> 이미지 bytes 목록 (쪽 순서)
    HymnTemplateError: 기준 PPT의 성가 묶음 구조가 맞지 않을 때
    """
    after_text, text_report = apply_text_fields(base_ppt_bytes, fields or {})
    result, hymn_report = replace_hymns(after_text, hymn_numbers or {}, get_hymn_images)
    report = {
        'changed_texts': text_report.get('changed') or [],
        'hymns': hymn_report.get('hymns') or [],
        'slide_count': hymn_report.get('slide_count'),
        'warnings': (text_report.get('warnings') or []) + (hymn_report.get('warnings') or []),
    }
    return result, report
