"""
매일미사 URL과 성가 정보로부터 완성된 PPT를 생성합니다.
"""
from typing import Callable, Optional, Tuple
from pathlib import Path

from .parser import parse_mass_url
from .text_edit import update_ppt
from .hymns import replace_hymn_slides


def generate(
    base_ppt_bytes: bytes,
    mass_url: str,
    hymn_numbers: dict,
    get_hymn_image: Callable[[int], Optional[bytes]]
) -> Tuple[bytes, dict]:
    """
    PPT를 생성합니다.

    Args:
        base_ppt_bytes: 템플릿 PPT (bytes)
        mass_url: 매일미사 웹사이트 URL
        hymn_numbers: {
            'entrance': int,
            'offertory': int,
            'communion': [int, ...],  # 1~2개
            'recessional': int
        }
        get_hymn_image: callable(hymn_number: int) -> bytes or None
                        성가 번호 → 이미지 bytes (없으면 None)

    Returns:
        (result_bytes, report)
        report: {
            'updated_texts': [...],
            'inserted_hymns': [...],
            'warnings': [...]
        }
    """
    # 1단계: 매일미사 URL 파싱
    parsed = parse_mass_url(mass_url)

    # 2단계: 텍스트 수정
    result = update_ppt(base_ppt_bytes, parsed)

    # 3단계: 성가 삽입 (준비 중)
    result, hymn_report = replace_hymn_slides(result, hymn_numbers, get_hymn_image)

    # 보고서 작성
    report = {
        'updated_texts': [
            f"제1독서: {parsed.get('first_reading_ref')}",
            f"화답송: {parsed.get('responsorial_refrain')[:30]}...",
            f"제2독서: {parsed.get('second_reading_ref')}",
            f"복음: {parsed.get('gospel_ref')}",
        ],
        'inserted_hymns': hymn_report.get('inserted_hymns', []),
        'warnings': hymn_report.get('warnings', [])
    }

    return result, report


def generate_from_local_hymns(
    base_ppt_bytes: bytes,
    mass_url: str,
    hymn_numbers: dict,
    songs_dir: Path = None
) -> Tuple[bytes, dict]:
    """
    로컬 파일에서 성가 이미지를 읽어 PPT를 생성합니다.

    Args:
        base_ppt_bytes: 템플릿 PPT
        mass_url: 매일미사 URL
        hymn_numbers: 성가 번호들
        songs_dir: 성가 이미지 디렉토리 (기본값: ./songs/)

    Returns:
        (result_bytes, report)
    """
    if songs_dir is None:
        songs_dir = Path(__file__).parent.parent / 'songs'

    def _get_hymn_image(hymn_num: int) -> Optional[bytes]:
        # 여러 장인 경우 첫 번째 이미지부터 순서대로 찾기
        for ext in ['.png', '.jpg']:
            # 한 장짜리
            path = songs_dir / f'{hymn_num}{ext}'
            if path.exists():
                return path.read_bytes()
            # 여러 장 (첫 번째)
            path = songs_dir / f'{hymn_num}-1{ext}'
            if path.exists():
                return path.read_bytes()

        return None

    return generate(base_ppt_bytes, mass_url, hymn_numbers, _get_hymn_image)
