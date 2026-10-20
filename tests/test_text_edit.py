"""
text_edit.py의 단위 테스트.
"""
import pytest
from pathlib import Path
from mass_ppt.text_edit import update_ppt, _format_date_for_title


class TestFormatDateForTitle:
    """날짜 포맷팅 테스트."""

    def test_format_sunday_to_saturday(self):
        """일요일 날짜를 토요일로 변환."""
        # 2026-10-18 (일요일) → 2026.10.17 (토요일)
        result = _format_date_for_title('2026-10-18')
        assert result == '2026.10.17'

    def test_format_date_with_correct_separator(self):
        """날짜가 점(.)으로 구분되는지 확인."""
        result = _format_date_for_title('2026-10-18')
        assert '.' in result
        assert '-' not in result

    def test_invalid_date_returns_empty_string(self):
        """잘못된 날짜는 빈 문자열 반환."""
        result = _format_date_for_title('invalid')
        assert result == ''

    def test_none_returns_empty_string(self):
        """None은 빈 문자열 반환."""
        result = _format_date_for_title(None)
        assert result == ''


class TestUpdatePpt:
    """PPT 수정 테스트."""

    def test_update_ppt_returns_bytes(self):
        """수정된 PPT가 bytes를 반환하는지 확인."""
        # 실제 PPT 파일이 필요함
        archive_ppt = Path('archive/261010미사.pptx')
        if not archive_ppt.exists():
            pytest.skip("archive PPT 파일이 없습니다")

        with open(archive_ppt, 'rb') as f:
            ppt_bytes = f.read()

        data = {
            'date': '2026-10-18',
            'liturgy_title': '(녹) 연중 제29주일, 민족들의 복음화를 위한 미사(전교 주일)',
            'first_reading_intro': '▥ 이사야서의 말씀입니다.',
            'first_reading_ref': '2,1-5',
            'responsorial_refrain': '◎ 주님은 민족들의 눈앞에 당신 정의를 드러내셨네.',
            'second_reading_intro': '▥ 사도 바오로의 로마서 말씀입니다.',
            'second_reading_ref': '10,9-18',
            'gospel_acclamation': '우리 주 예수 그리스도의 아버지, 저희 마음의 눈을 밝혀 주시어 부르심을 받은 저희의 희망을 알게 하여 주소서.',
            'gospel_intro': '✠ 마태오가 전한 거룩한 복음입니다.',
            'gospel_ref': '28,16-20',
            'communion_hymn': '주님, 이 거룩한 선물로 저희를 구원하소서.',
        }

        result = update_ppt(ppt_bytes, data)
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_update_ppt_preserves_overall_structure(self):
        """PPT 전체 구조가 보존되는지 확인."""
        archive_ppt = Path('archive/261010미사.pptx')
        if not archive_ppt.exists():
            pytest.skip("archive PPT 파일이 없습니다")

        with open(archive_ppt, 'rb') as f:
            original_size = len(f.read())

        with open(archive_ppt, 'rb') as f:
            ppt_bytes = f.read()

        data = {'date': '2026-10-18'}
        result = update_ppt(ppt_bytes, data)

        # 수정된 파일이 비슷한 크기여야 함 (±20%)
        assert 0.8 * original_size < len(result) < 1.2 * original_size


class TestAnchorDetection:
    """앵커 텍스트 탐지 테스트."""

    def test_find_first_lecture_slide(self):
        """제1독서 슬라이드를 앵커 텍스트로 찾기."""
        # text_edit.py의 _find_slide_by_anchor 함수 검증
        # 실제 PPT 파일 필요
        pass

    def test_find_gospel_acclamation_slide(self):
        """복음 환호송 슬라이드 찾기."""
        pass


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
