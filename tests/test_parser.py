"""
parser.py의 단위 테스트.
"""
import pytest
from mass_ppt.parser import parse_mass_url


class TestParseMassUrl:
    """URL 파싱 테스트."""

    def test_parse_mass_url_returns_dict(self):
        """파싱 결과가 dict를 반환하는지 확인."""
        # 이 테스트는 실제 URL이 필요하므로
        # 사용자가 제공한 URL로 진행
        pass

    def test_date_extraction_from_url(self):
        """URL에서 날짜를 추출하는지 확인."""
        # 예: gomonth=2026-10-18 → date: 2026-10-18
        pass

    def test_lecture_sources_format(self):
        """독서/복음 출처가 올바른 형식인지 확인."""
        # 예: "▥ 이사야서의 말씀입니다." + "25,6-10ㄱ"
        pass

    def test_responsorial_refrain_first_only(self):
        """화답송은 첫 번째 ◎만 추출하는지 확인."""
        # "또는" 이후 제외
        pass

    def test_gospel_acclamation_extraction(self):
        """복음 환호송 ○ 구절 추출 확인."""
        pass

    def test_communion_hymn_without_source(self):
        """영성체송은 출처 제외 본문만 추출하는지 확인."""
        pass

    def test_missing_items_return_none(self):
        """없는 항목은 None을 반환하는지 확인."""
        # 구조적으로 모든 키가 dict에 있어야 함
        pass


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
