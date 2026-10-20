"""
매일미사 웹사이트에서 전례 정보를 파싱하는 모듈.
"""
import re
from typing import Optional
from urllib.parse import urlparse, parse_qs

import requests
from bs4 import BeautifulSoup


def parse_mass_url(url: str) -> dict:
    """
    매일미사 웹사이트 URL에서 전례 정보를 추출합니다.

    Args:
        url: 매일미사 웹사이트 URL

    Returns:
        dict: 전례 정보 (없는 항목은 None)
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
                      'AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/91.0.4472.124 Safari/537.36'
    }

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.encoding = 'utf-8'
    except requests.RequestException as e:
        raise Exception(f"URL 접근 실패: {e}")

    soup = BeautifulSoup(response.text, 'html.parser')

    # URL에서 날짜 추출
    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    gomonth = query_params.get('gomonth', [None])[0]

    data = {
        'date': gomonth,
        'liturgy_title': None,
        'first_reading_intro': None,
        'first_reading_ref': None,
        'responsorial_refrain': None,
        'second_reading_intro': None,
        'second_reading_ref': None,
        'gospel_acclamation': None,
        'gospel_intro': None,
        'gospel_ref': None,
        'communion_hymn': None,
    }

    # 페이지 전체 텍스트에서 정보 추출
    page_text = soup.get_text()

    # 전례일 제목 (예: "(녹) 연중 제29주일, ...")
    title_match = re.search(
        r'\([^\)]*\)\s*연중 제\d+주일[^\n]*',
        page_text
    )
    if title_match:
        data['liturgy_title'] = title_match.group(0).strip()

    # 모든 "▥" 항목 찾기 (독서 출처들)
    all_readings = re.findall(
        r'▥\s*([^\d]+?(?:입니다\.))\s*(\d+[,:\d\-ㄱㄴㄷ]+)',
        page_text,
        re.DOTALL
    )

    # 첫 번째 독서
    if len(all_readings) > 0:
        intro = all_readings[0][0].strip()
        ref = all_readings[0][1].strip()
        # 참조 번호 숫자 부분만 (첫 번째 숫자 그룹까지)
        ref_clean = re.match(r'(\d+[,\d\-ㄱㄴㄷ]*)', ref)
        if ref_clean:
            data['first_reading_intro'] = '▥ ' + intro
            data['first_reading_ref'] = ref_clean.group(1)

    # 두 번째 독서
    if len(all_readings) > 1:
        intro = all_readings[1][0].strip()
        ref = all_readings[1][1].strip()
        ref_clean = re.match(r'(\d+[,\d\-ㄱㄴㄷ]*)', ref)
        if ref_clean:
            data['second_reading_intro'] = '▥ ' + intro
            data['second_reading_ref'] = ref_clean.group(1)

    # 화답송 후렴 (첫 번째 ◎만, "또는" 제외)
    responsorial_matches = re.findall(
        r'◎\s*([^\n◎]*?)(?:또는|\n|$)',
        page_text
    )
    if responsorial_matches:
        refrain = responsorial_matches[0].strip()
        if refrain:
            data['responsorial_refrain'] = '◎ ' + refrain

    # 복음 환호송 (○ 구절, 첫 번째만)
    gospel_acclamation_match = re.search(
        r'○\s*([^\n◎]+)',
        page_text
    )
    if gospel_acclamation_match:
        text = gospel_acclamation_match.group(1).strip()
        data['gospel_acclamation'] = text

    # 복음 출처 ("✠"로 시작하는 항목)
    gospel_match = re.search(
        r'✠\s*([^\d]+?(?:입니다\.))\s*(\d+[,:\d\-ㄱㄴㄷ]+)',
        page_text,
        re.DOTALL
    )
    if gospel_match:
        intro = gospel_match.group(1).strip()
        ref = gospel_match.group(2).strip()
        ref_clean = re.match(r'(\d+[,\d\-ㄱㄴㄷ]*)', ref)
        if ref_clean:
            data['gospel_intro'] = '✠ ' + intro
            data['gospel_ref'] = ref_clean.group(1)

    # 영성체송 본문
    # "영성체송" 섹션에서 출처를 제외한 본문
    communion_match = re.search(
        r'영성체송\s*(?:[\s\S]*?)([^\n]+)',
        page_text
    )
    if communion_match:
        hymn_text = communion_match.group(1).strip()
        # 불릿, 번호 등 제외
        if not hymn_text.startswith(('참조', '마')):
            data['communion_hymn'] = hymn_text

    return data
