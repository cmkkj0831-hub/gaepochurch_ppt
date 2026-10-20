"""매일미사 페이지에서 PPT에 필요한 전례 정보를 섹션 단위로 추출한다."""
import re
from datetime import date, timedelta
from typing import Optional

import requests
from bs4 import BeautifulSoup

USER_AGENT = (
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
    '(KHTML, like Gecko) Chrome/124.0 Safari/537.36'
)

FIELD_KEYS = [
    'cover_title', 'cover_date', 'liturgy_title',
    'first_reading_intro', 'first_reading_ref',
    'psalm_refrain',
    'second_reading_intro', 'second_reading_ref',
    'gospel_acclamation',
    'gospel_intro', 'gospel_ref',
    'communion_antiphon',
]


class ParseError(Exception):
    pass


def fetch_html(url: str) -> str:
    try:
        resp = requests.get(url, headers={'User-Agent': USER_AGENT}, timeout=15)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise ParseError(f'매일미사 페이지를 불러오지 못했습니다: {e}') from e
    resp.encoding = 'utf-8'
    return resp.text


def parse_mass_url(url: str) -> dict:
    return parse_html(fetch_html(url))


def make_cover_title(raw: Optional[str]) -> Optional[str]:
    """'(녹) 연중 제29주일, …미사(전교 주일)' -> '연중 제 29 주일(전교 주일)'"""
    if not raw:
        return None
    title = raw.strip().strip('[]').strip()
    title = re.sub(r'^\(\s*[^()]{1,3}\s*\)\s*', '', title)
    if ',' in title:
        head = title.split(',', 1)[0].strip()
        tail = re.search(r'(\([^()]*\))\s*$', title)
        if tail and tail.group(1) not in head:
            head += tail.group(1)
        title = head
    title = re.sub(r'제\s*(\d+)\s*주일', r'제 \1 주일', title)
    return title.strip()


def _section_lines(soup: BeautifulSoup) -> dict:
    form = soup.find('form', attrs={'name': 'contents'})
    if form is None:
        return {}
    sections = {}
    for h3 in form.find_all('h3'):
        name = h3.get_text(strip=True)
        lines = []
        for sib in h3.next_siblings:
            if getattr(sib, 'name', None) == 'h3':
                break
            text = sib.get_text('\n') if hasattr(sib, 'get_text') else str(sib)
            lines.extend(ln.strip() for ln in text.split('\n'))
        lines = [ln for ln in lines if ln]
        if name not in sections:
            sections[name] = lines
    return sections


def _source_and_ref(lines: Optional[list], pattern: str):
    if not lines:
        return None, None
    for i, line in enumerate(lines):
        m = re.match(pattern, line)
        if m:
            ref = line[m.end():].strip()
            if not ref and i + 1 < len(lines) and re.match(r'\d', lines[i + 1]):
                ref = lines[i + 1].strip()
            return m.group(1).strip(), (ref or None)
    return None, None


def _first_starting(lines: Optional[list], mark: str) -> Optional[str]:
    for line in lines or []:
        if line.startswith(mark):
            return line.split('또는')[0].strip()
    return None


def parse_html(html: str) -> dict:
    soup = BeautifulSoup(html, 'html.parser')
    fields = dict.fromkeys(FIELD_KEYS)

    today = soup.find('p', class_='today')
    if today:
        m = re.search(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', today.get_text())
        if m:
            sunday = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            fields['cover_date'] = (sunday - timedelta(days=1)).strftime('%Y.%m.%d')

    title_el = soup.find('p', class_='m_t_today01')
    if title_el:
        raw = title_el.get_text(strip=True).strip('[]').strip()
        fields['liturgy_title'] = raw or None
        fields['cover_title'] = make_cover_title(raw)

    sec = _section_lines(soup)

    fields['first_reading_intro'], fields['first_reading_ref'] = _source_and_ref(
        sec.get('제1독서'), r'(▥.*?말씀입니다\.)')
    fields['second_reading_intro'], fields['second_reading_ref'] = _source_and_ref(
        sec.get('제2독서'), r'(▥.*?말씀입니다\.)')
    fields['gospel_intro'], fields['gospel_ref'] = _source_and_ref(
        sec.get('복음'), r'(✠.*?복음입니다\.)')

    fields['psalm_refrain'] = _first_starting(sec.get('화답송'), '◎')
    fields['gospel_acclamation'] = _first_starting(sec.get('복음 환호송'), '○')

    comm = sec.get('영성체송')
    if comm and len(comm) > 1:
        body = []
        for line in comm[1:]:
            if re.fullmatch(r'<?\s*또는\s*>?', line):
                break
            body.append(line)
        fields['communion_antiphon'] = ' '.join(body).strip() or None

    return fields
