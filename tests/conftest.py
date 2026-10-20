import os
from pathlib import Path

import pytest
from lxml import etree

from mass_ppt.parser import parse_html
from mass_ppt.pptx_package import Package, qn
from mass_ppt.storage import LocalStorage

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).resolve().parent / 'fixtures'
# 최종본 PPT는 저장소에 없으므로(Releases 보관) 로컬 폴더를 지정해서 쓴다.
ARCHIVE_DIR = Path(os.environ.get('MASS_PPT_ARCHIVE_DIR', ROOT / 'archive'))


def _archive(name: str) -> bytes:
    path = ARCHIVE_DIR / name
    if not path.exists():
        pytest.skip(f'{path} 없음 (MASS_PPT_ARCHIVE_DIR로 최종본 폴더 지정)')
    return path.read_bytes()


@pytest.fixture(scope='session')
def ppt_1010() -> bytes:
    return _archive('261010미사.pptx')


@pytest.fixture(scope='session')
def ppt_1017() -> bytes:
    return _archive('261017미사.pptx')


@pytest.fixture(scope='session')
def fields_1018() -> dict:
    return parse_html((FIXTURES / 'missa_2026-10-18.html').read_text(encoding='utf-8'))


@pytest.fixture(scope='session')
def get_images():
    return LocalStorage(ARCHIVE_DIR, ROOT / 'songs').get_hymn_images


def run_table(ppt_bytes: bytes) -> dict:
    """{(슬라이드 번호, 문단/런 위치): (텍스트, rPr 직렬화)} — 슬라이드 순서 기준."""
    pkg = Package(ppt_bytes)
    out = {}
    for n, part in enumerate(pkg.slide_parts(), 1):
        root = pkg.xml(part)
        for i, r in enumerate(root.iter(qn('a:r'))):
            t = r.find(qn('a:t'))
            rpr = r.find(qn('a:rPr'))
            out[(n, i)] = ((t.text or '') if t is not None else '',
                           etree.tostring(rpr) if rpr is not None else b'')
    return out
