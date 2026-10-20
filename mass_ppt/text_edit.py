"""PPT 텍스트를 run 단위로만 교체한다. 제목 칸·서식(rPr/pPr)은 건드리지 않는다."""
import re

from .pptx_package import NS, Package, compact, qn, shape_text, slide_title

TEXT_SLIDE_TITLES = ('제1독서', '화답송', '제2독서', '복음환호송', '복음', '영성체송')

FIELD_LABELS = {
    'cover_title': '표지 제목',
    'cover_date': '표지 날짜',
    'first_reading_intro': '제1독서 출처',
    'first_reading_ref': '제1독서 장절',
    'psalm_refrain': '화답송',
    'second_reading_intro': '제2독서 출처',
    'second_reading_ref': '제2독서 장절',
    'gospel_acclamation': '복음 환호송',
    'gospel_intro': '복음 출처',
    'gospel_ref': '복음 장절',
    'communion_antiphon': '영성체송',
}

DATE_RE = re.compile(r'^\d{4}\.\d{2}\.\d{2}$')


def find_text_slides(pkg: Package) -> dict:
    """제목 칸 텍스트 -> 해당 슬라이드 번호(1부터) 목록."""
    found = {t: [] for t in TEXT_SLIDE_TITLES}
    for i, part in enumerate(pkg.slide_parts(), 1):
        t = slide_title(pkg.xml(part), TEXT_SLIDE_TITLES)
        if t:
            found[t].append(i)
    return found


def _run_text(r) -> str:
    t = r.find('a:t', NS)
    return (t.text or '') if t is not None else ''


def _runs(el) -> list:
    return list(el.iter(qn('a:r')))


def _next_run_in_paragraph(r):
    para = r.getparent()
    runs = para.findall('a:r', NS)
    idx = runs.index(r)
    return runs[idx + 1] if idx + 1 < len(runs) else None


class _Editor:
    def __init__(self, fields: dict, report: dict):
        self.fields = fields
        self.report = report

    def value(self, key):
        v = self.fields.get(key)
        if v is None or not str(v).strip():
            self.report['warnings'].append(f'{FIELD_LABELS[key]}: 값이 없어 수정하지 않았습니다.')
            return None
        return str(v).strip()

    def set_run(self, r, key, value, slide_no):
        t = r.find('a:t', NS)
        old = t.text or ''
        lead = old[:len(old) - len(old.lstrip())]
        new = lead + value
        if new != old:
            t.text = new
            self.report['changed'].append(
                {'slide': slide_no, 'field': FIELD_LABELS[key], 'before': old, 'after': new})

    def clear_runs(self, runs):
        for r in runs:
            t = r.find('a:t', NS)
            if t is not None and t.text:
                t.text = ''

    def warn(self, msg):
        self.report['warnings'].append(msg)


def _body_shapes(root, exclude):
    out = []
    for sp in root.iter(qn('p:sp')):
        text = compact(shape_text(sp))
        if text and text not in exclude:
            out.append(sp)
    return out


def _edit_cover(ed: _Editor, root, slide_no):
    title_val, date_val = ed.value('cover_title'), ed.value('cover_date')
    date_runs = [r for r in _runs(root) if DATE_RE.match(_run_text(r).strip())]
    if date_val is not None:
        if len(date_runs) == 1:
            ed.set_run(date_runs[0], 'cover_date', date_val, slide_no)
        else:
            ed.warn(f'표지: 날짜 칸을 {len(date_runs)}개 찾아서 수정하지 않았습니다.')
    if title_val is not None:
        cands = []
        for sp in root.iter(qn('p:sp')):
            text = shape_text(sp).strip()
            if not text or '개포동성당' in text or DATE_RE.match(text):
                continue
            cands.append(sp)
        if len(cands) != 1:
            ed.warn(f'표지: 제목 칸을 {len(cands)}개 찾아서 수정하지 않았습니다.')
            return
        runs = [r for r in _runs(cands[0]) if _run_text(r)]
        ed.set_run(runs[0], 'cover_title', title_val, slide_no)
        ed.clear_runs(runs[1:])


def _runs_starting(root, title, mark, contains=False):
    out = []
    for sp in _body_shapes(root, {title}):
        for r in _runs(sp):
            text = _run_text(r)
            if (mark in text) if contains else text.lstrip().startswith(mark):
                out.append(r)
    return out


def _edit_source(ed, root, slide_no, title, mark, intro_key, ref_key, contains=False):
    intro, ref = ed.value(intro_key), ed.value(ref_key)
    if intro is None and ref is None:
        return
    runs = _runs_starting(root, title, mark, contains)
    if len(runs) != 1:
        ed.warn(f'{title}: "{mark}" 출처 칸을 {len(runs)}개 찾아서 수정하지 않았습니다.')
        return
    if intro is not None:
        ed.set_run(runs[0], intro_key, intro, slide_no)
    if ref is not None:
        nxt = _next_run_in_paragraph(runs[0])
        if nxt is None or not re.match(r'\d', _run_text(nxt).strip()):
            ed.warn(f'{title}: 장절 칸을 찾지 못해 장절은 수정하지 않았습니다.')
        else:
            ed.set_run(nxt, ref_key, ref, slide_no)


def _edit_single_mark(ed, root, slide_no, title, mark, key):
    val = ed.value(key)
    if val is None:
        return
    runs = _runs_starting(root, title, mark)
    if len(runs) != 1:
        ed.warn(f'{title}: "{mark}" 칸을 {len(runs)}개 찾아서 수정하지 않았습니다.')
        return
    r = runs[0]
    ed.set_run(r, key, val, slide_no)
    para_runs = r.getparent().findall('a:r', NS)
    ed.clear_runs(para_runs[para_runs.index(r) + 1:])


def _edit_communion(ed, root, slide_no):
    val = ed.value('communion_antiphon')
    if val is None:
        return
    bodies = _body_shapes(root, {'영성체송', '영성체예식'})
    if len(bodies) != 1:
        ed.warn(f'영성체송: 본문 칸을 {len(bodies)}개 찾아서 수정하지 않았습니다.')
        return
    runs = [r for r in _runs(bodies[0]) if _run_text(r)]
    ed.set_run(runs[0], 'communion_antiphon', val, slide_no)
    ed.clear_runs(runs[1:])


def apply_text_fields(ppt_bytes: bytes, fields: dict):
    """fields 값으로 표지와 말씀 전례·영성체송 슬라이드의 지정 run만 교체한다."""
    pkg = Package(ppt_bytes)
    report = {'changed': [], 'warnings': []}
    ed = _Editor(fields, report)
    parts = pkg.slide_parts()

    cover = pkg.xml(parts[0])
    _edit_cover(ed, cover, 1)
    pkg.write_xml(parts[0], cover)

    slides = find_text_slides(pkg)
    edits = {
        '제1독서': lambda r, n: _edit_source(ed, r, n, '제1독서', '▥', 'first_reading_intro', 'first_reading_ref'),
        '화답송': lambda r, n: _edit_single_mark(ed, r, n, '화답송', '◎', 'psalm_refrain'),
        '제2독서': lambda r, n: _edit_source(ed, r, n, '제2독서', '▥', 'second_reading_intro', 'second_reading_ref'),
        '복음환호송': lambda r, n: _edit_single_mark(ed, r, n, '복음환호송', '○', 'gospel_acclamation'),
        '복음': lambda r, n: _edit_source(ed, r, n, '복음', '✠', 'gospel_intro', 'gospel_ref', contains=True),
        '영성체송': lambda r, n: _edit_communion(ed, r, n),
    }
    for title, fn in edits.items():
        nums = slides[title]
        if len(nums) != 1:
            report['warnings'].append(f'"{title}" 슬라이드가 {len(nums)}개라서 수정하지 않았습니다.')
            continue
        part = parts[nums[0] - 1]
        root = pkg.xml(part)
        fn(root, nums[0])
        pkg.write_xml(part, root)

    return pkg.save(), report
