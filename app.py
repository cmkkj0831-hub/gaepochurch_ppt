"""
개포동성당 초등부 미사 PPT 자동 생성 웹앱.
Streamlit 기반, 비개발자(교사) 친화적 UI.
"""
import hmac
import io
import re

import streamlit as st
from PIL import Image

from mass_ppt.generate import generate
from mass_ppt.hymns import HymnTemplateError, find_hymn_bundles
from mass_ppt.parser import FIELD_KEYS, ParseError, parse_mass_url
from mass_ppt.pptx_package import NS, Package
from mass_ppt.storage import SONG_NAME_RE, GitHubStorage, LocalStorage, Storage, StorageError
from mass_ppt.text_edit import DATE_RE, find_text_slides

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

st.set_page_config(
    page_title="개포동성당 초등부 미사 PPT",
    page_icon="⛪",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton > button, .stDownloadButton > button { width: 100%; }
</style>
""", unsafe_allow_html=True)


@st.cache_resource
def get_storage(token: str, repo: str, branch: str) -> Storage:
    """GitHub 토큰이 있으면 GitHubStorage, 없으면 LocalStorage. 토큰이 바뀌면 새로 만든다."""
    if token:
        return GitHubStorage(token, repo, branch)
    return LocalStorage()


def _secret(key: str, default: str = '') -> str:
    try:
        return st.secrets.get(key, default) or default
    except Exception:
        return default


storage = get_storage(_secret('GITHUB_TOKEN'), _secret('GITHUB_REPO', 'cmkkj0831-hub/gaepochurch_ppt'),
                      _secret('GITHUB_BRANCH', 'main'))

EDIT_FIELDS = [
    ('cover_title', '표지 제목', '예: 연중 제 29 주일(전교 주일)', False),
    ('cover_date', '표지 날짜 (토요일)', '예: 2026.10.17', False),
    ('first_reading_intro', '제1독서 출처 문장', '예: ▥ 이사야서의 말씀입니다.', False),
    ('first_reading_ref', '제1독서 장절', '예: 2,1-5', False),
    ('psalm_refrain', '화답송 후렴', '예: ◎ 주님은 민족들의 눈앞에 당신 정의를 드러내셨네.', True),
    ('second_reading_intro', '제2독서 출처 문장', '예: ▥ 사도 바오로의 로마서 말씀입니다.', False),
    ('second_reading_ref', '제2독서 장절', '예: 10,9-18', False),
    ('gospel_acclamation', '복음 환호송 구절', '예: ○ 주님이 말씀하신다. …', True),
    ('gospel_intro', '복음 출처 문장', '예: ✠ 마태오가 전한 거룩한 복음입니다.', False),
    ('gospel_ref', '복음 장절', '예: 28,16-20', False),
    ('communion_antiphon', '영성체송', '영성체송 본문', True),
]

HYMN_INPUTS = [
    ('entrance', '입당 성가'),
    ('offertory', '봉헌 성가'),
    ('communion1', '성체 성가 1'),
    ('communion2', '성체 성가 2 (없으면 비워두기)'),
    ('recessional', '파견 성가'),
]


def cover_date_of(ppt_bytes: bytes) -> str:
    """첫 슬라이드에서 YYYY.MM.DD 날짜를 찾는다."""
    try:
        pkg = Package(ppt_bytes)
        root = pkg.xml(pkg.slide_parts()[0])
        for t in root.iter('{%s}t' % NS['a']):
            if t.text and DATE_RE.match(t.text.strip()):
                return t.text.strip()
    except Exception:
        pass
    return ""


def check_template(ppt_bytes: bytes) -> list:
    """다음 주 자동 생성에 필요한 구조 점검. 문제 목록을 반환."""
    problems = []
    pkg = Package(ppt_bytes)
    bundles = find_hymn_bundles(pkg)
    if len(bundles) != 4:
        problems.append(f"성가 번호 슬라이드 묶음이 {len(bundles)}개입니다 (4개여야 함).")
    for title, nums in find_text_slides(pkg).items():
        if len(nums) != 1:
            problems.append(f'"{title}" 슬라이드가 {len(nums)}개입니다 (1개여야 함).')
    if not cover_date_of(ppt_bytes):
        problems.append("표지에서 날짜(예: 2026.10.17)를 찾지 못했습니다.")
    return problems


def filename_from_cover_date(cover_date: str) -> str:
    m = re.fullmatch(r'\d{2}(\d{2})\.(\d{2})\.(\d{2})', (cover_date or '').strip())
    return f"{m.group(1)}{m.group(2)}{m.group(3)}미사.pptx" if m else ""


def hymn_numbers_from_state() -> dict:
    def num(key):
        v = (st.session_state.get(f"hymn_{key}") or "").strip()
        return int(v) if v.isdigit() and int(v) > 0 else None

    return {
        'entrance': num('entrance'),
        'offertory': num('offertory'),
        'communion': [n for n in (num('communion1'), num('communion2')) if n],
        'recessional': num('recessional'),
    }


def page_create_ppt():
    st.header("📝 PPT 만들기", divider="blue")

    st.subheader("1. 기준 PPT")
    archive_files = storage.list_archive()
    base_ppt_bytes, base_label = None, ""
    use_upload = st.checkbox("다른 PPT 파일을 직접 올려서 기준으로 쓰기")
    if use_upload:
        up = st.file_uploader("기준 PPT 파일", type=["pptx"], key="base_upload")
        if up:
            base_ppt_bytes, base_label = up.getvalue(), up.name
    elif archive_files:
        names = [f.name for f in archive_files]
        base_label = st.selectbox("기준 최종본 (기본: 가장 최근)", names, index=0)
        with st.spinner("기준 PPT를 가져오는 중..."):
            base_ppt_bytes = storage.read_archive(base_label)
        if base_ppt_bytes is None:
            st.error(f"{base_label} 파일을 가져오지 못했습니다. 잠시 후 다시 시도해 주세요.")
    else:
        st.error("저장된 최종본이 없습니다. 위 체크박스로 기준 PPT를 직접 올려주세요.")
    if base_label:
        st.caption(f"기준: {base_label}")

    st.subheader("2. 매일미사 불러오기")
    url = st.text_input("매일미사 URL", placeholder="https://maria.catholic.or.kr/mi_pr/missa/missa.asp?...")
    if st.button("🔄 불러오기", disabled=not url):
        with st.spinner("매일미사 페이지를 읽는 중..."):
            try:
                parsed = parse_mass_url(url)
                for key in FIELD_KEYS:
                    st.session_state[f"f_{key}"] = parsed.get(key) or ""
                st.session_state.parse_missing = [k for k, *_ in EDIT_FIELDS if not parsed.get(k)]
                st.session_state.pop('generated_ppt', None)
                st.success("불러왔습니다. 아래 내용을 확인해 주세요.")
            except ParseError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"매일미사 페이지를 읽는 중 문제가 생겼습니다: {e}")

    st.subheader("3. 내용 확인·수정")
    st.caption("PPT는 아래 칸의 값으로만 만들어집니다. 틀린 곳은 직접 고쳐주세요.")
    missing = set(st.session_state.get('parse_missing', []))
    for key, label, placeholder, long in EDIT_FIELDS:
        if key in missing:
            st.error(f"⚠️ {label}: 페이지에서 찾지 못했습니다. 직접 입력해 주세요.")
        widget = st.text_area if long else st.text_input
        widget(label, key=f"f_{key}", placeholder=placeholder)

    st.subheader("4. 성가 번호")
    for key, label in HYMN_INPUTS:
        st.text_input(label, key=f"hymn_{key}", placeholder="번호만 입력 (예: 58)")
        raw = (st.session_state.get(f"hymn_{key}") or "").strip()
        if not raw:
            continue
        if not raw.isdigit():
            st.error(f"❌ '{raw}'는 숫자가 아닙니다.")
            continue
        files = storage.hymn_files(int(raw))
        if files:
            st.markdown(f"**{raw}번 ✅ {len(files)}장**")
            with st.expander(f"{raw}번 악보 보기"):
                for name in files:
                    data = storage.read_song_file(name)
                    if data:
                        st.image(Image.open(io.BytesIO(data)), caption=name, width=240)
        else:
            st.markdown(f"**{raw}번 ❌ 이미지 없음** — 이 성가는 PPT에 들어가지 않습니다.")

    st.subheader("5. PPT 만들기")
    if st.button("🚀 PPT 생성", type="primary", disabled=base_ppt_bytes is None):
        fields = {key: (st.session_state.get(f"f_{key}") or "").strip() or None for key in FIELD_KEYS}
        with st.spinner("PPT를 만드는 중..."):
            try:
                result, report = generate(base_ppt_bytes, fields, hymn_numbers_from_state(),
                                          storage.get_hymn_images)
                name = filename_from_cover_date(fields.get('cover_date')) or "미사.pptx"
                st.session_state.generated_ppt = (name, result, report)
            except HymnTemplateError as e:
                st.error(f"기준 PPT 구조 문제로 만들 수 없습니다: {e}")
            except StorageError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"PPT를 만드는 중 문제가 생겼습니다: {e}")

    generated = st.session_state.get('generated_ppt')
    if generated:
        name, data, report = generated
        st.success(f"완성! 슬라이드 {report.get('slide_count')}장")
        st.download_button("📥 " + name + " 다운로드", data=data, file_name=name, mime=PPTX_MIME)
        with st.expander("생성 내용 요약", expanded=True):
            for h in report.get('hymns') or []:
                st.write(f"• {h['role']} {h['num']}번 ({h['pages']}장)")
            for c in report.get('changed_texts') or []:
                st.write(f"• {c['field']}: {c['after'].strip()}")
            for w in report.get('warnings') or []:
                st.warning(w)
        st.info("다운로드한 파일을 다듬은 뒤 '최종본 올리기'에 꼭 올려주세요. 다음 주 PPT의 기준이 됩니다.")


def page_upload_archive():
    st.header("📤 최종본 올리기", divider="blue")

    uploaded = st.file_uploader("다듬은 최종본 PPT", type=["pptx"], key="archive_upload")
    if uploaded:
        ppt_bytes = uploaded.getvalue()
        try:
            problems = check_template(ppt_bytes)
        except Exception:
            st.error("PPT 파일을 열 수 없습니다. 올바른 .pptx 파일인지 확인해 주세요.")
            return
        if problems:
            st.warning("다음 주 자동 생성에 문제가 생길 수 있습니다:\n\n" + "\n".join(f"- {p}" for p in problems))
        else:
            st.success("PPT 구조 점검 통과")

        date_str = cover_date_of(ppt_bytes)
        if date_str:
            st.caption(f"표지 날짜 인식: {date_str}")
        filename = st.text_input("저장 파일명", value=filename_from_cover_date(date_str),
                                 placeholder="YYMMDD미사.pptx (예: 261017미사.pptx)")
        uploader = st.text_input("올린 사람 이름 (선택)")
        valid = bool(re.fullmatch(r'\d{6}미사\.pptx', filename or ''))
        if filename and not valid:
            st.error("파일명은 YYMMDD미사.pptx 형식이어야 합니다 (토요일 날짜, 예: 261017미사.pptx).")

        exists = valid and filename in [f.name for f in storage.list_archive()]
        overwrite = True
        if exists:
            overwrite = st.checkbox(f"{filename} 파일이 이미 있습니다. 덮어쓸까요?")

        if st.button("💾 저장", type="primary", disabled=not (valid and overwrite)):
            with st.spinner("저장 중... (파일이 커서 1분 정도 걸릴 수 있습니다)"):
                try:
                    storage.save_archive(filename, ppt_bytes, uploader or None)
                    st.success(f"{filename} 저장 완료. 다음 PPT 만들기의 기준으로 선택됩니다.")
                except StorageError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"저장 중 문제가 생겼습니다: {e}")

    st.subheader("최근 최종본")
    archive_files = storage.list_archive()
    if not archive_files:
        st.info("저장된 최종본이 없습니다")
    for info in archive_files[:10]:
        st.write(f"📄 **{info.name}** · {info.size / 1024 / 1024:.1f}MB"
                 + (f" · {info.uploader}" if info.uploader else ""))
        if st.button("파일 준비", key=f"prep_{info.name}"):
            with st.spinner("파일을 가져오는 중..."):
                st.session_state[f"dl_{info.name}"] = storage.read_archive(info.name)
        data = st.session_state.get(f"dl_{info.name}")
        if data:
            st.download_button(f"📥 {info.name} 다운로드", data=data, file_name=info.name,
                               mime=PPTX_MIME, key=f"download_{info.name}")


def admin_password() -> str:
    return _secret('ADMIN_PASSWORD')


def song_groups() -> dict:
    groups = {}
    for f in storage.list_songs():
        m = SONG_NAME_RE.match(f.name)
        if m:
            groups.setdefault(int(m.group(1)), []).append(f.name)
    return groups


def normalized_song_names(num: int, files) -> dict:
    """업로드 파일들 -> {저장 파일명: bytes}. 1장이면 num.ext, 여러 장이면 num-1.ext …"""
    files = sorted(files, key=lambda f: f.name)
    out = {}
    for i, f in enumerate(files, 1):
        ext = 'jpg' if f.name.lower().endswith(('.jpg', '.jpeg')) else 'png'
        name = f"{num}.{ext}" if len(files) == 1 else f"{num}-{i}.{ext}"
        out[name] = f.getvalue()
    return out


def confirm_box(state_key: str, message: str, on_confirm):
    """state_key가 설정돼 있으면 확인/취소 버튼을 보여준다."""
    if not st.session_state.get(state_key):
        return
    st.warning(message)
    if st.button("✅ 네, 삭제합니다", key=f"{state_key}_yes", type="primary"):
        try:
            on_confirm(st.session_state[state_key])
            st.session_state[f"{state_key}_done"] = f"{st.session_state[state_key]} 삭제 완료"
        except StorageError as e:
            st.error(str(e))
            return
        st.session_state[state_key] = None
        st.rerun()
    if st.button("취소", key=f"{state_key}_no"):
        st.session_state[state_key] = None
        st.rerun()


def page_manage_songs():
    st.header("🎵 성가 관리", divider="blue")

    expected = admin_password()
    if not st.session_state.get('admin_ok'):
        if not expected:
            st.error("관리자 비밀번호(ADMIN_PASSWORD)가 설정되지 않아 이 페이지를 사용할 수 없습니다.")
            return
        pw = st.text_input("관리자 비밀번호", type="password", key="admin_pw_input")
        if st.button("로그인"):
            if hmac.compare_digest(pw.encode(), expected.encode()):
                st.session_state.admin_ok = True
                st.rerun()
            else:
                st.error("비밀번호가 틀렸습니다.")
        return

    for key in ('pending_song_delete_done', 'pending_archive_delete_done'):
        if st.session_state.get(key):
            st.success(st.session_state.pop(key))

    # 1. 성가 찾기·교체·삭제
    st.subheader("1. 등록된 성가")
    groups = song_groups()
    st.caption(f"등록된 성가 {len(groups)}곡: " + ", ".join(str(n) for n in sorted(groups)))
    q = st.text_input("성가 번호 검색", placeholder="예: 58", key="song_search").strip()
    confirm_box('pending_song_delete', f"{st.session_state.get('pending_song_delete')}번 성가 이미지를 모두 삭제할까요?",
                lambda num: storage.change_songs({}, groups.get(int(num), []), f"성가 삭제: {num}"))
    if q:
        if not q.isdigit() or int(q) not in groups:
            st.info(f"{q}번 성가는 등록되어 있지 않습니다. 아래 '새 성가 올리기'로 추가하세요.")
        else:
            num = int(q)
            names = storage.hymn_files(num) or sorted(groups[num])
            st.markdown(f"**{num}번 ({len(groups[num])}개 파일)**")
            for name in names:
                data = storage.read_song_file(name)
                if data:
                    st.image(Image.open(io.BytesIO(data)), caption=name, width=200)
            new_files = st.file_uploader(f"{num}번 교체할 이미지 (여러 장이면 순서대로 이름 정렬)",
                                         type=["png", "jpg", "jpeg"], accept_multiple_files=True,
                                         key=f"replace_{num}")
            if new_files:
                add = normalized_song_names(num, new_files)
                st.caption("저장될 이름: " + ", ".join(add))
                if st.button(f"🔁 {num}번 교체 저장", type="primary"):
                    try:
                        with st.spinner("교체 중..."):
                            storage.change_songs(add, groups[num], f"성가 교체: {num}")
                        st.success(f"{num}번 교체 완료")
                    except StorageError as e:
                        st.error(str(e))
            if st.button(f"🗑️ {num}번 삭제", key=f"del_{num}"):
                st.session_state.pending_song_delete = str(num)
                st.rerun()

    # 2. 새 성가 올리기
    st.subheader("2. 새 성가 올리기")
    uploads = st.file_uploader("성가 이미지 (여러 파일 가능, 예: 58-1.png, 58-2.png, 128.png)",
                               type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="song_upload")
    if uploads:
        add, problems = {}, []
        for i, f in enumerate(uploads):
            m = SONG_NAME_RE.match(f.name)
            if m:
                num, page = m.group(1), m.group(2)
            else:
                st.warning(f"'{f.name}': 파일명에서 번호를 알 수 없습니다. 직접 입력해 주세요.")
                num = st.text_input(f"{f.name} 성가 번호", key=f"up_num_{i}").strip()
                page = st.text_input(f"{f.name} 쪽 번호 (한 장이면 비워두기)", key=f"up_page_{i}").strip()
            if not num or not num.isdigit() or (page and not page.isdigit()):
                problems.append(f.name)
                continue
            ext = 'jpg' if f.name.lower().endswith(('.jpg', '.jpeg')) else 'png'
            name = f"{int(num)}-{int(page)}.{ext}" if page else f"{int(num)}.{ext}"
            add[name] = f.getvalue()
        existing = {n for names in groups.values() for n in names}
        if add:
            st.caption("저장될 이름: " + ", ".join(sorted(add)))
            overwrite = sorted(set(add) & existing)
            if overwrite:
                st.warning("이미 있는 파일은 새 파일로 바뀝니다: " + ", ".join(overwrite))
        if problems:
            st.error("번호가 없는 파일이 있습니다: " + ", ".join(problems))
        if st.button("📤 업로드", type="primary", disabled=not add or bool(problems)):
            try:
                with st.spinner("업로드 중..."):
                    storage.change_songs(add, [], "성가 추가: " + ", ".join(sorted(add)))
                st.success("업로드 완료: " + ", ".join(sorted(add)))
            except StorageError as e:
                st.error(str(e))

    # 3. 최종본 삭제
    st.subheader("3. 최종본 PPT 삭제")
    archives = [f.name for f in storage.list_archive()]
    confirm_box('pending_archive_delete', f"{st.session_state.get('pending_archive_delete')} 파일을 삭제할까요? 되돌릴 수 없습니다.",
                storage.delete_archive)
    if archives:
        target = st.selectbox("삭제할 최종본", archives, key="archive_delete_target")
        if st.button("🗑️ 이 최종본 삭제", key="archive_delete_btn"):
            st.session_state.pending_archive_delete = target
            st.rerun()
    else:
        st.info("저장된 최종본이 없습니다.")

    st.divider()
    if st.button("🔓 로그아웃"):
        st.session_state.admin_ok = False
        st.rerun()


def guarded(page):
    try:
        page()
    except StorageError as e:
        st.error(str(e))


def main():
    st.title("⛪ 개포동성당 초등부 미사 PPT")
    tab1, tab2, tab3 = st.tabs(["📝 PPT 만들기", "📤 최종본 올리기", "🎵 성가 관리"])
    with tab1:
        guarded(page_create_ppt)
    with tab2:
        guarded(page_upload_archive)
    with tab3:
        guarded(page_manage_songs)


main()
