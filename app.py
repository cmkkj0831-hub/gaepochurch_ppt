"""
개포동성당 초등부 미사 PPT 자동 생성 웹앱.
Streamlit 기반, 비개발자(교사) 친화적 UI.
"""
import io
import re
from pathlib import Path

import streamlit as st
from PIL import Image

from mass_ppt.generate import generate
from mass_ppt.hymns import HymnTemplateError, find_hymn_bundles
from mass_ppt.parser import FIELD_KEYS, ParseError, parse_mass_url
from mass_ppt.pptx_package import NS, Package
from mass_ppt.storage import GitHubStorage, LocalStorage, Storage
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
def get_storage() -> Storage:
    """GitHub 설정이 있으면 GitHubStorage, 없으면 LocalStorage 사용."""
    try:
        if 'GITHUB_TOKEN' in st.secrets and st.secrets.get('GITHUB_TOKEN'):
            return GitHubStorage(
                st.secrets['GITHUB_TOKEN'],
                st.secrets.get('GITHUB_REPO', 'cmkkj0831-hub/gaepochurch_ppt'),
                st.secrets.get('GITHUB_BRANCH', 'main'),
            )
    except Exception:
        pass
    return LocalStorage()


storage = get_storage()

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
        base_ppt_bytes = storage.read_archive(base_label)
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
        try:
            files = storage.hymn_files(int(raw))
        except Exception as e:
            st.error(f"성가 목록을 읽지 못했습니다: {e}")
            continue
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
            with st.spinner("저장 중..."):
                try:
                    ok = storage.save_archive(filename, ppt_bytes, uploader or None)
                except Exception as e:
                    ok = False
                    st.error(f"저장 중 문제가 생겼습니다: {e}")
            if ok:
                st.cache_data.clear()
                st.success(f"{filename} 저장 완료. 다음 PPT 만들기의 기준으로 선택됩니다.")
            else:
                st.error("저장하지 못했습니다. 잠시 후 다시 시도해 주세요.")

    st.subheader("최근 최종본")
    archive_files = storage.list_archive()
    if not archive_files:
        st.info("저장된 최종본이 없습니다")
    for info in archive_files[:10]:
        st.write(f"📄 **{info.name}** · {info.size / 1024 / 1024:.1f}MB"
                 + (f" · {info.uploader}" if info.uploader else ""))
        if st.button("파일 준비", key=f"prep_{info.name}"):
            st.session_state[f"dl_{info.name}"] = storage.read_archive(info.name)
        data = st.session_state.get(f"dl_{info.name}")
        if data:
            st.download_button(f"📥 {info.name} 다운로드", data=data, file_name=info.name,
                               mime=PPTX_MIME, key=f"download_{info.name}")


def page_manage_songs():
    """페이지 3: 성가 관리 (관리자만 접근)."""
    st.header("🎵 성가 관리", divider="blue")

    # 비밀번호 인증
    if 'admin_authenticated' not in st.session_state:
        st.session_state.admin_authenticated = False

    if not st.session_state.admin_authenticated:
        st.subheader("🔐 관리자 인증")
        password = st.text_input("비밀번호", type="password", placeholder="관리자 비밀번호 입력")

        if password:
            try:
                admin_password = st.secrets.get('ADMIN_PASSWORD', '')
                if password == admin_password and admin_password:
                    st.session_state.admin_authenticated = True
                    st.success("✅ 인증 완료")
                    st.rerun()
                else:
                    st.error("❌ 비밀번호가 틀렸습니다")
            except Exception as e:
                st.error(f"❌ 인증 오류: {str(e)}")

        return

    # 관리자 메뉴
    st.subheader("성가 관리")

    # 현재 저장된 성가 목록
    st.subheader("1️⃣ 현재 등록된 성가")

    songs = storage.list_songs()

    if songs:
        # 번호별로 그룹화
        hymn_dict = {}
        for song in songs:
            # 파일명에서 번호 추출 (58.png, 58-1.png 등)
            match = re.match(r'(\d+)', song.name)
            if match:
                hymn_num = int(match.group(1))
                if hymn_num not in hymn_dict:
                    hymn_dict[hymn_num] = []
                hymn_dict[hymn_num].append(song)

        # 번호순으로 정렬
        for hymn_num in sorted(hymn_dict.keys()):
            hymn_files = hymn_dict[hymn_num]

            # 섬네일 + 파일 목록
            cols = st.columns([1, 3, 1])

            with cols[0]:
                images = storage.get_hymn_images(hymn_num)
                if images:
                    st.image(Image.open(io.BytesIO(images[0])), width=80, caption=f"{hymn_num}번")

            with cols[1]:
                st.write(f"**{hymn_num}번** ({len(hymn_files)}장)")
                for hf in hymn_files:
                    st.caption(f"  • {hf.name} ({hf.size/1024:.0f}KB)")

            with cols[2]:
                if st.button("🗑️", key=f"delete_{hymn_num}", help="삭제"):
                    if st.checkbox("정말 삭제하시겠습니까?", key=f"confirm_delete_{hymn_num}"):
                        for hf in hymn_files:
                            if hasattr(storage, '_delete'):
                                path = f"songs/{hf.name}"
                                file_info = storage._get_file(path)
                                if file_info:
                                    storage._delete(f"/contents/{path}", {
                                        "message": f"성가 삭제: {hf.name}",
                                        "sha": file_info.get('sha'),
                                        "branch": storage.branch
                                    })
                        st.cache_data.clear()
                        st.success(f"✅ {hymn_num}번 성가 삭제됨")
                        st.rerun()
    else:
        st.info("저장된 성가 이미지가 없습니다")

    # 새 성가 업로드
    st.subheader("2️⃣ 새 성가 업로드")

    uploaded_files = st.file_uploader(
        "이미지 파일 선택 (여러 파일 가능)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        help="58-1.png, 106-2.jpg 등의 형식으로 업로드"
    )

    if uploaded_files:
        files_to_upload = []

        for uploaded_file in uploaded_files:
            filename = uploaded_file.name

            # 파일명에서 번호 추출
            match = re.match(r'(\d+)(?:-(\d+))?', filename)
            if match:
                hymn_num = int(match.group(1))
                page = int(match.group(2)) if match.group(2) else None

                files_to_upload.append({
                    'file': uploaded_file,
                    'hymn_num': hymn_num,
                    'page': page,
                    'filename': filename
                })
            else:
                st.warning(f"❌ {filename}: 파일명 형식 인식 실패 (예: 58-1.png)")

        # 업로드 전 확인
        if files_to_upload:
            st.write("업로드할 파일:")
            for item in files_to_upload:
                page_str = f"-{item['page']}" if item['page'] else ""
                st.caption(f"  • {item['hymn_num']}{page_str} ({item['file'].name})")

            if st.button("📤 업로드", use_container_width=True, type="primary"):
                with st.spinner("업로드 중..."):
                    for item in files_to_upload:
                        try:
                            file_content = item['file'].read()

                            # GitHub Storage인 경우
                            if hasattr(storage, '_put_large_file'):
                                path = f"songs/{item['filename']}"
                                message = f"성가 추가: {item['filename']}"
                                success = storage._put_large_file(path, file_content, message)
                            else:
                                # Local Storage인 경우
                                path = Path(storage.songs_dir) / item['filename']
                                path.write_bytes(file_content)
                                success = True

                            if success:
                                st.success(f"✅ {item['filename']} 업로드됨")
                        except Exception as e:
                            st.error(f"❌ {item['filename']}: {str(e)}")

                    # 캐시 비우기
                    st.cache_data.clear()
                    st.rerun()

    # 로그아웃
    st.divider()
    if st.button("🔓 로그아웃", use_container_width=True):
        st.session_state.admin_authenticated = False
        st.rerun()


def main():
    st.title("⛪ 개포동성당 초등부 미사 PPT")
    tab1, tab2, tab3 = st.tabs(["📝 PPT 만들기", "📤 최종본 올리기", "🎵 성가 관리"])
    with tab1:
        page_create_ppt()
    with tab2:
        page_upload_archive()
    with tab3:
        page_manage_songs()


main()
