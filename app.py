"""
개포동성당 초등부 미사 PPT 자동 생성 웹앱.
Streamlit 기반, 비개발자(교사) 친화적 UI.
"""
import streamlit as st
from pathlib import Path
import io
from datetime import datetime, timedelta
from PIL import Image
import zipfile

from mass_ppt.parser import parse_mass_url
from mass_ppt.text_edit import update_ppt
from mass_ppt.generate import generate_from_local_hymns
from mass_ppt.storage import LocalStorage, GitHubStorage, Storage
from mass_ppt.hymns import _find_hymn_slides
from lxml import etree


# 설정
st.set_page_config(
    page_title="개포동성당 초등부 미사 PPT",
    page_icon="⛪",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 스타일
st.markdown("""
<style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
    .stButton > button { width: 100%; }
    .success { color: green; }
    .error { color: red; }
    .warning { color: orange; }
    .preview-img { max-width: 200px; border-radius: 4px; }
    input[type="text"], textarea { font-family: 'Courier New', monospace; }
</style>
""", unsafe_allow_html=True)

# 저장소 초기화
@st.cache_resource
def get_storage() -> Storage:
    """GitHub 설정이 있으면 GitHubStorage, 없으면 LocalStorage 사용."""
    try:
        # Streamlit secrets에서 GitHub 설정 읽기
        if hasattr(st, 'secrets') and 'GITHUB_TOKEN' in st.secrets:
            token = st.secrets.get('GITHUB_TOKEN')
            repo = st.secrets.get('GITHUB_REPO', 'cmkkj0831-hub/gaepochurch_ppt')
            branch = st.secrets.get('GITHUB_BRANCH', 'main')

            if token:
                return GitHubStorage(token, repo, branch)
    except Exception:
        pass

    # 기본값: LocalStorage
    return LocalStorage()

storage = get_storage()


def extract_date_from_ppt(ppt_bytes: bytes) -> str:
    """PPT 표지 슬라이드에서 날짜 추출 (YYYY.MM.DD)."""
    try:
        with zipfile.ZipFile(io.BytesIO(ppt_bytes)) as z:
            slide1 = z.read('ppt/slides/slide1.xml')

        tree = etree.fromstring(slide1)
        namespaces = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        }

        # 날짜 형식 찾기 (YYYY.MM.DD)
        for t in tree.findall('.//a:t', namespaces):
            if t.text and '.' in t.text:
                parts = t.text.split('.')
                if len(parts) == 3 and all(p.isdigit() for p in parts):
                    return t.text
    except Exception:
        pass

    return ""


def validate_ppt(ppt_bytes: bytes) -> tuple[bool, str]:
    """PPT 유효성 검증."""
    try:
        # PPTX 형식 확인
        with zipfile.ZipFile(io.BytesIO(ppt_bytes)) as z:
            # 필수 파일 확인
            z.getinfo('ppt/presentation.xml')

            # 슬라이드 확인
            slide_files = z.namelist()
            if not any(f.startswith('ppt/slides/slide') for f in slide_files):
                return False, "슬라이드가 없습니다"

        return True, "유효한 PPT입니다"
    except Exception as e:
        return False, f"PPT 읽기 실패: {str(e)}"


def page_create_ppt():
    """페이지 1: PPT 만들기."""
    st.header("📝 PPT 만들기", divider="blue")

    # 기준 PPT 선택
    st.subheader("1️⃣ 기준 PPT 선택")

    col1, col2 = st.columns([2, 1])
    with col1:
        archive_files = storage.list_archive()
        if archive_files:
            file_options = [f"{f.name} ({f.modified_time.strftime('%m월 %d일')})" for f in archive_files]
            selected_idx = st.selectbox(
                "기준 PPT",
                range(len(file_options)),
                format_func=lambda i: file_options[i],
                label_visibility="collapsed"
            )
            base_filename = archive_files[selected_idx].name
            st.caption(f"기준: {base_filename}")
        else:
            st.error("저장된 PPT가 없습니다")
            return

    with col2:
        if st.button("📂 직접 업로드"):
            uploaded_file = st.file_uploader("PPT 선택", type=["pptx"], key="base_upload")
            if uploaded_file:
                st.session_state.custom_base_ppt = uploaded_file.read()
                st.success("업로드됨")

    # 기준 PPT 결정
    if 'custom_base_ppt' in st.session_state:
        base_ppt_bytes = st.session_state.custom_base_ppt
        st.info("📂 업로드된 PPT를 사용합니다")
    else:
        base_ppt_bytes = storage.read_archive(base_filename)

    # 매일미사 URL 입력
    st.subheader("2️⃣ 매일미사 정보")

    url = st.text_input(
        "매일미사 URL",
        placeholder="https://maria.catholic.or.kr/mi_pr/missa/...",
        help="매일미사 웹사이트의 미사 정보 페이지 URL"
    )

    parsed = None
    if url:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.text_input("URL 상태", "✅ 입력됨", disabled=True)
        with col2:
            if st.button("🔄 불러오기", use_container_width=True):
                with st.spinner("파싱 중..."):
                    try:
                        parsed = parse_mass_url(url)
                        st.session_state.parsed = parsed
                        st.success("✅ 파싱 완료")
                    except Exception as e:
                        st.error(f"❌ 파싱 실패: {str(e)}")

    # 파싱 결과 수정
    if 'parsed' in st.session_state:
        parsed = st.session_state.parsed
        st.subheader("3️⃣ 전례 정보 수정")

        col1, col2 = st.columns(2)

        with col1:
            parsed['liturgy_title'] = st.text_input(
                "전례일 제목",
                value=parsed.get('liturgy_title') or "",
                placeholder="예: (녹) 연중 제29주일(전교 주일)"
            )

            parsed['first_reading_ref'] = st.text_input(
                "제1독서 장절",
                value=parsed.get('first_reading_ref') or "",
                placeholder="예: 2,1-5"
            )

            parsed['second_reading_ref'] = st.text_input(
                "제2독서 장절",
                value=parsed.get('second_reading_ref') or "",
                placeholder="예: 10,9-18"
            )

        with col2:
            date_val = st.text_input(
                "표지 날짜",
                value=parsed.get('date') or "",
                placeholder="YYYY-MM-DD (일요일)",
                help="일요일 날짜. 자동으로 토요일로 변환됩니다"
            )
            if date_val:
                parsed['date'] = date_val

            parsed['responsorial_refrain'] = st.text_area(
                "화답송 후렴",
                value=parsed.get('responsorial_refrain') or "",
                height=60,
                placeholder="예: ◎ 주님은 민족들의 눈앞에..."
            )

            parsed['gospel_ref'] = st.text_input(
                "복음 장절",
                value=parsed.get('gospel_ref') or "",
                placeholder="예: 28,16-20"
            )

        st.session_state.parsed = parsed

        # 성가 번호 입력
        st.subheader("4️⃣ 성가 번호")

        col1, col2 = st.columns(2)

        with col1:
            entrance = st.number_input("입당", min_value=0, value=0, step=1)
            offertory = st.number_input("봉헌", min_value=0, value=0, step=1)

        with col2:
            communion1 = st.number_input("성체 1", min_value=0, value=0, step=1)
            communion2 = st.number_input("성체 2", min_value=0, value=0, step=1, help="(선택)")

        recessional = st.number_input("파견", min_value=0, value=0, step=1)

        # 성가 번호 검증
        hymn_numbers = {
            'entrance': entrance,
            'offertory': offertory,
            'communion': [communion1] + ([communion2] if communion2 > 0 else []),
            'recessional': recessional
        }

        hymn_cols = st.columns(len([h for h in [entrance, offertory, communion1, communion2, recessional] if h > 0]))
        for col, hymn in zip(hymn_cols, [entrance, offertory, communion1, communion2, recessional]):
            if hymn > 0:
                img = storage.read_song(hymn)
                with col:
                    if img:
                        st.image(Image.open(io.BytesIO(img)), width=80, caption=f"{hymn}")
                        st.caption("✅", unsafe_allow_html=True)
                    else:
                        st.warning(f"{hymn}\n❌", unsafe_allow_html=True)

        # PPT 생성
        st.subheader("5️⃣ PPT 생성")

        if st.button("🚀 PPT 생성", use_container_width=True, type="primary"):
            with st.spinner("생성 중..."):
                try:
                    result_bytes, report = generate_from_local_hymns(
                        base_ppt_bytes,
                        url,
                        hymn_numbers
                    )

                    # 다운로드 파일명 (토요일 기준)
                    date_parts = parsed.get('date', '').split('-')
                    if len(date_parts) == 3:
                        from datetime import datetime, timedelta
                        sunday = datetime.strptime(parsed['date'], '%Y-%m-%d')
                        saturday = sunday - timedelta(days=1)
                        filename = saturday.strftime('%y%m%d') + '미사.pptx'
                    else:
                        filename = 'generated.pptx'

                    st.session_state.generated_ppt = (filename, result_bytes)
                    st.success("✅ PPT 생성 완료")

                    # 보고서
                    with st.expander("📋 생성 보고서"):
                        st.write("**변경된 텍스트**")
                        for text in report['updated_texts']:
                            st.caption(f"  • {text}")

                        if report['warnings']:
                            st.warning("**⚠️ 경고**")
                            for warn in report['warnings']:
                                st.caption(f"  • {warn}")

                except Exception as e:
                    st.error(f"❌ 생성 실패: {str(e)}")

        # 다운로드
        if 'generated_ppt' in st.session_state:
            filename, ppt_bytes = st.session_state.generated_ppt
            st.download_button(
                label="📥 PPT 다운로드",
                data=ppt_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

            st.info("💡 **다음 단계**: 다운로드한 파일을 필요에 따라 다듬은 뒤 '최종본 올리기' 탭에서 올려주세요. 이 파일이 다음 주 PPT의 기준이 됩니다.")


def page_upload_archive():
    """페이지 2: 최종본 올리기."""
    st.header("📤 최종본 올리기", divider="blue")

    # 파일 업로드
    st.subheader("최종본 PPT 업로드")

    uploaded_file = st.file_uploader(
        "PPT 파일 선택",
        type=["pptx"],
        help="다듬은 최종본 PPT를 올려주세요"
    )

    if uploaded_file:
        ppt_bytes = uploaded_file.read()

        # 유효성 검증
        valid, msg = validate_ppt(ppt_bytes)
        if not valid:
            st.error(f"❌ {msg}")
            return

        st.success("✅ 유효한 PPT입니다")

        # 파일명 제안
        date_str = extract_date_from_ppt(ppt_bytes)

        if date_str:
            # YYYY.MM.DD → YYMMDD
            parts = date_str.split('.')
            suggested_filename = f"{parts[0][2:]}{parts[1]}{parts[2]}미사.pptx"
            st.caption(f"자동 인식: {date_str}")
        else:
            suggested_filename = "미사.pptx"
            st.warning("⚠️ 표지 슬라이드에서 날짜를 찾을 수 없습니다")

        # 파일명 입력
        filename = st.text_input(
            "저장 파일명",
            value=suggested_filename,
            placeholder="YYMMDD미사.pptx",
            help="YYMMDD미사.pptx 형식 (예: 261017미사.pptx)"
        )

        if filename and not filename.endswith('.pptx'):
            filename += '.pptx'

        # 파일명 검증
        if filename and not storage._validate_ppt_filename(filename):
            st.error(f"❌ 파일명 형식이 잘못되었습니다. YYMMDD미사.pptx 형식을 사용하세요.")
            return

        # 올린 사람
        uploader = st.text_input(
            "올린 사람 이름 (선택)",
            placeholder="예: 김지현 선생님",
            help="(선택) 이 파일을 올린 사람의 이름을 기록합니다"
        )

        # 저장
        col1, col2 = st.columns(2)

        with col1:
            if st.button("💾 저장", use_container_width=True, type="primary"):
                # 기존 파일 확인
                if filename in [f.name for f in storage.list_archive()]:
                    if st.checkbox("같은 날짜의 파일이 있습니다. 덮어쓰시겠습니까?"):
                        if storage.save_archive(filename, ppt_bytes, uploader):
                            st.success(f"✅ {filename}로 저장되었습니다")
                            st.session_state.generated_ppt = None  # 초기화
                        else:
                            st.error("❌ 저장 실패")
                else:
                    if storage.save_archive(filename, ppt_bytes, uploader):
                        st.success(f"✅ {filename}로 저장되었습니다")
                        st.session_state.generated_ppt = None
                    else:
                        st.error("❌ 저장 실패")

        with col2:
            if st.button("❌ 취소", use_container_width=True):
                st.rerun()

    # 최근 최종본 목록
    st.subheader("최근 최종본")

    archive_files = storage.list_archive()

    if archive_files:
        for file_info in archive_files[:10]:  # 최근 10개
            col1, col2, col3, col4 = st.columns([2, 1, 1, 1])

            with col1:
                st.caption(f"📄 {file_info.name}")

            with col2:
                st.caption(f"{file_info.modified_time.strftime('%m월 %d일')}")

            with col3:
                size_kb = file_info.size / 1024
                st.caption(f"{size_kb:.0f}KB")

            with col4:
                if st.button("⬇️", key=f"dl_{file_info.name}", help="다운로드"):
                    ppt_bytes = storage.read_archive(file_info.name)
                    if ppt_bytes:
                        st.download_button(
                            label="다운로드",
                            data=ppt_bytes,
                            file_name=file_info.name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"download_{file_info.name}"
                        )
    else:
        st.info("저장된 최종본이 없습니다")


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
        import re
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
                img = storage.read_song(hymn_num)
                if img:
                    st.image(Image.open(io.BytesIO(img)), width=80, caption=f"{hymn_num}번")

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
        import re
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


# 메인
def main():
    st.title("⛪ 개포동성당 초등부 미사 PPT")
    st.markdown("비개발자를 위한 간편한 PPT 자동 생성 시스템")

    # 탭 구성
    tab1, tab2, tab3 = st.tabs(["📝 PPT 만들기", "📤 최종본 올리기", "🎵 성가 관리"])

    with tab1:
        page_create_ppt()

    with tab2:
        page_upload_archive()

    with tab3:
        page_manage_songs()


if __name__ == "__main__":
    main()
