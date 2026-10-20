# 개포동성당 초등부 미사 PPT

매일미사 페이지의 전례 정보와 성가 번호로, 지난주 최종본을 바탕으로 이번 주 초등부 주일학교 미사 PPT를 만드는 웹앱입니다.

앱 주소: https://gaepochurch-ppt.streamlit.app (링크만 있으면 누구나 사용, 성가 관리만 관리자 비밀번호)

## 매주 사용 순서

1. **PPT 만들기**
   - 기준 PPT는 가장 최근 최종본이 자동 선택됩니다 (다른 주 선택·직접 업로드 가능).
   - 매일미사 URL을 넣고 [불러오기] → 표지 제목·날짜, 독서·복음 출처, 화답송, 복음 환호송, 영성체송을 확인하고 틀린 곳은 고칩니다.
   - 성가 번호(입당, 봉헌, 성체 1~2곡, 파견)를 넣으면 "58번 ✅ 2장"처럼 이미지 유무가 표시됩니다.
   - [PPT 생성] → `YYMMDD미사.pptx` 다운로드.
2. **다듬기**: PowerPoint에서 열어 필요한 부분을 손봅니다.
3. **최종본 올리기**: 다듬은 파일을 꼭 올려주세요. 다음 주 PPT의 기준이 됩니다.

매일미사 URL: https://maria.catholic.or.kr 에서 해당 주일 미사를 열고 주소창 URL을 복사합니다.

## 파일명 규칙

- 최종본: `YYMMDD미사.pptx` — 미사가 있는 **토요일** 날짜, 6자리. 예) 2026년 10월 17일 토요일 → `261017미사.pptx`. 표지 날짜(2026.10.17)로 자동 제안됩니다.
- 성가 이미지: 한 장이면 `128.png`, 여러 장이면 `58-1.png`, `58-2.png` … (png/jpg).

## 저장 위치

- 최종본 PPT: GitHub 저장소의 Releases → **archive** 릴리스 첨부파일. 최근 12주만 남고 오래된 것은 올릴 때 자동 삭제됩니다.
  (GitHub 첨부파일 이름에는 한글을 쓸 수 없어 `261017_missa.pptx`로 저장되고, 앱에서는 `261017미사.pptx`로 보이고 받아집니다.)
- 성가 이미지: 저장소의 `songs/` 폴더.

## 관리자 기능 (성가 관리 탭)

관리자 비밀번호로 로그인 후:
- 성가 번호 검색 → 악보 확인, **교체**(새 이미지 올리기), **삭제**(확인 버튼을 한 번 더 눌러야 삭제).
- 새 성가 올리기: 여러 파일을 한 번에. 파일명(58-1.png 등)으로 번호·순서를 자동 인식하고, 모르면 직접 입력.
- 최종본 PPT 삭제 (확인 후).
모든 변경은 GitHub에 커밋되어 기록이 남습니다.

## 앱이 잠들었을 때

Streamlit 무료 서버는 한동안 접속이 없으면 잠듭니다. 링크를 열면 "Yes, get this app back up!" 버튼이 보이는데, 누르고 1분 정도 기다리면 다시 켜집니다.

## GitHub 토큰 갱신 (관리자)

토큰이 만료되면 앱에 **"GitHub 연결 오류: 토큰이 만료되었거나 권한이 없습니다. 관리자에게 토큰 갱신을 요청하세요."** 가 표시됩니다.

1. GitHub → Settings → Developer settings → Personal access tokens → **Fine-grained tokens** → Generate new token
   - Repository access: Only select repositories → `gaepochurch_ppt`
   - Permissions → Repository permissions → **Contents: Read and write** (이것 하나면 됩니다. Releases 첨부파일도 이 권한으로 동작)
   - 만료 기간을 정하고 생성 → 토큰 복사
2. share.streamlit.io → 앱 오른쪽 메뉴(⋮) → **Settings → Secrets** 에서 `GITHUB_TOKEN = "..."` 값을 새 토큰으로 바꾸고 Save
3. 앱이 자동으로 다시 시작됩니다. 오류가 사라졌는지 확인하고, 예전 토큰은 GitHub에서 삭제합니다.

## 문제가 생겼을 때

- **GitHub 연결 오류**: 위 토큰 갱신.
- **"성가 번호 슬라이드 묶음이 N개입니다"**: 기준 PPT의 성가 슬라이드(그림 1장 + 번호 칸)가 입당/봉헌/성체/파견 4묶음이 아닙니다. 지난 최종본을 확인하거나 다른 주 최종본을 기준으로 고르세요.
- **내용이 비어 있음(빨간 경고)**: 매일미사 페이지에서 찾지 못한 항목입니다. 직접 입력하세요.
- **성가가 ❌**: 해당 번호 이미지가 없습니다. 관리자에게 성가 등록을 요청하세요.
- **최종본 올리기 경고**: 다음 주 자동 생성에 쓰이는 슬라이드(제1독서, 영성체송 등)를 찾지 못했다는 뜻입니다. 저장은 되지만 다음 주 생성 결과를 꼭 확인하세요.

## 개발자용

- Python 3.13 (Streamlit Cloud 배포 시 Advanced settings에서 3.13 선택).
- 설치: `pip install -r requirements-dev.txt` (앱만 실행할 때는 `requirements.txt`)
- 로컬 설정: `.streamlit/secrets.toml.example`을 복사해 `.streamlit/secrets.toml`을 만들고 값을 채웁니다 (커밋 금지). `GITHUB_TOKEN`이 없으면 로컬 `archive/`, `songs/` 폴더를 씁니다.
- 실행: `streamlit run app.py`
- 테스트: `MASS_PPT_ARCHIVE_DIR=<최종본 PPT 폴더> python -m pytest` (폴더가 없으면 PPT가 필요한 테스트는 건너뜀)
