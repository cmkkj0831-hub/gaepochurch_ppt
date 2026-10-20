# 개포동성당 초등부 주일학교 미사 PPT 자동 생성

## 목적
매일미사 페이지의 전례 정보와 교사가 고른 성가 번호로, 지난주 최종본 PPT를 바탕으로 이번 주 미사 PPT를 만드는 웹앱. LLM API는 쓰지 않고 규칙 기반으로만 구현한다. 사용자는 여러 명의 비개발자 교사(폰/PC 브라우저).

## 진행 단계
1. 파서 + 텍스트 수정
2. 성가 교체
3. Streamlit UI (로컬)
4. GitHub 저장
5. 성가 관리 + 용량
6. 배포

## 폴더 구조
```
app.py                  Streamlit 앱 (PPT 만들기 / 최종본 올리기 / 성가 관리)
mass_ppt/
  parser.py             매일미사 페이지 → 필드 dict
  text_edit.py          필드 → 표지·말씀 전례·영성체송 run 교체
  hymns.py              성가 번호 슬라이드 묶음 교체
  generate.py           text_edit + hymns 통합
  pptx_package.py       PPTX zip 읽기/쓰기 공통 헬퍼 (메모리에서만)
  storage.py            Storage 인터페이스, LocalStorage, GitHubStorage
songs/                  성가 이미지 (원본 불변, 저장소에 있음)
archive/                로컬 모드용 최종본 폴더 (.gitignore, 저장소에는 없음)
tests/                  pytest, tests/fixtures/ 에 매일미사 최소 발췌 HTML
requirements.txt        앱 실행용 (버전 고정), requirements-dev.txt 테스트용
```
- Python 3.13. Streamlit Cloud는 파일이 아니라 배포 화면 Advanced settings에서 Python 버전을 고른다.

## 파일명 규칙
- 최종본: `YYMMDD미사.pptx`, 미사가 있는 **토요일** 날짜, 항상 6자리 (예: 261017미사.pptx). "가장 최근 최종본"은 파일명 날짜로 판단.
- 성가 이미지: `songs/{번호}.png` (한 장) 또는 `songs/{번호}-{쪽}.png` (여러 장). png/jpg/jpeg, 대소문자 무관. 쪽 번호 파일이 있으면 그것만 쪽 순서대로 사용.

## 공통 원칙
- **슬라이드는 항상 내용으로 찾는다.** 슬라이드 번호를 고정으로 쓰지 않는다. 슬라이드 순서는 presentation.xml의 sldIdLst 기준(slideN.xml 파일 번호와 다를 수 있음).
- PPT는 bytes로 받고 bytes로 반환. 디스크에 고정 이름 파일을 쓰지 않는다 (여러 명 동시 사용).
- 이미지 재압축·자르기 금지. 최종본·성가 원본 내용 수정 금지.
- 공개 저장소: 토큰·비밀번호·secrets.toml 커밋 금지, 외부 사이트 HTML 전체 커밋 금지(fixture는 필요한 줄만 남긴 최소본).

## 1. 파서 (mass_ppt/parser.py)
- `fetch_html(url)`(User-Agent 지정) / `parse_html(html)` / `parse_mass_url(url)`. 실패 시 `ParseError`(한국어).
- 페이지 전체 텍스트에 정규식 금지(달력·메뉴에 같은 기호가 있음). `form[name=contents]` 안의 `h3` 섹션 제목을 찾고, 각 섹션은 "그 제목부터 다음 h3 전까지"만 본다.
- 날짜: `p.today`("2026년 10월 18일 주일") → `cover_date` = 토요일 "YYYY.MM.DD". URL 파라미터에 의존하지 않는다.
- 제목: `p.m_t_today01`("[(녹) …]") → `liturgy_title`, `cover_title = make_cover_title()`.
- `make_cover_title`: 색 표시 "(녹)" 제거 → 쉼표가 있으면 쉼표 앞 + 문장 끝 괄호 → "제29주일" → "제 29 주일".
  - "(녹) 연중 제29주일, 민족들의 복음화를 위한 미사(전교 주일)" → "연중 제 29 주일(전교 주일)"
- 제1·2독서: 섹션 안 "▥ …말씀입니다." + 바로 뒤 장절. 복음: "✠ …복음입니다." + 장절.
- 화답송: 섹션 안 첫 ◎ 줄, "또는" 이후 무시. 복음 환호송: ○ 줄 ("(◎ 알렐루야.)" 제외).
- 영성체송: 첫 줄(출처 "…참조") 버리고 나머지. "<또는>" 이후 대체 영성체송은 버림.
- 섹션이 없으면 None. 다른 섹션 값으로 대신 채우지 않는다.
- 반환 키: `cover_title, cover_date, liturgy_title, first_reading_intro, first_reading_ref, psalm_refrain, second_reading_intro, second_reading_ref, gospel_acclamation, gospel_intro, gospel_ref, communion_antiphon`.

## 2. 텍스트 수정 (mass_ppt/text_edit.py)
- `apply_text_fields(ppt_bytes, fields) -> (bytes, {'changed', 'warnings'})`.
- 앵커는 "어느 슬라이드인지" 판단에만 쓴다. 앵커가 든 제목 칸은 절대 수정하지 않는다. 수정할 run은 run 텍스트의 시작 기호로 찾는다.
- 대상 슬라이드: 어떤 sp의 텍스트(공백 제거)가 정확히 `제1독서 / 화답송 / 제2독서 / 복음환호송 / 복음 / 영성체송`인 슬라이드. 본문에 "복음"이 들어 있을 뿐인 슬라이드는 대상이 아니다. 각각 정확히 1개여야 하며, 0개나 2개 이상이면 수정하지 않고 경고.
- `<a:t>` 텍스트만 바꾸고 rPr/pPr는 그대로. `text_frame.text = …` 금지. 기존 run의 앞 공백은 유지.
- 표지(첫 슬라이드): 제목 run(앞 공백 유지), `\d{4}\.\d{2}\.\d{2}` 날짜 run.
- 제1·2독서: "▥"로 시작하는 run → 출처, 바로 다음 run → 장절. "...", "주님의 말씀입니다", "◎ 하느님, 감사합니다"는 그대로.
- 화답송: "◎" run 하나. 복음 환호송: "○" run 하나("◎ 알렐루야"는 그대로).
- 복음: "✠" 포함 run → 출처(앞 공백 유지), 바로 다음 run → 장절.
- 영성체송: 제목 칸과 "영성체 예식" 칸을 뺀 본문 칸.
- 값이 None이면 해당 항목은 수정하지 않고 경고.

## 3. 성가 교체 (mass_ppt/hymns.py)
- 성가 번호 슬라이드 = spTree에 `p:pic` 정확히 1개 + 텍스트가 숫자뿐인 `p:sp` 정확히 1개, 그 외 도형 없음. 그 외 이미지 슬라이드(대영광송, 배경 등)는 절대 건드리지 않는다.
- 연속된 성가 번호 슬라이드 = 한 묶음 (번호가 달라도 연속이면 같은 묶음).
- 템플릿에는 묶음이 정확히 4개: 순서대로 입당 / 봉헌 / 성체 / 파견.
  - 봉헌 묶음 바로 다음 = "예물 준비 기도" 슬라이드, 성체 묶음 바로 앞 = 영성체송 슬라이드.
  - 묶음 수가 4가 아니거나 검증 실패 시 `HymnTemplateError`로 생성 중단.
- 영성체송 슬라이드는 영성체 기도문("합당치 않사오나") 바로 다음, 성체 성가보다 앞. 어긋나면 먼저 재배치하고 경고.
- 각 묶음을 지우고 같은 자리에, 성가마다 이미지 장수만큼 슬라이드를 넣는다 (58-1, 58-2 … / 128.png는 1장). 성체 2곡이면 1곡 다음 2곡.
- 새 슬라이드는 묶음의 기존 성가 슬라이드를 복제하고 숫자와 이미지만 교체:
  - 번호 칸: off(0,0) ext(1567543,1015663), 60pt bold, 가운데, 그림자 — 템플릿 서식 그대로.
  - 이미지 영역: x=1567543 ~ 슬라이드 끝(12192000), y=0 ~ 6858000. 영역 크기 기준 비율 계산으로 최대 크기, 가로·세로 가운데 (`place_image`). 자르기(srcRect)·재압축 금지.
- 이미지가 없는 성가는 빼고 경고.
- 갱신 대상: presentation.xml sldIdLst와 p14 sectionLst, presentation.xml.rels, slide rels, [Content_Types].xml(slide Override, png/jpg/jpeg Default), docProps/app.xml 슬라이드 수. 지운 슬라이드(및 notes)와 고아 media 삭제. zip은 [Content_Types].xml을 맨 앞에, 원본 순서·압축 방식 유지.

## 4. 이미지 가져오기와 generate
- `Storage.get_hymn_images(num) -> list[bytes]`: 해당 번호 모든 장을 순서대로. 로컬·GitHub 모두 storage를 통해.
- `generate(base_ppt_bytes, fields, hymn_numbers, get_hymn_images) -> (bytes, report)`:
  - URL을 다시 파싱하지 않고 화면에서 확인·수정한 fields만 사용.
  - `hymn_numbers = {'entrance': 58, 'offertory': 106, 'communion': [128, 129], 'recessional': 61}`
  - report: `changed_texts, hymns[{role, num, pages}], slide_count, warnings` (None이 있어도 오류 없음).

## Streamlit 앱 (app.py)
- `layout="centered"`, 한국어, 한 줄에 하나씩(st.columns 최소화). 사용자 데이터는 st.session_state에만, 생성 파일은 메모리에서 바로 다운로드.
- PPT 만들기: 기준 PPT(기본 최신 최종본 / 다른 주 선택 / 직접 업로드) → URL 불러오기 → 수정 칸(표지 제목, 표지 날짜, 독서·복음 출처+장절, 화답송, 복음환호송, 영성체송) → 성가 번호(✅/❌ + 장수) → 생성·다운로드(YYMMDD미사.pptx).
- 최종본 올리기: 구조 점검(성가 묶음, 텍스트 슬라이드, 표지 날짜) 실패 시 경고만 하고 업로드 허용. 표지 날짜로 파일명 제안, 규칙 안 맞으면 저장 불가, 같은 날짜는 덮어쓰기 확인. 이 페이지에는 삭제 기능 없음.
- 성가 관리: ADMIN_PASSWORD(st.secrets)로 보호. 다른 페이지는 비밀번호 없음. 번호 검색 → 악보·교체·삭제, 새 성가 여러 장 업로드(파일명 인식 실패 시 직접 입력), 최종본 삭제. 삭제는 버튼 → 확인 버튼 2단계(session_state의 pending 값).
- GitHub 오류(StorageError)는 각 탭에서 한국어로 표시하고 앱은 죽지 않는다.

## 저장소 (storage.py)
- secrets에 GITHUB_TOKEN이 있으면 GitHubStorage, 없으면 LocalStorage(로컬 archive/, songs/). 토큰은 st.secrets에서만 읽는다. `get_storage`는 토큰을 캐시 키에 포함해서 토큰이 바뀌면 새로 만든다.
- **최종본 PPT: GitHub Releases** — 릴리스 하나(태그 `archive`)의 첨부파일.
  - GitHub 첨부파일 이름에서 한글이 지워지므로 저장 이름은 `YYMMDD_missa.pptx`, 앱에서 보이는 이름은 `YYMMDD미사.pptx`. 올린 사람은 첨부파일 label("올린 사람: OOO")에 저장.
  - 같은 이름은 지우고 다시 올림(교체). 올린 뒤 파일명 날짜 기준 **최근 12개만 남기고** 나머지 자동 삭제.
  - 목록·최신본 선택·다운로드 모두 Releases 기준.
- **성가 이미지: 저장소 songs/** — Git Data API(blob → base_tree로 tree → commit → ref). ref 갱신이 충돌(422)하면 최신 HEAD로 다시 만들어 재시도. 추가·교체·삭제는 `change_songs(add, delete, message)` 한 번의 커밋.
  - 커밋 메시지: "성가 추가: 58-1.png", "성가 교체: 58", "성가 삭제: 58".
- 읽기 캐시: 목록은 30초 TTL, 파일 내용은 blob sha·asset id 키(내용이 바뀌면 키가 바뀜). 쓰기 직후 캐시 비움.
- 오류: 401/403(사용량 한도 제외)은 `GitHubAuthError` — 앱은 "GitHub 연결 오류: 토큰이 만료되었거나 권한이 없습니다. 관리자에게 토큰 갱신을 요청하세요."를 그대로 표시. 그 외 실패도 `StorageError`(한국어)로 올리고, 빈 목록으로 숨기지 않는다.
- 토큰: Fine-grained, 저장소 gaepochurch_ppt 하나, **Contents: Read and write**만 (Releases 첨부파일도 이 권한으로 동작 확인).
- **토큰 만료 시 갱신**: GitHub에서 새 Fine-grained 토큰 발급(위 권한) → share.streamlit.io 앱 Settings → Secrets의 `GITHUB_TOKEN` 교체 → 앱 재시작 확인 → 예전 토큰 삭제. (README에도 같은 절차)

## 용량
- 최종본은 Releases에 최근 12개(약 31MB × 12)만 보관하므로 git 저장소 크기와 무관. 저장소에는 코드와 songs/만 남는다. (이전 커밋 기록에 옛 archive/ PPT가 남아 있지만 기록 정리는 하지 않기로 함.)

## 테스트
- `MASS_PPT_ARCHIVE_DIR=<최종본 폴더> ppt_env/bin/python -m pytest` (오프라인). 최종본 PPT가 저장소에 없으므로 폴더를 지정하고, 없으면 해당 테스트는 skip. 파서는 tests/fixtures/의 최소 발췌 HTML로 검증. GitHubStorage는 가짜 API로 검증(tests/test_storage.py).
- 텍스트 수정은 run 단위 전체 diff로 지정 run 외 변화가 없는지 확인.
- 시각 확인은 LibreOffice로 PDF 변환 후 이미지로 본다(한글 글꼴 대체 필요).

## Git
- push 전 토큰·비밀번호 문자열 검사. force push 금지.
