# 천주교 초등부 주일학교 미사 PPT 자동 생성 프로젝트

## 프로젝트 목적
매주 천주교 매일미사 웹사이트에서 전례 정보를 파싱하여 개포동성당 초등부 주일학교 미사 PPT를 자동으로 생성하는 웹앱.

## 전체 진행 단계
1. **1단계**: 파서(parser.py) + 텍스트 수정기(text_edit.py) 구현 및 검증
2. **2단계**: 성가 규칙 정의 및 성가 이미지 삽입 기능
3. **3단계**: 웹앱 백엔드 (Flask/FastAPI)
4. **4단계**: 웹앱 프론트엔드
5. **5단계**: 데이터베이스 & 스케줄링
6. **6단계**: 배포 & 테스트

## 폴더 구조
```
gaepochurch_ppt/
├── CLAUDE.md                 # 이 파일
├── mass_ppt/                 # 핵심 로직 패키지
│   ├── __init__.py
│   ├── parser.py            # URL 파싱
│   ├── text_edit.py         # PPT 텍스트 수정
│   ├── hymns.py             # 성가 이미지 교체
│   ├── generate.py          # 전체 생성 함수
│   └── models.py            # 데이터 모델
├── songs/                    # 성가 이미지 (수정금지)
├── archive/                  # 최종본 PPT (수정금지)
└── tests/                    # 단위 테스트
    ├── __init__.py
    ├── test_parser.py
    ├── test_text_edit.py
    ├── test_hymns.py
    └── test_generate.py
```

## 파일명 규칙
- **최종본 PPT 파일명**: `YYMMDD미사.pptx`
  - `YY`: 연도 뒤 2자리 (예: 26 = 2026)
  - `MM`: 월 2자리
  - `DD`: 미사가 있는 토요일 기준 (예: 261017미사.pptx → 2026년 10월 17일)
  - 항상 6자리 고정

## 구현 규칙

### URL 파싱 (mass_ppt/parser.py)
- **입력**: 매일미사 웹사이트 URL
  - 예: `https://maria.catholic.or.kr/mi_pr/missa/missa.asp?menu=missa&missaid=16874&gomonth=2026-10-18&missatype=DA`
- **기술**: requests + BeautifulSoup (LLM API 미사용, 규칙 기반만)
- **User-Agent 필수**: 웹사이트 접근 시 반드시 User-Agent 헤더 지정

**추출 항목**:
1. **전례일 제목**: 예) "(녹) 연중 제29주일, 민족들의 복음화를 위한 미사(전교 주일)"
2. **날짜**: URL의 gomonth 파라미터에서 추출
3. **제1독서 출처**: "▥ 이사야서의 말씀입니다." + "2,1-5" (두 부분을 각각 추출)
4. **화답송 후렴**: 첫 번째 `◎` 문장만 (있으면), "또는" 뒤는 무시
5. **제2독서 출처**: "▥ 사도 바오로의 필리피서 말씀입니다." + "4,12-14.19-20"
6. **복음 환호송의 ○ 구절**: "우리 주 예수 그리스도의 아버지..." 부분
7. **복음 출처**: "✠ 마태오가 전한 거룩한 복음입니다." + "28,16-20"
8. **영성체송 본문**: 출처 줄 제외한 본문만

**중요**: 독서 본문(말씀) 자체는 추출하지 않음. PPT의 본문 위치는 "..."로 두는 것이 원래 형식.

**반환값**: dict (없는 항목은 None)

### 텍스트 수정 (mass_ppt/text_edit.py)
- **PPT 접근**: 앵커 텍스트로 슬라이드 찾기 (고정 슬라이드 번호 사용 금지)
  - 표지: "연중 제" 또는 "주일" 포함 첫 번째 슬라이드
  - "제1독서" / "화답송" / "제2독서" / "복 음 환 호 송" / "복음" / "영 성 체 송" 등으로 검색
  
- **서식 보존**: 최우선 원칙
  - **반드시** `<a:r>` run 단위로 텍스트만 변경
  - `rPr` (run properties), `pPr` (paragraph properties) 유지
  - python-pptx의 `text_frame.text = ...` 절대 금지 (서식 삭제됨)
  - lxml 직접 조작 권장

- **표지 날짜**: 형식 "2026.10.17" (토요일)
  - 입력된 주일 날짜(일요일)에서 하루를 빼서 토요일 계산
  - 기본값으로 사용

- **표지 제목**: 형식 예) "연중 제 29 주일(전교 주일)"
  - 숫자 앞뒤에 공백 필수
  - 괄호는 그대로 유지

- **함수 시그니처**:
  ```python
  def update_ppt(ppt_bytes: bytes, data: dict) -> bytes:
      """
      PPT를 bytes로 받아 수정 후 bytes로 반환.
      작업 폴더는 tempfile로 생성/삭제.
      """
  ```

## 슬라이드 위치 (현재 템플릿 기준)
- 슬라이드 1: 표지 (제목 + 날짜)
- 슬라이드 27: 제1독서
- 슬라이드 28: 화답송
- 슬라이드 29: 제2독서
- 슬라이드 30: 복음 환호송
- 슬라이드 31: 복음
- 슬라이드 56: 영성체송

## 2단계: 성가 이미지 교체 (mass_ppt/hymns.py, mass_ppt/generate.py)

### 성가 번호 슬라이드 정의 (가장 중요)
- **성가 번호 슬라이드**: 그림 1개 + 숫자만 적힌 텍스트박스 1개로 이루어진 슬라이드
  - 이 슬라이드만 교체 대상
  - 그 외 이미지 슬라이드(대영광송, 기도문 등)는 절대 건드리지 않음
  
### 성가 배치 위치 (순서 중요)
성가 번호 슬라이드 묶음은 **순서대로**:
1. **입당** (Entrance): 첫 번째 묶음
2. **봉헌** (Offertory): "예물 준비 기도" 앞
3. **성체** (Communion): "영 성 체 송" 뒤
4. **파견** (Recessional): 성체 이후

### 성가 슬라이드 교체 규칙
- **기존 슬라이드 처리**: 기존 묶음 완전히 제거
- **새 슬라이드 삽입**: 같은 자리에 새 슬라이드 삽입 (한 장당 한 슬라이드)
- **입력 포맷**: 
  - 입당: 1개 (예: 58)
  - 봉헌: 1개 (예: 106)
  - 성체: 1~2개 (예: 128, 129)
  - 파견: 1개 (예: 61)

### 이미지 파일명 규칙
- **형식**: `songs/{번호}-{장}.png` 또는 한 장짜리면 `songs/{번호}.png`
  - 예: `58-1.png`, `58-2.png`, `128.png`
- **확장자**: `.png` 또는 `.jpg` 모두 허용
- **없는 번호**: 이미지 파일이 없으면 슬라이드 삽입 안 하고 경고 목록에 추가

### 슬라이드 배치 규칙 (EMU 좌표)
슬라이드 기본 크기: 12192000 x 6858000 EMU (10" x 5.625")

**왼쪽 위 번호 텍스트박스**:
- 위치(offset): (0, 0)
- 크기(extent): (1567543, 1015663)
- 포맷: 60pt bold, 가운데 정렬
- 그림자 효과: 기존 템플릿 그대로 유지

**이미지 영역**:
- x 좌표: 1567543부터 오른쪽 끝(12192000)까지
- 비율 유지하며 최대 크기로 맞추기
- 영역 안에서 가로·세로 가운데 정렬
- **원본 이미지 자르지 않음** (전체 이미지 표시)

### 슬라이드 복제 방법
- **템플릿**: 기존 성가 번호 슬라이드 XML 사용
- **파일 갱신**:
  - `presentation.xml`의 `sldIdLst` 갱신
  - `ppt/slides/_rels/*.xml` 관계 파일 갱신
  - `[Content_Types].xml` 타입 정보 갱신
  - 고아 media 파일 정리

### 고정 규칙
**영성체송 슬라이드 배치**:
- 영성체송 슬라이드는 반드시 영성체 기도문("합당치 않사오나") 슬라이드 바로 다음에 와야 함
- 성체성가보다 앞에 배치해야 함
- 어긋나 있으면 자동으로 재배치

### 전체 생성 함수 (mass_ppt/generate.py)
```python
def generate(base_ppt_bytes: bytes, 
             parsed_fields: dict,
             hymn_numbers: dict,
             get_hymn_image: callable) -> tuple[bytes, dict]:
    """
    PPT 생성.
    
    Args:
        base_ppt_bytes: 템플릿 PPT bytes
        parsed_fields: parser.parse_mass_url의 결과
        hymn_numbers: {'entrance': 58, 'offertory': 106, 'communion': [128, 129], 'recessional': 61}
        get_hymn_image: callable(hymn_number: int) -> bytes | None
                        성가 번호 → 이미지 bytes (없으면 None)
    
    Returns:
        (result_bytes, report_dict)
        report: {
            'updated_texts': [...],
            'inserted_hymns': [{'num': 58, 'position': 'entrance', 'slides_added': 1}, ...],
            'warnings': ['Hymn image not found: 999', ...]
        }
    """
```

**report 내용**:
- `updated_texts`: 변경된 텍스트 항목 리스트
- `inserted_hymns`: 삽입된 성가 정보 (번호, 위치, 슬라이드 수)
- `warnings`: 주의사항 (없는 이미지, 배치 변경 등)

## Git 커밋 규칙
- Co-Author 포함: `Co-Authored-By: Claude Haiku 4.5 <noreply@anthropic.com>`
- 명확한 메시지: 무엇(what)보다는 왜(why) 중심

## 테스트
- pytest 사용
- 파서: URL 파싱 결과 검증
- 텍스트 수정: 앵커 탐지 및 서식 보존 검증
