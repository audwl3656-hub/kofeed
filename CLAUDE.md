# CLAUDE.md — 숙련도 시험 플랫폼 (Kofeed)

## 프로젝트 개요

한국 사료 업계의 **회원사 비교분석 시험(숙련도 시험) 플랫폼**으로 Streamlit으로 구축되어 있습니다. 회원사들이 사료 시료의 분석 데이터를 제출하면, 시스템이 Robust Z-score를 계산하여 기관 간 비교를 수행하고 PDF 보고서를 이메일로 발송합니다.

앱은 **Streamlit Community Cloud**에 배포되며, **Google Sheets**를 데이터베이스 및 설정 저장소로 사용합니다.

---

## 저장소 구조

```
kofeed/
├── app.py                  # 참가자 데이터 제출 폼 (진입점)
├── pages/
│   └── admin.py            # 관리자 대시보드 (비밀번호 보호)
├── utils/
│   ├── __init__.py
│   ├── config.py           # 전체 설정 읽기/쓰기; Google Sheets "config" 탭
│   ├── sheets.py           # 제출 데이터 읽기/쓰기; Google Sheets "제출데이터" 탭
│   ├── zscore.py           # Robust Z-score 계산 헬퍼
│   ├── report.py           # PDF 생성 (ReportLab, 한글 폰트)
│   └── email_sender.py     # SMTP 이메일 발송 (Gmail)
├── requirements.txt        # Python 의존성 패키지
├── packages.txt            # 시스템 패키지 (한글 PDF용 fonts-nanum)
└── .gitignore
```

---

## 기술 스택

| 계층 | 기술 |
|---|---|
| UI 프레임워크 | Streamlit ≥ 1.35 |
| 데이터 저장소 | Google Sheets (gspread ≥ 6.0) |
| 인증 | Google 서비스 계정 (`st.secrets["gcp_service_account"]`) |
| PDF 생성 | ReportLab ≥ 4.2 |
| 한글 폰트 | NanumGothic TTF (`packages.txt`: `fonts-nanum`) |
| 이메일 | Gmail SMTP SSL 465 포트 |
| 데이터 처리 | pandas ≥ 2.0, numpy ≥ 1.26 |

---

## 시크릿 설정

시크릿은 `.streamlit/secrets.toml`에 저장됩니다 (gitignore 처리됨). 필수 키:

```toml
[gcp_service_account]
# Google 서비스 계정 JSON 전체 내용을 TOML 키로 입력

[sheet]
name = "스프레드시트 파일명"   # Google Sheets 파일명

[admin]
password = "관리자_비밀번호"

[email]
sender   = "sender@gmail.com"
password = "gmail_앱_비밀번호"
```

---

## Google Sheets 구조

두 개의 워크시트를 사용합니다.

### `config` 탭
모든 앱 설정이 저장됩니다. 컬럼: `type | group | name | samples | order | enabled | use_equip | use_solvent | free_decimal`

설정 행 타입:

| `type` | 용도 | 주요 필드 |
|---|---|---|
| `method_option` | 분석 방법 드롭다운 선택지 | `name`=방법명, `group`=성분명 (비워두면 전체 공통) |
| `info_field` | 참가자 정보 폼 필드 | `name`=필드명, `group`=placeholder, `samples`=플래그 (`required`, `email`) |
| `sample` | 사료 종류 | `name`=사료명 |
| `group` | 성분 섹션 | `name`=섹션명, `samples`=`nir` (NIR 대상인 경우) |
| `component` | 개별 분석 항목 | `group`=섹션명, `name`=성분명, `samples`=`all` 또는 쉼표로 구분된 사료명 |
| `question` | 추가 설문 항목 | `group`=질문 ID, `name`=질문 내용, `samples`=유형 명세 (아래 참고) |
| `participant` | 참가코드 → 회사명 매핑 | `group`=코드, `name`=회사명 |

질문 `samples` 유형 명세:
- `text` 또는 `text:힌트` → 주관식
- `choice:옵션1|옵션2|옵션3` → 단일 선택 (라디오)
- `multicheck:옵션1|옵션2|옵션3` → 복수 선택 (체크박스)

### `제출데이터` 탭
제출 행이 저장됩니다. 헤더 행은 첫 제출 시 자동 생성되며, config에 새 성분이 추가되면 이후 제출 시 자동으로 열이 추가됩니다. 컬럼 명명 규칙:
- `{성분}_{사료}` — 수치값 (예: `수분_축우사료`)
- `{성분}_방법` — 분석 방법 문자열
- `{성분}_기기` — 기기명
- `{성분}_용매` — 용매명
- `NIR_{성분}_{사료}` — NIR 측정값

---

## 앱 흐름

### 참가자 폼 (`app.py`)
1. config에 `participant` 항목이 있으면 폼 진입 전에 참가코드 입력 게이트를 표시합니다.
2. Google Sheets에서 설정을 로드합니다 (`@st.cache_data(ttl=120)`로 120초 캐시).
3. 기관 정보 필드, 그룹/섹션별 성분 입력 테이블, 선택적 설문 항목을 렌더링합니다.
4. 제출 시: 필수 필드 및 방법 선택 규칙을 검증한 후 `submit_data()`로 Google Sheets에 행을 추가하고, 제출 확인 이메일(PDF 첨부)을 즉시 발송합니다.

성분 입력 테이블 컬럼: **성분 | 방법 (드롭다운) | 기기명 | 용매 | [사료 종류 컬럼들...]**

### 관리자 대시보드 (`pages/admin.py`)
`st.secrets["admin"]["password"]`로 보호됩니다. 네 개의 탭:
- **제출 현황** — 전체 제출 데이터를 데이터프레임으로 조회
- **Z-score 분석** — 성분별 Robust Z-score 분석 (전체 및 방법별), 색상 코딩 테이블
- **보고서 발송** — 개별 PDF 보고서 생성 (전체 + 방법별), 다운로드 버튼 및 일괄 이메일 발송
- **설정** — `st.data_editor`를 통한 모든 설정 타입 CRUD

---

## Robust Z-score (`utils/zscore.py`)

공식: `Z = (x − 중앙값) / (1.4826 × MAD)`

- MAD = 0인 경우 표준편차 기반 Z-score로 폴백합니다.
- Z-score 계산에 **최소 n > 5**개 기관이 필요합니다 (미만이면 `NaN` 반환).
- 방법별 Z-score도 동일 방법을 사용하는 기관이 5개 초과인 경우에만 계산합니다.
- 판정 기준: `|Z| ≤ 2.0` = 적합, `2.0 < |Z| ≤ 3.0` = 경고, `|Z| > 3.0` = 부적합.

---

## PDF 보고서 (`utils/report.py`)

세 가지 보고서 유형:
- **전체 Z-score** (`generate_pdf_overall`) — 전체 기관 대비 비교.
- **방법별 Z-score** (`generate_pdf_by_method`) — 동일 분석 방법 사용 기관 내 비교.
- **제출 확인서** (`generate_submission_pdf`) — 제출 즉시 발송; Z-score 없음.

한글 폰트: `/usr/share/fonts/truetype/nanum/NanumGothic.ttf` 로드. 로컬 개발 환경에서 `fonts-nanum`이 설치되지 않은 경우 Helvetica로 폴백합니다 (한글 깨질 수 있음).

---

## 주요 규칙 및 컨벤션

### Session State 키 패턴
성분 입력 위젯의 키는 `{그룹명}_{성분명}_{사료명}` 형식을 사용합니다 (예: `일반성분_수분_축우사료`). 폼 제출 시 `st.session_state`에서 직접 값을 읽어 stale data 문제를 방지합니다.

### 컬럼명 규칙
- `is_value_col(col, samples)` — 컬럼명이 `_{사료명}`으로 끝나면 True 반환.
- `get_component_from_col(col, samples)` — 사료 suffix (및 `NIR_` prefix) 제거 후 성분명 반환.
- `get_sample_from_col(col, samples)` — 컬럼명에서 사료명 부분 반환.

### 기관명 필드 감지
기관/회사명 필드는 키워드 매칭으로 찾습니다: `info_field` 중 `name`에 `기관`, `회사`, `기업`, `업체` 중 하나를 포함하는 첫 번째 필드. 없으면 첫 번째 info_field를 기본값으로 사용합니다.

### 캐시 무효화
- `get_config()`는 `ttl=120` 초.
- 관리자의 `load_data()`는 `ttl=60` 초.
- `st.cache_data.clear()` 또는 `get_config.clear()` 호출로 강제 갱신 (설정 저장 시 자동 수행됨).

### 불리언 컬럼 처리
`use_equip`, `use_solvent`, `free_decimal` 컬럼은 기본값이 각각 `True`/`True`/`False`이며, 기존 시트에 해당 컬럼이 없어도 graceful하게 처리됩니다.

---

## 개발 워크플로

### 로컬 실행

```bash
pip install -r requirements.txt
# 시스템 폰트 설치 (Linux):
sudo apt-get install fonts-nanum

# .streamlit/secrets.toml 파일을 위의 필수 키 형식에 맞게 작성
streamlit run app.py
```

### 새 성분 추가
1. 관리자 **설정** 탭의 **성분 종목**에 행 추가 (`group`, `name`, `samples`, 옵션 플래그 입력).
2. **설정 저장** 클릭 — config 시트가 업데이트되고 2분 내에 폼에 반영됩니다 (캐시 TTL).
3. 새 제출 컬럼은 다음 제출 시 데이터 시트에 자동으로 추가됩니다.

### 새 사료 종류 추가
1. 관리자 설정의 **사료 종류**에 행 추가.
2. `all`로 커버되지 않는 성분 항목의 `samples` 필드를 필요시 업데이트.

### PDF 레이아웃 수정
`utils/report.py` 수정. `_make_table_style()` 함수에서 표 스타일을 일괄 관리합니다. 열 너비는 `mm` 단위입니다.

### 이메일 설정
Gmail App Password(계정 비밀번호가 아님)를 사용합니다. 시크릿에 `email.sender`와 `email.password`를 설정하세요. 첨부 파일의 한글 파일명은 네이버/다음 메일 호환을 위해 RFC 2047 base64로 인코딩됩니다.

---

## 배포 (Streamlit Community Cloud)

- 모든 시크릿은 Streamlit Cloud 시크릿 관리자에 설정합니다 (git에 커밋하지 않음).
- `packages.txt`가 한글 PDF 렌더링을 위한 시스템 레벨 `fonts-nanum`을 설치합니다.
- 연결된 브랜치에 푸시할 때마다 앱이 자동 재시작됩니다.
- Google 서비스 계정은 대상 스프레드시트에 대한 **편집자** 권한이 필요합니다.
