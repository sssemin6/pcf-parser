# PCF 파서 모듈 상세 PRD (Product Requirements Document)

## 1. 목적 및 범위

* **목적**: 사내 실적 선주·선급 교신 PCF(Plan Comments Form) 파일(`.xlsx`, `.docx`)을 일관된 **`Comment Thread` 단위 데이터**로 추출·정규화하여, 향후 RAG 검색 및 선주 회신 초안 생성 웹앱의 기반 DB(SQLite)를 구축한다.
* **핵심 전략 (Hybrid AI 검토)**:
  * **1차(Rule-based)**: Python(`openpyxl`, `python-docx`)이 85~90%의 정형 데이터를 빠르고 정확하게 구조화하여 고유번호(`Hull No.`, `Plan No.`) 환각 방지.
  * **2차(Selective AI)**: 파싱 신뢰도가 낮은 예외 건 복구(Fallback)와 스레드 단위 핵심 요약·결론 태깅(Enrichment)만 선별적으로 수행.
* **범위 (Phase 1)**: 로컬 폴더 내 PCF 파일 스캔 → 파싱 → 중간 AI 검증/정제 → SQLite 저장 및 파싱 품질 리포트 출력. *(웹 UI 및 회신 생성 프롬프트는 Phase 2에서 구현)*

---

## 2. 입력 파일 스펙 및 지원 정책

| 구분 | 지원 포맷 | 처리 정책 |
| :--- | :--- | :--- |
| **1차 지원** | `.xlsx`, `.docx` | `openpyxl`, `python-docx`로 직접 파싱 |
| **구버전 포맷** | `.xls`, `.doc` | 파싱 전 LibreOffice(Headless) 또는 Windows COM(`win32com`)으로 `.xlsx`/`.docx` 일괄 변환 후 파이프라인 투입 |
| **제외 대상** | 임시 파일(`~$*`), 암호화 파일, 스캔 PDF | `skipped_files` 로그 테이블에 사유 기록 후 건너뜀 |

---

## 3. 파싱 파이프라인 아키텍처 (4-Stage)

### Stage 1: 파일 스캔 및 증분 처리
* 지정된 루트 디렉터리를 재귀 탐색하며 파일별 `SHA-256` 해시 계산.
* DB(`pcf_documents.file_hash`)와 비교해 **신규 또는 수정된 파일만 증분 파싱**.
* 파일명/경로에서 보조 메타데이터(선종, 호선번호, 도면번호 등)를 1차 정규식으로 추출해 문서 내부 파싱 실패 시 백업 값으로 활용.

### Stage 2: 규칙 기반 구조화 파싱 (Deterministic Parser)

#### A. 공통 헤더 메타데이터 추출
* **대상 필드**: `hull_no`, `ship_type`, `plan_title`, `plan_no`, `comment_party` (Buyer / Class 구분), `approval_status` (Approved / Approved with comments 등).
* **추출 방식**: 고정 셀 좌표(예: `B3`)를 하드코딩하지 않고, 상단 15행(Excel) 또는 첫 번째 표(Word)에서 `Plan Title`, `Plan No.`, `Hull No.` 키워드 셀을 찾은 뒤 **우측 인접 셀(또는 병합된 다음 셀)** 값을 동적 추출.

#### B. 본문 테이블 파싱 및 스레드 그룹화
* **컬럼 동적 매핑**: 헤더 행에서 정규식 매칭으로 열 인덱스 자동 식별.
  * `no_col`: `No\.?`, `Item`
  * `buyer_col`: `BUYER'?S?\s*COMMENT`, `OWNER'?S?\s*COMMENT`, `CLASS\s*COMMENT`
  * `builder_col`: `BUILDER'?S?\s*REPLY`, `YARD'?S?\s*REPLY`
  * `status_col`: `STATUS`, `REMARK`
  * `round_col`: 별도 열이 없으면 본문 텍스트 선두의 `a)`, `b)`, `1st`, `2nd` 패턴으로 분리
* **Excel (`.xlsx`) 병합 셀 처리**:
  1. `ws.merged_cells.ranges`를 순회하며 모든 병합 영역의 좌상단 값을 하위 셀로 전개(Unmerge & Forward-Fill).
  2. `No.` 값이 동일한 연속 행을 하나의 `Thread`로 묶고, 각 행의 `round` 순서대로 `Turn` 리스트 구성.
* **Word (`.docx`) 병합 셀 처리**:
  1. 본문 표(`tables[1]` 등 컬럼 매핑에 성공한 표)의 행을 순회.
  2. `No.` 셀의 XML 객체(`cell._tc`)가 이전 행과 동일하거나 빈 값이면 기존 `Thread`에 병합, 변경되면 새 `Thread` 시작.
  3. 한 셀 안에 `a) ... \n b) ...`가 줄바꿈으로 묶여 있는 경우, 정규식(`^(?:[a-z]\)|[1-9](?:st|nd|rd|th))`)으로 턴 분리.
* **이미지 및 OLE 첨부파일 태깅**:
  * 셀 내부에 삽입된 이미지(`w:drawing`, `openpyxl` 이미지 앵커)나 임베디드 객체(`Attachment #1-a.pdf` 등)를 감지해 해당 턴 텍스트 끝에 `[IMAGE_ATTACHED]`, `[FILE_ATTACHED: 파일명]` 태그 추가.

### Stage 3: 중간 AI 검토 및 정제 (Selective AI Layer)
로컬 Ollama 또는 사내/클라우드 LLM API를 선택할 수 있도록 인터페이스를 추상화하며, **2가지 역할**로 분리해 실행한다.

1. **예외 복구용 AI (Fallback — 조건부 호출)**
   * **트리거 조건**:
     * 헤더에서 `Plan Title` 또는 `Hull No.` 추출 실패
     * 본문 테이블 헤더 매핑 실패 (비정형 양식)
     * `buyer_comment`와 `builder_reply`의 턴 개수가 맞지 않거나, 단일 셀에 500자 이상 텍스트와 여러 라운드가 뒤섞여 정규식 분리 실패
   * **동작**: 해당 표의 원시 텍스트(Markdown/CSV 형태)만 LLM에 전달해 표준 JSON 스키마로 복구. 복구된 건은 `parse_method = 'llm_fallback'`으로 마킹.
2. **스레드 정제용 AI (Enrichment — 스레드당 1회 호출)**
   * 파싱이 완료된 각 `Thread` 전체(1st~최종)를 입력받아 아래 4개 필드를 JSON으로 추출:
     * `issue_summary`: 선주 최초 요구사항 및 핵심 기술 쟁점 (한글/영문 1~2문장)
     * `final_conclusion`: 최종 합의/종결 내용 (예: *"1호선 Gas trial 시에만 시험 수행하기로 합의"*)
     * `inferred_status`: 표의 `STATUS`가 비어있거나 애매할 때 마지막 턴(`Noted, thanks`, `Confirmed/Agreed` 등)을 보고 `Closed` / `Open` 판정
     * `search_keywords`: 핵심 장비명, 규정(SOLAS, LSA 등), 시험 항목(FAT, Sea trial 등) 태그
   * **효과**: `Waiting confirmation`, `Can Builder understand...` 같은 행정성 노이즈를 배제하고, 검색용 텍스트(`search_text`)의 품질을 극대화함.

### Stage 4: 저장 및 검증 리포트
* SQLite DB에 트랜잭션 단위로 저장.
* 파싱 완료 후 콘솔/로그에 **파싱 요약 리포트** 출력:
  * 전체 파일 수 / 성공 / AI Fallback 복구 건수 / 실패 파일 목록
  * 추출된 총 Thread 수 / Turn 수 / `STATUS` 미식별 건수

---

## 4. 데이터베이스 스키마 (SQLite)

Parent-Child(Small-to-Big) RAG 구조를 지원하기 위해 검색용 텍스트(`search_text`, Child)와 LLM 전달용 전체 스레드(`pcf_threads` + `pcf_turns`, Parent)를 분리 설계한다.

```sql
-- 1. PCF 문서 메타데이터
CREATE TABLE pcf_documents (
    doc_id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path TEXT UNIQUE NOT NULL,
    file_hash TEXT NOT NULL,
    file_type TEXT NOT NULL,          -- 'xlsx' | 'docx'
    hull_no TEXT,
    ship_type TEXT,
    plan_title TEXT,
    plan_no TEXT,
    comment_party TEXT,               -- 'Buyer' | 'Class'
    parse_method TEXT,                -- 'rule' | 'llm_fallback'
    parsed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Comment No. 단위 스레드 (LLM 전달 단위 = Parent)
CREATE TABLE pcf_threads (
    thread_id INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_id INTEGER NOT NULL REFERENCES pcf_documents(doc_id) ON DELETE CASCADE,
    comment_no TEXT NOT NULL,
    initial TEXT,
    raw_status TEXT,                  -- 원문 STATUS
    inferred_status TEXT,             -- AI 보정 STATUS ('Closed' | 'Open' | 'Unknown')
    issue_summary TEXT,               -- AI 요약: 핵심 쟁점
    final_conclusion TEXT,            -- AI 요약: 최종 결론
    search_keywords TEXT,             -- AI 추출 키워드 (JSON 배열)
    search_text TEXT,                 -- FTS5 및 임베딩 대상 정제 텍스트 (Child)
    confidence_score REAL DEFAULT 1.0 -- 파싱 신뢰도 (0.0 ~ 1.0)
);

-- 3. 개별 교신 턴 (1st/2nd 또는 a/b)
CREATE TABLE pcf_turns (
    turn_id INTEGER PRIMARY KEY AUTOINCREMENT,
    thread_id INTEGER NOT NULL REFERENCES pcf_threads(thread_id) ON DELETE CASCADE,
    turn_order INTEGER NOT NULL,
    round_label TEXT,                 -- '1st', '2nd', 'a)', 'b)' 등
    buyer_comment TEXT,
    builder_reply TEXT,
    has_image BOOLEAN DEFAULT 0,
    attachments TEXT                  -- 첨부파일명 목록 (JSON 배열)
);

-- 4. 키워드 검색용 FTS5 가상 테이블
CREATE VIRTUAL TABLE pcf_threads_fts USING fts5(
    plan_title,
    issue_summary,
    search_keywords,
    search_text,
    content='pcf_threads',
    content_rowid='thread_id'
);
```

---

## 5. 주요 엣지 케이스 및 대응 규칙

* **다중 시트 Excel**: 시트 이름에 `PCF`, `Comment`, `Form`이 포함되거나 본문 헤더(`BUYER'S COMMENTS`)가 감지되는 시트만 파싱하고, 커버 시트나 빈 시트는 자동 무시.
* **멀티 페이지 반복 헤더**: Word/Excel에서 페이지가 넘어갈 때 `No. | BUYER'S COMMENTS | BUILDER'S REPLY` 헤더 행이 중간에 다시 등장하는 경우, 중복 헤더 행을 감지해 스킵하고 직전 스레드 연결 유지.
* **빈 턴 필터링**: `Buyer`가 비어있고 `Builder`만 독촉 메일(`Builder look forward to receive Buyer's reply`)을 쓴 턴은 DB에는 원문 보존하되, `search_text` 조합 시에는 제외.

---

## 6. 검증 기준 (Acceptance Criteria) 및 개발 순서

1. **Step 1 (Rule Parser 구현)**: 보유한 대표 샘플 파일(Excel 5종, Word 5종)을 대상으로 `Stage 1~2` 실행 → 수동 대조 시 `Hull No.`, `Plan Title`, `Comment No.`, `Turn` 분리 정확도 **90% 이상** 확인.
2. **Step 2 (AI 검토 레이어 연동)**: `Stage 3`의 Fallback 및 Enrichment 프롬프트 연결 → JSON 파싱 실패율 0%(Structured Output 적용) 및 `inferred_status` 정확도 검증.
3. **Step 3 (DB 적재 및 FTS5 테스트)**: 전체 실적 폴더 일괄 파싱 후, 특정 키워드(예: `Economizer capacity test`)로 검색했을 때 해당 `thread_id`의 `1st~5th` 전체 이력이 누락 없이 출력되는지 확인.