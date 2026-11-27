# PCF Parser

Plan Comments Form의 Excel·Word 파일을 읽어 문서 → 코멘트 스레드 → 교신 회차 구조로 SQLite에 저장하는 로컬 CLI입니다. [상세 PRD](docs/PCF-PARSER-PRD.md)를 참고하세요.

## 설치와 실행

### Windows에서 더블클릭 실행

Windows용 ZIP은 압축을 푼 폴더에서 `run-pcf.cmd`를 더블클릭하면 실행됩니다. GitHub 소스에서 실행하는 경우에는 Python 3.10 이상을 설치하고 프로젝트 폴더에서 `py -3 -m pip install .`을 한 번 실행하세요.

`run-pcf.cmd` 옆에 `pcf-source.txt`를 만들고 PCF 입력 폴더의 경로를 UTF-8로 한 줄에 적으면 기본 경로로 사용합니다. 로컬 경로와 `\\server\share\folder` 형식의 UNC 경로를 지원합니다. 실행 후 1번은 적재, 2번은 검색, 3번은 입력 폴더 변경입니다. 검색 DB는 `%LOCALAPPDATA%\PCF Parser\pcf.sqlite3`에 저장됩니다. PowerShell에서 `.\run-pcf.cmd --self-test`로 런타임을 점검할 수 있습니다.

### 명령줄에서 실행

Python 3.10 이상과 FTS5가 활성화된 SQLite가 필요합니다. 구형 `.xls`·`.doc` 변환에는 LibreOffice를 설치하고 `libreoffice` 또는 `soffice`를 PATH에서 실행할 수 있어야 합니다.
구형 파일은 변환 과정에서 표 구조가 손실될 수 있습니다. 이 경우 파싱 리포트에 실패 파일로 표시됩니다.

```sh
python -m pip install .
pcf-parser ingest ./sample-pcf --db ./pcf.sqlite3
pcf-parser search '"pump capacity"' --db ./pcf.sqlite3
python -m unittest discover -s tests -v
```

`ingest`는 재귀 탐색 후 SHA-256이 새롭거나 바뀐 파일을 적재합니다. 건너뛴 파일의 사유는 `skipped_files`에 남으며, 파일별 실패와 확인이 필요한 항목은 JSON 리포트에 나옵니다. 같은 경로의 파일을 다시 적재할 때 문서와 검색 인덱스를 한 트랜잭션에서 갱신합니다. 수정본 파싱에 실패하면 마지막으로 적재에 성공한 내용이 유지됩니다.

`search`는 SQLite FTS5 쿼리를 지원합니다. 키워드가 하나의 교신 회차에만 들어 있어도 검색 결과에는 해당 스레드 전체가 표시됩니다. 병합 셀에 붙은 이미지·첨부파일은 태그로 표시합니다. 첨부 개체에 원래 파일명이 없으면 Office 패키지의 내부 이름이 표시될 수 있습니다.

기존 SQLite의 `initial`은 다음 적재 때 `initial_comment`로 자동 변환됩니다. 예전에 적재한 문서에도 새 파싱 규칙과 `reviewer_initial`을 적용하려면 `ingest --force`로 다시 적재하세요.

## JSON 추출과 LLM 구조 검토

```powershell
.\run-pcf.cmd parse "C:\PCF" --out "$env:LOCALAPPDATA\PCF Parser\review\parsed"
.\run-pcf.cmd review "$env:LOCALAPPDATA\PCF Parser\review\parsed" --out "$env:LOCALAPPDATA\PCF Parser\review\feedback"
.\run-pcf.cmd pipeline "C:\PCF" --out "$env:LOCALAPPDATA\PCF Parser\review"
```

`pcf-parser` 명령으로도 같은 하위 명령을 실행할 수 있습니다. 대화형 메뉴에서는 **4번**을 사용합니다.

- `parse`는 `.xlsx`와 `.docx`를 하위 폴더까지 읽어 `이름.xlsx.json`처럼 파일별로 저장합니다. SHA-256이 같은 파일은 재사용하고, `--force`는 강제로 다시 파싱합니다. `_parse_report.json`에 이번 실행에 포함된 파일과 오류를 기록합니다.
- `review`는 추출한 메타데이터, 스레드, 경고, 길이가 제한된 원본 표 미리보기를 설정된 LLM 게이트웨이에 전송합니다. `feedback_report.json`에는 파일별 판정과 오류, `feedback_summary.md`에는 FAIL/WARN 파일, 공통 패턴, 수정 제안을 저장합니다.
- `pipeline`은 두 단계를 연결해 작업 폴더의 `parsed`와 `feedback` 아래에 저장합니다. 같은 자료에 파서 수정 효과를 확인할 때 `pipeline --force`로 다시 실행하세요.

| PCF 환경변수 | onlyou 폴백 |
| --- | --- |
| `PCF_LLM_URL` | `MAILSEARCH_GATEWAY_URL` |
| `PCF_LLM_API_KEY` | `MAILSEARCH_GATEWAY_API_KEY` |
| `PCF_LLM_MODEL` | `MAILSEARCH_GATEWAY_REWRITE_MODEL`, 없으면 `MAILSEARCH_GATEWAY_ANSWER_MODELS`의 첫 모델 |
| `PCF_LLM_TLS_VERIFY` | `MAILSEARCH_GATEWAY_TLS_VERIFY`, 기본값 `1` |
| `PCF_LLM_CA_FILE` | `MAILSEARCH_GATEWAY_CA_FILE` |

URL은 HTTPS를 사용합니다. Windows에서 사내 CA가 필요하면 `PCF_LLM_CA_FILE`에 PEM 파일 경로를 지정하세요. `PASS`·`WARN`·`FAIL`은 검토 판정이며, 명령은 파싱 실패나 검토 실행 오류가 있을 때 종료 코드 1을 반환합니다.
