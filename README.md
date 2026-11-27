# PCF Parser

Plan Comments Form의 Excel·Word 파일을 읽어 문서 → 코멘트 스레드 → 교신 회차 구조로 SQLite에 저장하는 로컬 CLI입니다. [상세 PRD](docs/PCF-PARSER-PRD.md)를 참고하세요.

## 설치와 실행

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

현재 구현은 규칙 기반 파싱과 로컬 검색 단계입니다. AI 예외 복구·요약과 실제 업무 샘플 정확도 검증은 후속 작업입니다.
