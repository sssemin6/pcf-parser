"""Windows console launcher for PCF Parser."""

import os
from pathlib import Path
import sqlite3
import sys


def read_source(config: Path) -> str:
    try:
        return config.read_text(encoding="utf-8-sig").strip().strip('"')
    except FileNotFoundError:
        return ""


def run_menu(pcf_main) -> int:
    config = Path(__file__).with_name("pcf-source.txt")
    source = read_source(config)
    local_data = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "PCF Parser"
    local_data.mkdir(parents=True, exist_ok=True)
    database = local_data / "pcf.sqlite3"

    while True:
        print(f"\nPCF Parser\n입력 폴더: {source or '미설정'}\nSQLite: {database}\n")
        print("1. PCF 파일 적재 또는 갱신\n2. 코멘트 검색\n3. 입력 폴더 변경\n"
              "4. JSON 추출 및 LLM 자동 검토(FB 생성)\n0. 종료")
        choice = input("선택 [Enter = 1]: ").strip()
        if choice == "0":
            return 0
        if choice == "3" or (choice in {"", "1", "4"} and not source):
            selected = input("입력 폴더 경로: ").strip().strip('"')
            if selected:
                source = selected
            continue
        if choice == "2":
            if not database.is_file():
                print("먼저 1번에서 PCF 파일을 적재하세요.")
                continue
            query = input("검색어 (예: pump capacity): ").strip()
            if not query:
                continue
            args = ["search", query, "--db", str(database)]
        elif choice in {"", "1", "4"}:
            try:
                if not Path(source).is_dir():
                    print("입력 폴더를 찾을 수 없습니다. 경로와 공유 폴더 연결을 확인하세요.")
                    continue
            except OSError as exc:
                print(f"입력 폴더에 접근할 수 없습니다. {exc}")
                continue
            if choice == "4":
                default_output = local_data / "review"
                output = input(f"작업 폴더 [Enter = {default_output}]: ").strip().strip('"')
                force = input("변경되지 않은 파일도 다시 파싱할까요? [y/N]: ").strip().lower()
                args = ["pipeline", source, "--out", output or str(default_output)]
                if force == "y":
                    args.append("--force")
            else:
                args = ["ingest", source, "--db", str(database)]
        else:
            print("0~4 중에서 선택하세요.")
            continue

        try:
            if pcf_main(args):
                print("일부 파일을 처리하지 못했습니다. 파싱 요약 또는 피드백 리포트의 오류를 확인하세요.")
        except SystemExit as exc:
            if exc.code:
                print("위 오류 메시지를 확인하세요.")


def main() -> int:
    if sys.version_info < (3, 10):
        print("Python 3.10 or later is required.", file=sys.stderr)
        return 1
    try:
        from pcf_parser.cli import main as pcf_main
    except ModuleNotFoundError as exc:
        print(f"PCF Parser dependency missing: {exc.name}", file=sys.stderr)
        print('From the PCF Parser folder, run: py -3 -m pip install .', file=sys.stderr)
        return 1

    if sys.argv[1:] == ["--self-test"]:
        from docx import Document
        from openpyxl import Workbook
        from PIL import Image

        Document()
        Workbook().close()
        Image.new("RGB", (1, 1))
        connection = sqlite3.connect(":memory:")
        try:
            connection.execute("CREATE VIRTUAL TABLE check_fts USING fts5(text)")
        finally:
            connection.close()
        print("PCF Parser runtime OK")
        return 0
    if len(sys.argv) > 1:
        return pcf_main(sys.argv[1:])
    try:
        return run_menu(pcf_main)
    except (EOFError, KeyboardInterrupt):
        print()
        return 0
    except (OSError, UnicodeError) as exc:
        print(f"PCF Parser: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
