from argparse import ArgumentParser
from contextlib import closing
from dataclasses import asdict
import json
from pathlib import Path
import sqlite3

from .database import connect, search
from .extraction import parse_folder
from .ingest import ingest
from .review import review_folder


def run_pipeline(root: Path, output: Path, *, force: bool = False) -> tuple[dict, int]:
    parsed = parse_folder(root, output / "parsed", force=force)
    reviewed = review_folder(output / "parsed", output / "feedback")
    return {"parse": parsed, "review": reviewed["counts"]}, int(bool(parsed["failed"] or reviewed["counts"]["ERROR"]))


def main(argv: list[str] | None = None) -> int:
    parser = ArgumentParser(description="Parse and search Office comment threads")
    commands = parser.add_subparsers(dest="command", required=True)
    scan = commands.add_parser("ingest", help="Scan a folder and update the SQLite index")
    scan.add_argument("root", type=Path)
    scan.add_argument("--db", type=Path, required=True)
    scan.add_argument("--force", action="store_true", help="Reparse unchanged source documents")
    lookup = commands.add_parser("search", help="Search FTS5 and return matching threads with all turns")
    lookup.add_argument("query")
    lookup.add_argument("--db", type=Path, required=True)
    lookup.add_argument("--limit", type=int, default=20)
    for command, help_text in (
        ("parse", "Extract .xlsx and .docx files to JSON"),
        ("review", "Review extracted JSON with the configured LLM gateway"),
        ("pipeline", "Extract JSON and generate LLM feedback in one run"),
    ):
        action = commands.add_parser(command, help=help_text)
        action.add_argument("root", type=Path)
        action.add_argument("--out", type=Path, required=True)
        if command != "review":
            action.add_argument("--force", action="store_true", help="Reparse unchanged source documents")
    args = parser.parse_args(argv)

    if args.command in {"parse", "review", "pipeline"}:
        try:
            if args.command == "parse":
                report = parse_folder(args.root, args.out, force=args.force)
                failed = bool(report["failed"])
            elif args.command == "review":
                report = review_folder(args.root, args.out)
                failed = report["counts"]["ERROR"] != 0
            else:
                report, failed = run_pipeline(args.root, args.out, force=args.force)
        except (OSError, ValueError) as exc:
            parser.exit(1, f"pcf-parser: {exc}\n")
        print(json.dumps(report, indent=2, ensure_ascii=False))
        return int(failed)

    if args.command == "ingest":
        if not args.root.is_dir():
            parser.error("root must be an existing directory")
        try:
            with closing(connect(args.db)) as db:
                report = ingest(db, args.root, force=args.force)
        except (OSError, sqlite3.Error) as exc:
            parser.exit(1, f"pcf-parser: {exc}\n")
        print(json.dumps(asdict(report), indent=2, ensure_ascii=False))
        return int(bool(report.failed))

    if args.limit < 1 or not args.query.strip():
        parser.error("search requires a nonempty query and a positive limit")
    try:
        with closing(connect(args.db, create=False)) as db:
            results = search(db, args.query, args.limit)
    except sqlite3.Error as exc:
        parser.exit(1, f"pcf-parser: {exc}\n")
    print(json.dumps(results, indent=2, ensure_ascii=False))
    return 0
