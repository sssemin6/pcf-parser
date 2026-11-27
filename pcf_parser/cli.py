from argparse import ArgumentParser
from contextlib import closing
from dataclasses import asdict
import json
from pathlib import Path
import sqlite3

from .database import connect, search
from .ingest import ingest


def main(argv: list[str] | None = None) -> int:
    parser = ArgumentParser(description="Parse and search Office comment threads")
    commands = parser.add_subparsers(dest="command", required=True)
    scan = commands.add_parser("ingest", help="Scan a folder and update the SQLite index")
    scan.add_argument("root", type=Path)
    scan.add_argument("--db", type=Path, required=True)
    lookup = commands.add_parser("search", help="Search FTS5 and return matching threads with all turns")
    lookup.add_argument("query")
    lookup.add_argument("--db", type=Path, required=True)
    lookup.add_argument("--limit", type=int, default=20)
    args = parser.parse_args(argv)

    if args.command == "ingest":
        if not args.root.is_dir():
            parser.error("root must be an existing directory")
        try:
            with closing(connect(args.db)) as db:
                report = ingest(db, args.root)
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
