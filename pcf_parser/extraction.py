from dataclasses import asdict
from datetime import datetime, timezone
import json
from pathlib import Path
from tempfile import NamedTemporaryFile

from .ingest import SkipFile, _check_encryption, _digest, _paths
from .models import METADATA_FIELDS
from .parser import ParseError, parse_document


PARSE_REPORT = "_parse_report.json"
_FORMAT_VERSION = 1


def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(mode="w", encoding="utf-8", newline="\n", dir=path.parent, delete=False) as temp:
        staged = Path(temp.name)
        try:
            temp.write(text)
            temp.close()
            staged.replace(path)
        finally:
            temp.close()
            staged.unlink(missing_ok=True)


def write_json(path: Path, data) -> None:
    write_text(path, json.dumps(data, ensure_ascii=False, indent=2) + "\n")


def _current_document(output: Path, source: Path, digest: str) -> dict | None:
    try:
        data = read_json(output)
    except (OSError, ValueError):
        return None
    if (isinstance(data, dict) and data.get("schema_version") == _FORMAT_VERSION
            and data.get("file_path") == str(source) and data.get("file_hash") == digest
            and isinstance(data.get("metadata"), dict)
            and all(isinstance(data.get(key), list) for key in ("threads", "warnings", "raw_preview"))):
        return data
    return None


def parse_folder(root: Path, output: Path, *, force: bool = False) -> dict:
    root = root.resolve()
    if not root.is_dir():
        raise NotADirectoryError(root)
    output.mkdir(parents=True, exist_ok=True)
    report = {"total": 0, "parsed": 0, "unchanged": 0, "skipped": {}, "failed": {},
              "warnings": {}, "files": []}
    for source in _paths(root):
        if source.suffix.lower() not in {".xlsx", ".docx"} or source.name.startswith("~$"):
            continue
        report["total"] += 1
        relative = source.relative_to(root)
        target = output / relative.with_name(relative.name + ".json")
        try:
            digest = _digest(source)
            parsed = None if force else _current_document(target, source, digest)
            if parsed is not None:
                report["unchanged"] += 1
            else:
                _check_encryption(source)
                document = parse_document(source, include_preview=True)
                if _digest(source) != digest:
                    raise ParseError("File changed while parsing. Retry to process the current contents.")
                parsed = asdict(document)
                parsed["metadata"] = {key: document.metadata.get(key) for key in METADATA_FIELDS}
                for thread in parsed["threads"]:
                    for order, turn in enumerate(thread["turns"], start=1):
                        turn["turn_order"] = order
                write_json(target, {
                    "schema_version": _FORMAT_VERSION,
                    "file_path": str(source),
                    "file_hash": digest,
                    "file_type": source.suffix.lower().lstrip("."),
                    "parsed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    **parsed,
                })
                report["parsed"] += 1
            report["files"].append(target.relative_to(output).as_posix())
            warnings = parsed["warnings"]
            if warnings:
                report["warnings"][relative.as_posix()] = warnings
        except SkipFile as exc:
            report["skipped"][relative.as_posix()] = str(exc)
        except Exception as exc:
            report["failed"][relative.as_posix()] = f"{type(exc).__name__}: {exc}"
    write_json(output / PARSE_REPORT, report)
    return report
