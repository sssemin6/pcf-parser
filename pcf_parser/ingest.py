from contextlib import contextmanager
from dataclasses import dataclass, field
from hashlib import sha256
import os
from pathlib import Path
import shutil
import sqlite3
import subprocess
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from . import database
from .parser import ParseError, parse_document


_LEGACY = {".xls": "xlsx", ".doc": "docx"}
_MODERN = {".xlsx", ".docx"}


class SkipFile(ValueError):
    pass


@dataclass
class Report:
    total_files: int = 0
    parsed: int = 0
    unchanged: int = 0
    skipped: int = 0
    failed: dict[str, str] = field(default_factory=dict)
    threads: int = 0
    turns: int = 0
    missing_status: int = 0
    warnings: dict[str, list[str]] = field(default_factory=dict)


def _raise(error: OSError) -> None:
    raise error


def _paths(root: Path):
    for directory, dirs, files in os.walk(root, onerror=_raise):
        dirs.sort()
        for name in sorted(files):
            path = Path(directory, name)
            if not path.is_symlink() and path.is_file():
                if name.startswith("~$") or path.suffix.lower() in _MODERN | _LEGACY.keys() | {".pdf"}:
                    yield path


def _digest(path: Path) -> str:
    checksum = sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            checksum.update(chunk)
    return checksum.hexdigest()


def _check_encryption(path: Path) -> None:
    with path.open("rb") as source:
        prefix = source.read(65536)
    if prefix.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        if "EncryptedPackage".encode("utf-16-le") in prefix:
            raise SkipFile("encrypted_office")
    elif path.suffix.lower() in _MODERN:
        with ZipFile(path) as archive:
            if any(member.flag_bits & 1 for member in archive.infolist()):
                raise SkipFile("encrypted_office")


@contextmanager
def _modern_file(path: Path):
    suffix = path.suffix.lower()
    _check_encryption(path)
    if suffix in _MODERN:
        yield path
        return
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if executable is None:
        raise SkipFile("legacy_converter_unavailable")
    with TemporaryDirectory(prefix="pcf-convert-") as directory:
        output = Path(directory)
        profile = (output / "profile").as_uri()
        target = _LEGACY[suffix]
        if target == "xlsx":
            output_format = "xlsx:Calc MS Excel 2007 XML"
        else:
            output_format = "docx:Office Open XML Text"
        result = subprocess.run(
            [executable, f"-env:UserInstallation={profile}", "--headless", "--convert-to", output_format,
             "--outdir", str(output), str(path)], capture_output=True, timeout=120, check=False,
        )
        converted = output / f"{path.stem}.{target}"
        if b"password" in (result.stderr + result.stdout).lower():
            raise SkipFile("encrypted_office")
        if result.returncode or not converted.is_file():
            raise ParseError(f"LibreOffice could not convert {suffix}")
        yield converted


def ingest(db: sqlite3.Connection, root: Path) -> Report:
    report = Report()
    for path in _paths(root.resolve()):
        report.total_files += 1
        try:
            if path.name.startswith("~$"):
                raise SkipFile("office_temporary_file")
            if path.suffix.lower() == ".pdf":
                raise SkipFile("unsupported_pdf")
            digest = _digest(path)
            if database.is_current(db, path, digest):
                with db:
                    database.clear_skip(db, path)
                report.unchanged += 1
                continue
            with _modern_file(path) as modern:
                document = parse_document(modern, original_path=path)
                file_type = modern.suffix.lower().lstrip(".")
            if _digest(path) != digest:
                raise ParseError("File changed while parsing. Retry ingestion to process the current contents.")
        except SkipFile as exc:
            database.skip(db, path, str(exc))
            report.skipped += 1
            continue
        except Exception as exc:
            report.failed[str(path)] = f"{type(exc).__name__}: {exc}"
            continue
        if not database.save(db, path, digest, file_type, document):
            report.unchanged += 1
            continue
        report.parsed += 1
        report.threads += len(document.threads)
        report.turns += sum(len(thread.turns) for thread in document.threads)
        report.missing_status += sum(not thread.raw_status for thread in document.threads)
        if document.warnings:
            report.warnings[str(path)] = document.warnings
    return report
