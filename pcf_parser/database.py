import json
from pathlib import Path
import sqlite3

from .models import Document
from .parser import search_text


_SCHEMA = """
CREATE TABLE IF NOT EXISTS pcf_documents (
    doc_id INTEGER PRIMARY KEY,
    file_path TEXT NOT NULL UNIQUE,
    file_hash TEXT NOT NULL,
    file_type TEXT NOT NULL,
    hull_no TEXT,
    ship_type TEXT,
    plan_title TEXT,
    plan_no TEXT,
    comment_party TEXT,
    approval_status TEXT,
    parse_method TEXT NOT NULL DEFAULT 'rule',
    parsed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS pcf_threads (
    thread_id INTEGER PRIMARY KEY,
    doc_id INTEGER NOT NULL REFERENCES pcf_documents(doc_id) ON DELETE CASCADE,
    comment_no TEXT NOT NULL,
    source TEXT NOT NULL,
    initial TEXT,
    raw_status TEXT,
    inferred_status TEXT NOT NULL DEFAULT 'Unknown',
    issue_summary TEXT,
    final_conclusion TEXT,
    search_keywords TEXT,
    search_text TEXT NOT NULL,
    confidence_score REAL
);
CREATE INDEX IF NOT EXISTS pcf_threads_doc_id ON pcf_threads(doc_id);
CREATE TABLE IF NOT EXISTS pcf_turns (
    turn_id INTEGER PRIMARY KEY,
    thread_id INTEGER NOT NULL REFERENCES pcf_threads(thread_id) ON DELETE CASCADE,
    turn_order INTEGER NOT NULL,
    round_label TEXT,
    buyer_comment TEXT,
    builder_reply TEXT,
    has_image INTEGER NOT NULL,
    attachments TEXT NOT NULL,
    UNIQUE(thread_id, turn_order)
);
CREATE VIRTUAL TABLE IF NOT EXISTS pcf_threads_fts USING fts5(
    plan_title, issue_summary, search_keywords, search_text
);
CREATE TRIGGER IF NOT EXISTS pcf_threads_insert_fts AFTER INSERT ON pcf_threads BEGIN
    INSERT INTO pcf_threads_fts(rowid, plan_title, issue_summary, search_keywords, search_text)
    SELECT new.thread_id, plan_title, new.issue_summary, new.search_keywords, new.search_text
    FROM pcf_documents WHERE doc_id = new.doc_id;
END;
CREATE TRIGGER IF NOT EXISTS pcf_threads_delete_fts AFTER DELETE ON pcf_threads BEGIN
    DELETE FROM pcf_threads_fts WHERE rowid = old.thread_id;
END;
CREATE TRIGGER IF NOT EXISTS pcf_threads_update_fts
AFTER UPDATE OF issue_summary, search_keywords, search_text ON pcf_threads BEGIN
    UPDATE pcf_threads_fts SET issue_summary = new.issue_summary,
        search_keywords = new.search_keywords, search_text = new.search_text
    WHERE rowid = new.thread_id;
END;
CREATE TRIGGER IF NOT EXISTS pcf_documents_update_fts AFTER UPDATE OF plan_title ON pcf_documents BEGIN
    UPDATE pcf_threads_fts SET plan_title = new.plan_title
    WHERE rowid IN (SELECT thread_id FROM pcf_threads WHERE doc_id = new.doc_id);
END;
CREATE TABLE IF NOT EXISTS skipped_files (
    file_path TEXT PRIMARY KEY,
    reason TEXT NOT NULL,
    detected_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
"""

_FIELDS = ("hull_no", "ship_type", "plan_title", "plan_no", "comment_party", "approval_status")


def connect(path: Path, *, create: bool = True) -> sqlite3.Connection:
    location = path.resolve().as_uri() + ("?mode=rwc" if create else "?mode=ro")
    db = sqlite3.connect(location, uri=True)
    try:
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys = ON")
        if create:
            db.executescript(_SCHEMA)
    except Exception:
        db.close()
        raise
    return db


def is_current(db: sqlite3.Connection, path: Path, file_hash: str) -> bool:
    return db.execute(
        "SELECT 1 FROM pcf_documents WHERE file_path = ? AND file_hash = ?", (str(path), file_hash)
    ).fetchone() is not None


def save(db: sqlite3.Connection, path: Path, file_hash: str, file_type: str, document: Document) -> bool:
    with db:
        db.execute("BEGIN IMMEDIATE")
        if is_current(db, path, file_hash):
            return False
        db.execute("""
            INSERT INTO pcf_documents (file_path, file_hash, file_type, hull_no, ship_type,
                                       plan_title, plan_no, comment_party, approval_status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(file_path) DO UPDATE SET file_hash = excluded.file_hash, file_type = excluded.file_type,
                hull_no = excluded.hull_no, ship_type = excluded.ship_type, plan_title = excluded.plan_title,
                plan_no = excluded.plan_no, comment_party = excluded.comment_party,
                approval_status = excluded.approval_status, parse_method = 'rule', parsed_at = CURRENT_TIMESTAMP
        """, (str(path), file_hash, file_type, *(document.metadata.get(key) for key in _FIELDS)))
        doc_id = db.execute("SELECT doc_id FROM pcf_documents WHERE file_path = ?", (str(path),)).fetchone()[0]
        db.execute("DELETE FROM pcf_threads WHERE doc_id = ?", (doc_id,))
        for thread in document.threads:
            initial = next((turn.buyer_comment for turn in thread.turns if turn.buyer_comment), None)
            cursor = db.execute("""
                INSERT INTO pcf_threads (doc_id, comment_no, source, initial, raw_status, search_text)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (doc_id, thread.comment_no, thread.source, initial, thread.raw_status or None, search_text(thread)))
            db.executemany("""
                INSERT INTO pcf_turns (thread_id, turn_order, round_label, buyer_comment, builder_reply, has_image, attachments)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, ((cursor.lastrowid, order, turn.round_label, turn.buyer_comment, turn.builder_reply,
                   int(turn.has_image), json.dumps(turn.attachments, ensure_ascii=False))
                  for order, turn in enumerate(thread.turns, start=1)))
        clear_skip(db, path)
    return True


def skip(db: sqlite3.Connection, path: Path, reason: str) -> None:
    with db:
        db.execute("""
            INSERT INTO skipped_files(file_path, reason) VALUES (?, ?)
            ON CONFLICT(file_path) DO UPDATE SET reason = excluded.reason, detected_at = CURRENT_TIMESTAMP
        """, (str(path), reason))


def clear_skip(db: sqlite3.Connection, path: Path) -> None:
    db.execute("DELETE FROM skipped_files WHERE file_path = ?", (str(path),))


def search(db: sqlite3.Connection, query: str, limit: int) -> list[dict]:
    matches = db.execute("""
        SELECT t.thread_id, t.comment_no, t.source, t.raw_status, d.file_path, d.plan_title, d.hull_no
        FROM pcf_threads_fts
        JOIN pcf_threads AS t ON t.thread_id = pcf_threads_fts.rowid
        JOIN pcf_documents AS d ON d.doc_id = t.doc_id
        WHERE pcf_threads_fts MATCH ? ORDER BY rank, t.thread_id LIMIT ?
    """, (query, limit)).fetchall()
    results = []
    for match in matches:
        turns = []
        for row in db.execute("""
            SELECT turn_order, round_label, buyer_comment, builder_reply, has_image, attachments
            FROM pcf_turns WHERE thread_id = ? ORDER BY turn_order
        """, (match["thread_id"],)):
            turns.append({**dict(row), "attachments": json.loads(row["attachments"]), "has_image": bool(row["has_image"])})
        results.append({**dict(match), "turns": turns})
    return results
