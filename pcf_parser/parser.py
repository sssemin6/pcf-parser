from dataclasses import dataclass, field
from pathlib import Path
import re

from .models import Document, Thread, Turn
from .office import Cell, Table, read_excel, read_word


class ParseError(ValueError):
    pass


_METADATA = {
    "hull_no": r"hull\s*(?:no\.?|number)",
    "ship_type": r"(?:ship|vessel)\s*type",
    "plan_title": r"(?:plan|drawing)\s*title",
    "plan_no": r"(?:plan|drawing)\s*(?:no\.?|number)",
    "comment_party": r"comment\s*party",
    "approval_status": r"approval\s*status",
}
_COLUMN = {
    "number": r"(?:(?:comment|item)\s+)?(?:no\.?|number)|item|#",
    "buyer": r"(?:buyer|owner|class)(?:['’]s|s)?\s+comments?",
    "builder": r"(?:builder|yard)(?:['’]s|s)?\s+(?:repl(?:y|ies)|responses?)",
    "status": r"status|remarks?",
    "round": r"round(?:\s+no\.?)?|revision|rev\.?",
}
_ROUND = re.compile(r"^\s*\(?(?P<label>[a-z]\)|\d{1,3}(?:st|nd|rd|th)\b\)?)[.\s:–-]*", re.I | re.M)
_NUMBER = re.compile(r"[A-Za-z0-9]+(?:[./-][A-Za-z0-9]+)*\.?\)?")
_FOLLOW_UP = re.compile(r"^(?:the\s+)?(?:builder|yard)\s+(?:is\s+)?look(?:s|ing)?\s+forward\s+to\b", re.I)
_EMPTY = Cell()


def _normalize(text: str) -> str:
    return " ".join(text.split()).removesuffix(":").removesuffix("：").strip().casefold()


def _metadata(table: Table, limit: int | None) -> dict[str, str]:
    found: dict[str, str] = {}
    for index, row in sorted(table.rows.items()):
        if limit is not None and index >= limit:
            break
        cells = list(sorted(row.items()))
        for position, (column, cell) in enumerate(cells):
            if not cell.text:
                continue
            for field_name, pattern in _METADATA.items():
                if field_name in found:
                    continue
                match = re.fullmatch(rf"(?:{pattern})\s*[:：]\s*(.+)", cell.text, re.I)
                if match:
                    found[field_name] = match[1].strip()
                elif re.fullmatch(pattern, _normalize(cell.text)):
                    edge = column
                    for next_column, neighbor in cells[position + 1:]:
                        if neighbor.origin == cell.origin:
                            edge = next_column
                            continue
                        if next_column == edge + 1 and neighbor.text and not any(
                            re.match(rf"(?:{label})(?:$|\s*[:：])", _normalize(neighbor.text))
                            for label in _METADATA.values()
                        ):
                            found[field_name] = neighbor.text
                        break
    return found


def _path_metadata(path: Path) -> dict[str, str]:
    result = {}
    for part in (path.stem, *reversed(path.parent.parts)):
        for field_name in ("hull_no", "plan_no", "ship_type"):
            if field_name in result:
                continue
            if field_name == "ship_type":
                pattern = r"(?:ship|vessel)[ _-]*type[ _:=-]+([A-Za-z]+(?:[ -][A-Za-z]+)*)"
            else:
                label = "hull" if field_name == "hull_no" else "(?:plan|drawing)"
                pattern = rf"{label}(?:[ _-]*(?:no\.?|number))?[ _:=-]+([A-Z0-9]+(?:[.-][A-Z0-9]+)*)"
            match = re.search(rf"(?:^|[^\w]|_){pattern}(?=$|[^\w]|_)", part, re.I)
            if match and (field_name == "ship_type" or any(char.isdigit() for char in match[1])):
                result[field_name] = match[1]
    return result


def _headers(row: dict[int, Cell]) -> dict[str, int]:
    columns: dict[str, int] = {}
    for column, cell in sorted(row.items()):
        for role, pattern in _COLUMN.items():
            if role not in columns and re.fullmatch(pattern, _normalize(cell.text)):
                columns[role] = column
    return columns if {"number", "buyer", "builder"} <= columns.keys() else {}


@dataclass
class _Part:
    label: str
    text: str
    has_image: bool = False
    attachments: tuple[str, ...] = ()


def _parts(cell: Cell) -> list[_Part]:
    if not (cell.text or cell.has_image or cell.attachments):
        return []
    markers = list(_ROUND.finditer(cell.text))
    parts = []
    start = 0
    label = ""
    for marker in markers:
        text = cell.text[start:marker.start()].strip()
        if text or label:
            parts.append(_Part(label, text))
        label = marker["label"].casefold()
        start = marker.end()
    parts.append(_Part(label, cell.text[start:].strip()))

    for filename in cell.attachments:
        part = next((part for part in parts if filename in part.text), parts[-1])
        part.attachments += (filename,)
    parts[-1].has_image = cell.has_image
    return parts


def _round_key(label: str) -> tuple[str, int]:
    label = label.strip().lower()
    match = _ROUND.fullmatch(label)
    if match:
        token = match["label"]
        if token[0].isdigit():
            return "number", int(re.match(r"\d+", token)[0])
        return "letter", ord(token[0])
    return label, 0


def _join_parts(buyer: list[_Part], builder: list[_Part]) -> list[tuple[_Part | None, _Part | None]]:
    pairs: list[tuple[_Part | None, _Part | None]] = []
    remaining = list(builder)
    for left in buyer:
        position = next((i for i, right in enumerate(remaining)
                         if _round_key(left.label) == _round_key(right.label)), None)
        pairs.append((left, remaining.pop(position) if position is not None else None))
    one_side_is_unlabeled = not (any(part.label for part in buyer) and any(part.label for part in builder))
    if len(buyer) == len(builder) and one_side_is_unlabeled:
        for index, (left, matched) in enumerate(pairs):
            if matched is None and remaining:
                pairs[index] = (left, remaining.pop(0))
    pairs.extend((None, right) for right in remaining)
    return pairs


def _with_tags(part: _Part | None) -> str:
    if part is None:
        return ""
    tags = [f"[FILE_ATTACHED: {filename}]" for filename in part.attachments]
    if part.has_image:
        tags.insert(0, "[IMAGE_ATTACHED]")
    return "\n".join(filter(None, [part.text, *tags]))


@dataclass
class _ThreadState:
    thread: Thread
    seen: set[object] = field(default_factory=set)
    turns_by_round: dict[tuple[str, int], list[Turn]] = field(default_factory=dict)

    def _add_turn(self, turn: Turn) -> None:
        key = _round_key(turn.round_label)
        if turn.round_label:
            same_round = self.turns_by_round.setdefault(key, [])
            for previous in same_round:
                if (previous.buyer_comment and turn.buyer_comment) or (previous.builder_reply and turn.builder_reply):
                    continue
                previous.buyer_comment = previous.buyer_comment or turn.buyer_comment
                previous.builder_reply = previous.builder_reply or turn.builder_reply
                previous.has_image = previous.has_image or turn.has_image
                previous.attachments = tuple(dict.fromkeys(previous.attachments + turn.attachments))
                return
            same_round.append(turn)
        self.thread.turns.append(turn)

    def add(self, buyer: Cell, builder: Cell, round_label: str) -> bool:
        parts = []
        for cell in (buyer, builder):
            parts.append(_parts(cell) if cell.origin not in self.seen else [])
            self.seen.add(cell.origin)
        left_parts, right_parts = parts
        pairs = _join_parts(left_parts, right_parts)
        for left, right in pairs:
            label = (left and left.label) or (right and right.label) or round_label
            image = bool((left and left.has_image) or (right and right.has_image))
            files = tuple(dict.fromkeys((left.attachments if left else ()) + (right.attachments if right else ())))
            self._add_turn(Turn(label, _with_tags(left), _with_tags(right), image, files))
        unmatched = left_parts and right_parts and any(left is None or right is None for left, right in pairs)
        ambiguous = any(len(cell.text) >= 500 and len(parts) <= 1 and re.search(
            r"\b\d+(?:st|nd|rd|th)\b.*\b\d+(?:st|nd|rd|th)\b", cell.text, re.I | re.S
        ) for cell, parts in ((buyer, left_parts), (builder, right_parts)) if parts)
        return bool(unmatched or ambiguous)

def _parse_table(table: Table) -> tuple[list[Thread], list[str], str]:
    threads: list[Thread] = []
    warnings = []
    columns: dict[str, int] = {}
    current: _ThreadState | None = None
    party = ""
    for index, row in sorted(table.rows.items()):
        headers = _headers(row)
        if headers:
            columns = headers
            party = "Class" if _normalize(row[columns["buyer"]].text).startswith("class") else "Buyer"
            continue
        if not columns:
            continue

        number = row.get(columns["number"], _EMPTY).text.strip()
        buyer = row.get(columns["buyer"], _EMPTY)
        builder = row.get(columns["builder"], _EMPTY)
        status = row.get(columns.get("status", -1), _EMPTY).text
        round_label = row.get(columns.get("round", -1), _EMPTY).text
        if number:
            if not _NUMBER.fullmatch(number):
                current = None
                continue
            if current is None or current.thread.comment_no != number:
                thread = Thread(number, table.name)
                current = _ThreadState(thread)
                threads.append(thread)
        if current is None:
            if any(cell.text or cell.attachments or cell.has_image for cell in (buyer, builder)):
                warnings.append(f"{table.name}, row {index + 1}: comment number missing")
            continue
        if status:
            current.thread.raw_status = status
        if current.add(buyer, builder, round_label):
            warnings.append(f"{table.name}, row {index + 1}: ambiguous rounds for comment {current.thread.comment_no}")
    for thread in threads:
        families = {_round_key(turn.round_label)[0] for turn in thread.turns}
        if families in ({"number"}, {"letter"}):
            thread.turns.sort(key=lambda turn: _round_key(turn.round_label)[1])
    return [thread for thread in threads if thread.turns], warnings, party


def parse_document(path: Path, original_path: Path | None = None) -> Document:
    file_type = path.suffix.lower()
    if file_type == ".xlsx":
        tables = read_excel(path)
    elif file_type == ".docx":
        tables = read_word(path)
    else:
        raise ParseError(f"Unsupported document format: {file_type}")

    threads: list[Thread] = []
    warnings = []
    parties = set()
    accepted_tables = []
    for table in tables:
        parsed, table_warnings, party = _parse_table(table)
        if parsed:
            threads.extend(parsed)
            accepted_tables.append(table)
            parties.add(party)
        if party:
            warnings.extend(table_warnings)
    if not threads:
        raise ParseError("No comment threads found in a table with recognized headers")

    metadata: dict[str, str] = {}
    if file_type == ".xlsx":
        names = {table.name for table in accepted_tables}
        metadata_sources = accepted_tables + [table for table in tables if table.name not in names]
    else:
        metadata_sources = tables[:1]
    for table in metadata_sources:
        for key, value in _metadata(table, 15 if file_type == ".xlsx" else None).items():
            metadata.setdefault(key, value)
    for key, value in _path_metadata(original_path or path).items():
        metadata.setdefault(key, value)
    if "comment_party" in metadata:
        party = _normalize(metadata["comment_party"])
        if party in {"owner", "buyer", "class"}:
            metadata["comment_party"] = "Class" if party == "class" else "Buyer"
    elif len(parties) == 1:
        metadata["comment_party"] = parties.pop()
    for key in ("hull_no", "plan_title"):
        if not metadata.get(key):
            warnings.append(f"Missing metadata: {key}")
    return Document(metadata, threads, warnings)


def search_text(thread: Thread) -> str:
    passages = []
    for turn in thread.turns:
        if not turn.buyer_comment and _FOLLOW_UP.match(turn.builder_reply):
            continue
        passages.extend(filter(None, (turn.buyer_comment, turn.builder_reply)))
    return "\n".join(passages)
