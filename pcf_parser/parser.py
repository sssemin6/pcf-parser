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
    "initial": r"(?:reviewer(?:'s)?\s+)?initials?",
}
_ROUND = re.compile(r"^\s*\(?(?P<label>[a-z]\)|\d{1,3}(?:st|nd|rd|th)\b\)?)[.\s:–-]*", re.I | re.M)
_NUMBER = re.compile(r"[A-Za-z0-9]+(?:[./-][A-Za-z0-9]+)*\.?\)?")
_FOLLOW_UP = re.compile(r"^(?:the\s+)?(?:builder|yard)\s+(?:is\s+)?look(?:s|ing)?\s+forward\s+to\b", re.I)
_HULL_INLINE = re.compile(
    r"hull\s*(?:no\.?|number)(?:[ \t]*[:：][ \t]*|[ \t]+)(?=[\w /-]*\d)"
    r"([A-Z0-9]+(?:[ \t]*[/-][ \t]*[A-Z0-9]+)*)", re.I,
)
_SHIP_DESCRIPTION = re.compile(
    r"(?:\d[\d,.]*\s*(?:CBM|M[³3]|DWT|TEU)\s+.{0,100}|(?:LNG|LPG|VLCC|BULK|CONTAINER|OIL|CHEMICAL)\s+.{0,50})"
    r"\b(?:CARRIER|TANKER|SHIP|VESSEL)\b.*", re.I,
)
_APPROVAL = {label.casefold(): label for label in ("Approved", "Approved with comments", "Not approved", "Noted")}
_CHECK = {"x", "✓", "✔", "v", "o"}
_ADMIN = re.compile(
    r"(?:waiting\s+(?:(?:for|your|buyer['’]?s|builder['’]?s|owner['’]?s|class|the|final)\s+)*confirmation"
    r"|can\s+(?:the\s+)?builder\s+understand\s+that\s+(?:the|this)\s+comment\s+is\s+closed"
    r"|noted[,.\s]+(?:many\s+)?thanks?"
    r"|thank\s+you[,.!\s]+(?:confirmed|agreed)(?:\s*/\s*(?:confirmed|agreed))?)\s*[.!?]*",
    re.I,
)
_EMPTY = Cell()


def _normalize(text: str) -> str:
    return " ".join(text.split()).removesuffix(":").removesuffix("：").strip().casefold()


def _metadata(table: Table, limit: int | None) -> dict[str, str]:
    found: dict[str, str] = {}
    ship_fallback = ""
    checked_approval = ""
    for index, row in sorted(table.rows.items()):
        if limit is not None and index >= limit:
            break
        cells = list(sorted(row.items()))
        for position, (column, cell) in enumerate(cells):
            if not cell.text:
                continue
            lines = cell.text.splitlines()
            hull = _HULL_INLINE.fullmatch(lines[0].strip())
            if hull:
                found.setdefault("hull_no", re.sub(r"\s*([/-])\s*", r"\1", hull[1]))
            if not ship_fallback and (hull or re.fullmatch(_METADATA["hull_no"], _normalize(lines[0]))):
                below = table.rows.get(index + 1, {}) if limit is None or index + 1 < limit else {}
                first_line = [below[c].text.splitlines()[0].strip() for c in sorted(below)
                              if c >= column and below[c].origin != cell.origin and below[c].text]
                ship_fallback = next(
                    (value for value in [*(line.strip() for line in lines[1:]), *first_line]
                     if _SHIP_DESCRIPTION.fullmatch(value)), "",
                )
            approval = _APPROVAL.get(_normalize(cell.text))
            if approval and not checked_approval:
                # A merged label may occupy multiple columns before the checkbox.
                span = [c for c, value in cells if value.origin == cell.origin]
                if any(_normalize(row.get(c, _EMPTY).text) in _CHECK for c in (min(span) - 1, max(span) + 1)):
                    checked_approval = approval
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
                            if field_name == "hull_no":
                                value, *description = neighbor.text.splitlines()
                                found[field_name] = value.strip()
                                if not ship_fallback:
                                    ship_fallback = next((line.strip() for line in description
                                                          if _SHIP_DESCRIPTION.fullmatch(line.strip())), "")
                            else:
                                found[field_name] = neighbor.text
                        break
    if ship_fallback:
        found.setdefault("ship_type", ship_fallback)
    if checked_approval:
        found["approval_status"] = checked_approval
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


def _headers(row: dict[int, Cell]) -> dict[str, tuple[int, ...]]:
    columns: dict[str, tuple[int, ...]] = {}
    for column, cell in sorted(row.items()):
        for role, pattern in _COLUMN.items():
            if role not in columns and re.fullmatch(pattern, _normalize(cell.text)):
                columns[role] = tuple(c for c in sorted(row) if row[c].origin == cell.origin)
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


def _side_parts(cells: list[Cell], seen: set[object]) -> list[_Part]:
    parts: list[_Part] = []
    label = ""
    for cell in cells:
        marker = _ROUND.fullmatch(cell.text)
        if marker:
            label = marker["label"].casefold()
        if cell.origin in seen:
            continue
        seen.add(cell.origin)
        if marker:
            cell = Cell(origin=cell.origin, has_image=cell.has_image, attachments=cell.attachments)
        for part in _parts(cell):
            label = part.label or label
            part.label = label
            if parts and _round_key(parts[-1].label) == _round_key(label):
                previous = parts[-1]
                previous.text = "\n".join(filter(None, (previous.text, part.text)))
                previous.has_image |= part.has_image
                previous.attachments = tuple(dict.fromkeys(previous.attachments + part.attachments))
            else:
                parts.append(part)
    return parts


def _row_cells(row: dict[int, Cell], columns: tuple[int, ...]) -> list[Cell]:
    return list({row[column].origin: row[column] for column in columns if column in row}.values())


def _field(row: dict[int, Cell], columns: dict[str, tuple[int, ...]], role: str) -> str:
    return " ".join(cell.text for cell in _row_cells(row, columns.get(role, ())) if cell.text)


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

    def add(self, buyer: list[Cell], builder: list[Cell], round_label: str) -> bool:
        left_parts = _side_parts(buyer, self.seen)
        right_parts = _side_parts(builder, self.seen)
        pairs = _join_parts(left_parts, right_parts)
        for left, right in pairs:
            label = (left and left.label) or (right and right.label) or round_label
            image = bool((left and left.has_image) or (right and right.has_image))
            files = tuple(dict.fromkeys((left.attachments if left else ()) + (right.attachments if right else ())))
            self._add_turn(Turn(label, _with_tags(left), _with_tags(right), image, files))
        unmatched = left_parts and right_parts and any(left is None or right is None for left, right in pairs)
        ambiguous = any(len(cell.text) >= 500 and len(parts) <= 1 and re.search(
            r"\b\d+(?:st|nd|rd|th)\b.*\b\d+(?:st|nd|rd|th)\b", cell.text, re.I | re.S
        ) for cells, parts in ((buyer, left_parts), (builder, right_parts)) if parts for cell in cells)
        return bool(unmatched or ambiguous)


def _parse_table(table: Table) -> tuple[list[Thread], list[str], str]:
    threads: list[Thread] = []
    warnings = []
    columns: dict[str, tuple[int, ...]] = {}
    current: _ThreadState | None = None
    party = ""
    for index, row in sorted(table.rows.items()):
        headers = _headers(row)
        if headers:
            columns = headers
            party = "Class" if _normalize(_field(row, columns, "buyer")).startswith("class") else "Buyer"
            continue
        if not columns:
            continue

        number = _field(row, columns, "number")
        buyer = _row_cells(row, columns["buyer"])
        builder = _row_cells(row, columns["builder"])
        status = _field(row, columns, "status")
        round_label = _field(row, columns, "round")
        if number:
            if not _NUMBER.fullmatch(number):
                current = None
                continue
            if current is None or current.thread.comment_no != number:
                thread = Thread(number, table.name)
                current = _ThreadState(thread)
                threads.append(thread)
        if current is None:
            if any(cell.text or cell.attachments or cell.has_image for cell in buyer + builder):
                warnings.append(f"{table.name}, row {index + 1}: comment number missing")
            continue
        if status:
            current.thread.raw_status = status
        if not current.thread.reviewer_initial:
            current.thread.reviewer_initial = _field(row, columns, "initial")
        if current.add(buyer, builder, round_label):
            warnings.append(f"{table.name}, row {index + 1}: ambiguous rounds for comment {current.thread.comment_no}")
    for thread in threads:
        families = {_round_key(turn.round_label)[0] for turn in thread.turns}
        if families in ({"number"}, {"letter"}):
            thread.turns.sort(key=lambda turn: _round_key(turn.round_label)[1])
    return [thread for thread in threads if thread.turns], warnings, party


def _raw_preview(tables: list[Table], accepted: list[Table]) -> list[dict]:
    preview = []
    preferred = {table.name: table for table in tables[:1] + accepted + tables}
    for table in list(preferred.values())[:5]:
        remaining = 35
        rows = []
        ordered = sorted(table.rows.items())
        header_start = next((i for i, (_, row) in enumerate(ordered) if _headers(row)), None)
        head = ordered[:6]
        body = ordered[header_start:header_start + 8] if header_start is not None else ordered[6:12]
        for index, row in dict(head + body).items():
            cells = []
            columns = tuple(sorted(row)[:20])
            for cell in _row_cells(row, columns):
                if not (cell.text or cell.has_image or cell.attachments):
                    continue
                span = [column + 1 for column in columns if row[column].origin == cell.origin]
                cells.append({"columns": span, "text": cell.text[:160], "has_image": cell.has_image,
                              "attachments": [name[:80] for name in cell.attachments[:3]]})
                remaining -= 1
                if not remaining:
                    break
            if cells:
                rows.append({"row": index + 1, "cells": cells})
            if not remaining:
                break
        preview.append({"source": table.name[:150], "rows": rows})
    return preview


def parse_document(path: Path, original_path: Path | None = None, *, include_preview: bool = False) -> Document:
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
        message = "No comment threads found in a table with recognized headers"
        if not include_preview:
            raise ParseError(message)
        warnings.append(message)

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
    return Document(metadata, threads, warnings, _raw_preview(tables, accepted_tables) if include_preview else [])


def search_text(thread: Thread) -> str:
    passages = []
    for turn in thread.turns:
        if not turn.buyer_comment and _FOLLOW_UP.match(turn.builder_reply):
            continue
        for text in (turn.buyer_comment, turn.builder_reply):
            passages.extend(line for line in text.splitlines() if line.strip() and not _ADMIN.fullmatch(line.strip()))
    return "\n".join(passages)
