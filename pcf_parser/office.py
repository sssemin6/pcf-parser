from dataclasses import dataclass, replace
from itertools import repeat
from pathlib import Path, PurePosixPath
import posixpath
import re
from xml.etree import ElementTree
from zipfile import ZipFile

from docx import Document as open_word
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class Cell:
    text: str = ""
    origin: object = None
    has_image: bool = False
    attachments: tuple[str, ...] = ()


@dataclass
class Table:
    name: str
    rows: dict[int, dict[int, Cell]]


_FILE_LABEL = re.compile(
    r"\bAttachment(?:[ \t]*:[ \t]*|[ \t]+)(?:#[\w-]+[ \t]*[:–-]?[ \t]*)?([^/\\<>:\r\n]+?\."
    r"(?:pdf|xlsx?|docx?|pptx?|png|jpe?g|zip|dwg|txt))\b",
    re.IGNORECASE,
)
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _filenames(text: str) -> tuple[str, ...]:
    return tuple(dict.fromkeys(match.strip() for match in _FILE_LABEL.findall(text)))


def _embeddings(path: Path) -> dict[str, dict[tuple[int, int], tuple[str, ...]]]:
    result = {}
    with ZipFile(path) as archive:
        names = set(archive.namelist())
        for name in sorted(names):
            if not re.fullmatch(r"xl/worksheets/[^/]+\.xml", name):
                continue
            sheet = ElementTree.fromstring(archive.read(name))
            objects = sheet.findall(".//{*}oleObject")
            if not objects:
                continue
            parent, filename = posixpath.split(name)
            rel_path = f"{parent}/_rels/{filename}.rels"
            if rel_path not in names:
                continue
            relationships = {
                entry.get("Id"): entry.get("Target", "")
                for entry in ElementTree.fromstring(archive.read(rel_path))
                if entry.get("TargetMode") != "External"
            }
            anchors: dict[tuple[int, int], list[str]] = {}
            for embedded in objects:
                target = relationships.get(embedded.get(_REL_NS + "id"))
                anchor = embedded.find(".//{*}from")
                if not target or anchor is None:
                    continue
                row, column = anchor.findtext("{*}row"), anchor.findtext("{*}col")
                if row is not None and column is not None and row.isdecimal() and column.isdecimal():
                    anchors.setdefault((int(row), int(column)), []).append(PurePosixPath(target).name)
            result[name] = {
                coordinate: tuple(dict.fromkeys(files)) for coordinate, files in anchors.items()
            }
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        relationships = {
            entry.get("Id"): entry.get("Target", "")
            for entry in ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        }
        result_by_sheet = {}
        for sheet in workbook.findall(".//{*}sheet"):
            name = sheet.get("name")
            target = relationships.get(sheet.get(_REL_NS + "id"), "")
            sheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(f"xl/{target}")
            if name is not None and sheet_path in result:
                result_by_sheet[name] = result[sheet_path]
        return result_by_sheet


def _excel_table(sheet: Worksheet, embeddings: dict[tuple[int, int], tuple[str, ...]]) -> Table:
    rows: dict[int, dict[int, Cell]] = {}
    # Visit stored cells so formatting an entire column cannot force a dense scan.
    for (row, column), cell in sheet._cells.items():
        if cell.value is None:
            continue
        text = str(cell.value).strip()
        if text:
            origin = (row - 1, column - 1)
            rows.setdefault(origin[0], {})[origin[1]] = Cell(text, origin, attachments=_filenames(text))

    images = set()
    for image in sheet._images:
        anchor = getattr(image.anchor, "_from", None)
        if anchor is not None:
            images.add((anchor.row, anchor.col))

    merged_ranges = tuple(sheet.merged_cells.ranges)
    for row, column in images | embeddings.keys():
        origin = next(
            ((region.min_row - 1, region.min_col - 1) for region in merged_ranges
             if region.min_row <= row + 1 <= region.max_row and region.min_col <= column + 1 <= region.max_col),
            (row, column),
        )
        cells = rows.setdefault(origin[0], {})
        cell = cells.get(origin[1], Cell(origin=origin))
        files = tuple(dict.fromkeys(cell.attachments + embeddings.get((row, column), ())))
        cells[origin[1]] = replace(cell, has_image=cell.has_image or (row, column) in images, attachments=files)

    for region in merged_ranges:
        row_start, column_start = region.min_row - 1, region.min_col - 1
        cell = rows.get(row_start, {}).get(column_start)
        if cell is None:
            continue
        for row in range(row_start, region.max_row):
            cells = rows.setdefault(row, {})
            columns = range(column_start, region.max_col)
            cells.update(zip(columns, repeat(cell)))
    return Table(sheet.title, rows)


def read_excel(path: Path) -> list[Table]:
    embeddings = _embeddings(path)
    with path.open("rb") as source:
        workbook = load_workbook(source, data_only=True, keep_links=False)
        try:
            return [_excel_table(sheet, embeddings.get(sheet.title, {})) for sheet in workbook.worksheets]
        finally:
            workbook.close()


def read_word(path: Path) -> list[Table]:
    document = open_word(path)
    tables = []
    for index, table in enumerate(document.tables, start=1):
        rows: dict[int, dict[int, Cell]] = {}
        read_cells: dict[object, Cell] = {}
        for row_index, row in enumerate(table.rows):
            cells = rows.setdefault(row_index, {})
            for column, cell in enumerate(row.cells, start=row.grid_cols_before):
                if cell._tc not in read_cells:
                    text = cell.text.strip()
                    embedded = tuple(
                        PurePosixPath(relation.target_ref).name
                        for node in cell._tc.xpath(".//*[local-name()='OLEObject']")
                        if (relation := cell.part.rels.get(node.get(_REL_NS + "id"))) is not None
                        and not relation.is_external
                    )
                    files = tuple(dict.fromkeys(_filenames(text) + embedded))
                    image = bool(cell._tc.xpath(".//w:drawing | .//w:pict"))
                    read_cells[cell._tc] = Cell(text, cell._tc, image, files)
                cells[column] = read_cells[cell._tc]
        tables.append(Table(f"Table {index}", rows))
    return tables
