#!/usr/bin/env python3

import argparse
import json
import math
import posixpath
import re
import sys
from typing import Dict, Set, Tuple, Union
from urllib.parse import unquote
import xml.etree.ElementTree as ElementTree
import zipfile

MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLUMNS = 16_384
CELL_REFERENCE = re.compile(r"^([A-Z]+)([1-9][0-9]*)$")

CellValue = Union[str, float]
Coordinate = Tuple[int, int]


def positive_integer(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate the semantic workload in a generated XLSX file."
    )
    parser.add_argument("--file", required=True)
    parser.add_argument("--rows", required=True, type=positive_integer)
    parser.add_argument("--columns", required=True, type=positive_integer)
    args = parser.parse_args()

    if args.rows > MAX_EXCEL_ROWS:
        parser.error(f"rows must not exceed Excel's limit of {MAX_EXCEL_ROWS}")
    if args.columns > MAX_EXCEL_COLUMNS:
        parser.error(f"columns must not exceed Excel's limit of {MAX_EXCEL_COLUMNS}")
    return args


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def relationship_id(sheet: ElementTree.Element) -> str:
    for attribute, value in sheet.attrib.items():
        if local_name(attribute) == "id":
            return value
    raise ValueError("Data worksheet has no relationship id")


def normalize_part_path(base_part: str, target: str) -> str:
    decoded = unquote(target).replace("\\", "/")
    if decoded.startswith("/"):
        candidate = decoded.lstrip("/")
    else:
        candidate = posixpath.join(posixpath.dirname(base_part), decoded)
    normalized = posixpath.normpath(candidate)
    if normalized == ".." or normalized.startswith("../"):
        raise ValueError(f"relationship target escapes the XLSX archive: {target}")
    return normalized


def worksheet_part(archive: zipfile.ZipFile) -> Tuple[str, str]:
    workbook_part = "xl/workbook.xml"
    workbook = ElementTree.fromstring(archive.read(workbook_part))
    sheet = next(
        (
            element
            for element in workbook.iter()
            if local_name(element.tag) == "sheet" and element.attrib.get("name") == "Data"
        ),
        None,
    )
    if sheet is None:
        raise ValueError("workbook does not contain a worksheet named Data")

    relation_id = relationship_id(sheet)
    relationships = ElementTree.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )
    for relationship in relationships:
        if relationship.attrib.get("Id") == relation_id:
            if relationship.attrib.get("TargetMode") == "External":
                raise ValueError("Data worksheet points to an external relationship")
            return normalize_part_path(
                workbook_part, relationship.attrib["Target"]
            ), workbook_part
    raise ValueError(f"worksheet relationship not found: {relation_id}")


def shared_strings_part(archive: zipfile.ZipFile, workbook_part: str) -> str:
    relationships = ElementTree.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )
    for relationship in relationships:
        if relationship.attrib.get("Type", "").endswith("/sharedStrings"):
            return normalize_part_path(workbook_part, relationship.attrib["Target"])
    fallback = "xl/sharedStrings.xml"
    if fallback in archive.namelist():
        return fallback
    raise ValueError("worksheet references shared strings but no shared string table exists")


def column_name(one_based_column: int) -> str:
    result = []
    while one_based_column:
        one_based_column, remainder = divmod(one_based_column - 1, 26)
        result.append(chr(ord("A") + remainder))
    return "".join(reversed(result))


def column_index(name: str) -> int:
    result = 0
    for character in name:
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def sample_coordinates(rows: int, columns: int) -> Set[Coordinate]:
    sample_rows = {0, rows // 2, rows - 1}
    sample_columns = {0, columns // 2, columns - 1}
    if columns > 1:
        sample_columns.add(1)
    return {
        (row, column)
        for row in sample_rows
        for column in sample_columns
    }


def child_text(element: ElementTree.Element, child_name: str) -> str:
    child = next(
        (child for child in element.iter() if local_name(child.tag) == child_name),
        None,
    )
    if child is None:
        raise ValueError(
            f"cell is missing its {child_name} element"
        )
    return child.text or ""


def parse_sample_value(
    cell: ElementTree.Element,
) -> Union[CellValue, Tuple[str, int]]:
    data_type = cell.attrib.get("t")
    if data_type == "inlineStr":
        return "".join(
            child.text or ""
            for child in cell.iter()
            if local_name(child.tag) == "t"
        )
    raw_value = child_text(cell, "v")
    if data_type == "s":
        return ("shared", int(raw_value))
    if data_type == "str":
        return raw_value
    if data_type not in (None, "n"):
        raise ValueError(f"unsupported sampled cell type: {data_type}")
    return float(raw_value)


def load_shared_strings(
    archive: zipfile.ZipFile, part: str, needed: Set[int]
) -> Dict[int, str]:
    resolved: Dict[int, str] = {}
    current_index = 0
    with archive.open(part) as stream:
        for _, element in ElementTree.iterparse(stream, events=("end",)):
            if local_name(element.tag) != "si":
                continue
            if current_index in needed:
                resolved[current_index] = "".join(
                    child.text or ""
                    for child in element.iter()
                    if local_name(child.tag) == "t"
                )
                if len(resolved) == len(needed):
                    break
            current_index += 1
            element.clear()
    missing = needed.difference(resolved)
    if missing:
        raise ValueError(f"shared string indexes not found: {sorted(missing)}")
    return resolved


def expected_value(row: int, column: int) -> CellValue:
    if column % 2 == 0:
        return f"Row {row} Col {column}"
    return row * column * 0.123


def assert_sample(coordinate: Coordinate, actual: CellValue) -> None:
    expected = expected_value(*coordinate)
    if isinstance(expected, str):
        if actual != expected:
            raise ValueError(
                f"unexpected value at row {coordinate[0]}, column {coordinate[1]}: "
                f"expected {expected!r}, got {actual!r}"
            )
        return
    if not isinstance(actual, float) or not math.isclose(
        actual, expected, rel_tol=1e-12, abs_tol=1e-12
    ):
        raise ValueError(
            f"unexpected value at row {coordinate[0]}, column {coordinate[1]}: "
            f"expected {expected!r}, got {actual!r}"
        )


def validate(filename: str, rows: int, columns: int) -> dict:
    samples = sample_coordinates(rows, columns)
    target_references = {
        f"{column_name(column + 1)}{row + 1}": (row, column)
        for row, column in samples
    }
    sampled_values: Dict[Coordinate, Union[CellValue, Tuple[str, int]]] = {}

    with zipfile.ZipFile(filename) as archive:
        worksheet, workbook = worksheet_part(archive)
        row_count = 0
        cell_count = 0
        cells_in_row = 0

        with archive.open(worksheet) as stream:
            for event, element in ElementTree.iterparse(
                stream, events=("start", "end")
            ):
                element_name = local_name(element.tag)
                if event == "start" and element_name == "row":
                    row_count += 1
                    row_number = int(element.attrib.get("r", row_count))
                    if row_number != row_count:
                        raise ValueError(
                            f"worksheet row sequence is not contiguous at row {row_count}"
                        )
                    cells_in_row = 0
                    continue

                if event == "end" and element_name == "c":
                    cell_count += 1
                    cells_in_row += 1
                    reference = element.attrib.get("r")
                    match = CELL_REFERENCE.fullmatch(reference or "")
                    if match is None:
                        raise ValueError(f"invalid or missing cell reference: {reference!r}")
                    cell_column = column_index(match.group(1))
                    cell_row = int(match.group(2))
                    if cell_row != row_count or cell_column != cells_in_row:
                        raise ValueError(
                            f"worksheet cells are not contiguous at {reference}"
                        )
                    coordinate = target_references.get(reference)
                    if coordinate is not None:
                        sampled_values[coordinate] = parse_sample_value(element)
                    element.clear()
                    continue

                if event == "end" and element_name == "row":
                    if cells_in_row != columns:
                        raise ValueError(
                            f"row {row_count} contains {cells_in_row} cells instead of {columns}"
                        )
                    element.clear()

        expected_cells = rows * columns
        if row_count != rows:
            raise ValueError(f"worksheet contains {row_count} rows instead of {rows}")
        if cell_count != expected_cells:
            raise ValueError(
                f"worksheet contains {cell_count} cells instead of {expected_cells}"
            )
        missing_samples = samples.difference(sampled_values)
        if missing_samples:
            raise ValueError(f"sampled cells are missing: {sorted(missing_samples)}")

        shared_indexes = {
            value[1]
            for value in sampled_values.values()
            if isinstance(value, tuple) and value[0] == "shared"
        }
        shared_strings = (
            load_shared_strings(
                archive, shared_strings_part(archive, workbook), shared_indexes
            )
            if shared_indexes
            else {}
        )

    for coordinate, value in sampled_values.items():
        if isinstance(value, tuple):
            value = shared_strings[value[1]]
        assert_sample(coordinate, value)

    return {
        "worksheet": "Data",
        "rows": rows,
        "columns": columns,
        "cells": rows * columns,
        "samples": len(samples),
    }


def main() -> None:
    args = parse_args()
    print(json.dumps(validate(args.file, args.rows, args.columns), separators=(",", ":")))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Workbook validation failed: {error}", file=sys.stderr)
        raise SystemExit(1) from error
