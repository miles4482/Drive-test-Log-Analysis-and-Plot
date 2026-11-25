#!/usr/bin/env python3
"""Strict OpenXML checks for the Bogura report workbook.

Excel rejects a package as corrupt on problems that openpyxl and LibreOffice
accept, so this checks what Excel is strict about:

* every XML / rels part is well-formed with all namespace prefixes bound
* every relationship target exists; every r:id / r:embed resolves
* every part has a content type; no override points at a missing part
* workbook sheets, rIds and sheetIds are unique
* worksheet child elements follow the schema order
* rows and cells are strictly ascending, cell refs sit in their row
* merges do not overlap, are not single cells, and the count matches
* style indexes exist in cellXfs
* drawing anchors have from/ext|to/object/clientData in order, ids are unique
* chart formulas only reference sheets that exist

Usage: python3 validate_xlsx_package.py output/Bogura_DriveTest_Report_v2.06.xlsx [...]
"""
from __future__ import annotations

import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"

WORKSHEET_ORDER = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
    "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
    "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
    "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
    "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
    "webPublishItems", "tableParts", "extLst",
]
ANCHOR_OBJECTS = {"sp", "grpSp", "graphicFrame", "cxnSp", "pic", "contentPart"}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _col_index(letters: str) -> int:
    value = 0
    for ch in letters:
        value = value * 26 + ord(ch) - 64
    return value


def _ref(ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", ref)
    if match is None:
        raise ValueError(ref)
    return _col_index(match.group(1)), int(match.group(2))


def _rels_path(part: str) -> str:
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def _resolve(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part) if source_part else ""
    return posixpath.normpath(posixpath.join(base, target))


def _rels(parts: dict[str, bytes], part: str) -> dict[str, tuple[str, str, str]]:
    """rId -> (type, resolved target, target mode) for one source part."""
    name = "_rels/.rels" if part == "" else _rels_path(part)
    if name not in parts:
        return {}
    out = {}
    for rel in ET.fromstring(parts[name]):
        mode = rel.attrib.get("TargetMode", "Internal")
        target = rel.attrib["Target"]
        resolved = target if mode == "External" else _resolve(part, target)
        out[rel.attrib["Id"]] = (rel.attrib["Type"], resolved, mode)
    return out


def check(path: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad:
            problems.append(f"zip CRC failure in {bad}")
        names = zf.namelist()
        if len(names) != len(set(names)):
            problems.append("duplicate zip entries")
        parts = {name: zf.read(name) for name in names}

    trees: dict[str, ET.Element] = {}
    for name, body in parts.items():
        if name.endswith((".xml", ".rels")):
            try:
                trees[name] = ET.fromstring(body)
            except ET.ParseError as exc:
                problems.append(f"XML not well-formed: {name}: {exc}")
    if problems:
        return problems

    # Content types
    types = trees["[Content_Types].xml"]
    defaults = {d.attrib["Extension"].lower() for d in types if _local(d.tag) == "Default"}
    overrides = {o.attrib["PartName"].lstrip("/") for o in types if _local(o.tag) == "Override"}
    for name in names:
        if name == "[Content_Types].xml" or name.endswith("/"):
            continue
        if name not in overrides and name.rsplit(".", 1)[-1].lower() not in defaults:
            problems.append(f"no content type: {name}")
    for name in overrides:
        if name not in parts:
            problems.append(f"content type override for missing part: {name}")

    # Relationships resolve
    for name in names:
        if not name.endswith(".rels"):
            continue
        folder, rel_file = posixpath.split(name)
        source = posixpath.join(posixpath.dirname(folder), rel_file[: -len(".rels")]) if rel_file != ".rels" else ""
        if source and source not in parts:
            problems.append(f"rels for missing part: {name}")
        ids = [rel.attrib.get("Id") for rel in trees[name]]
        if len(ids) != len(set(ids)):
            problems.append(f"duplicate relationship Id in {name}")
        for rid, (_, target, mode) in _rels(parts, source).items():
            if mode != "External" and target not in parts:
                problems.append(f"{name} {rid} -> missing {target}")

    # Workbook
    wb = trees["xl/workbook.xml"]
    wb_rels = _rels(parts, "xl/workbook.xml")
    sheets = wb.find(f"{{{MAIN}}}sheets")
    sheet_names, sheet_ids, sheet_rids = [], [], []
    sheet_parts: dict[str, str] = {}
    for sheet in sheets:
        rid = sheet.attrib.get(f"{{{REL_NS}}}id")
        sheet_names.append(sheet.attrib["name"])
        sheet_ids.append(sheet.attrib["sheetId"])
        sheet_rids.append(rid)
        if rid not in wb_rels:
            problems.append(f"workbook sheet {sheet.attrib['name']!r} r:id {rid} has no relationship")
            continue
        sheet_parts[sheet.attrib["name"]] = wb_rels[rid][1]
        if len(sheet.attrib["name"]) > 31:
            problems.append(f"sheet name longer than 31 chars: {sheet.attrib['name']!r}")
    for label, values in (("name", [n.lower() for n in sheet_names]), ("sheetId", sheet_ids), ("r:id", sheet_rids)):
        if len(values) != len(set(values)):
            problems.append(f"duplicate workbook sheet {label}")
    worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    linked = {target for kind, target, _ in wb_rels.values() if kind == worksheet_type}
    if linked != set(sheet_parts.values()):
        problems.append(f"workbook worksheet rels {sorted(linked ^ set(sheet_parts.values()))} not matched to <sheet>")
    visible = [s for s in sheets if s.attrib.get("state", "visible") == "visible"]
    if not visible:
        problems.append("no visible sheet")

    styles = trees.get("xl/styles.xml")
    xf_count = 0
    if styles is not None:
        xfs = styles.find(f"{{{MAIN}}}cellXfs")
        xf_count = len(list(xfs)) if xfs is not None else 0
        problems.extend(_check_styles(styles))

    # Worksheets
    for sheet_name, part in sheet_parts.items():
        root = trees[part]
        order = [_local(child.tag) for child in root]
        positions = [WORKSHEET_ORDER.index(tag) if tag in WORKSHEET_ORDER else -1 for tag in order]
        if -1 in positions:
            problems.append(f"{sheet_name}: unknown worksheet child {order[positions.index(-1)]}")
        if positions != sorted(positions):
            problems.append(f"{sheet_name}: worksheet children out of order {order}")
        last_row = 0
        for row in root.iter(f"{{{MAIN}}}row"):
            r = int(row.attrib["r"])
            if r <= last_row:
                problems.append(f"{sheet_name}: row {r} after row {last_row}")
            last_row = r
            last_col = 0
            for cell in row:
                ref = cell.attrib.get("r")
                if ref is None:
                    continue
                col, cell_row = _ref(ref)
                if cell_row != r:
                    problems.append(f"{sheet_name}: cell {ref} inside row {r}")
                if col <= last_col:
                    problems.append(f"{sheet_name}: cell {ref} out of column order in row {r}")
                last_col = col
                style = int(cell.attrib.get("s", "0"))
                if xf_count and style >= xf_count:
                    problems.append(f"{sheet_name}: cell {ref} style {style} >= cellXfs {xf_count}")
                if cell.attrib.get("t") == "inlineStr" and cell.find(f"{{{MAIN}}}is") is None:
                    problems.append(f"{sheet_name}: inlineStr cell {ref} without <is>")
        if last_row > 1048576:
            problems.append(f"{sheet_name}: row beyond Excel limit")
        merges = root.find(f"{{{MAIN}}}mergeCells")
        if merges is not None:
            boxes = []
            for merge in merges:
                a, _, b = merge.attrib["ref"].partition(":")
                (c1, r1), (c2, r2) = _ref(a), _ref(b or a)
                if (c1, r1) == (c2, r2):
                    problems.append(f"{sheet_name}: single-cell merge {merge.attrib['ref']}")
                boxes.append((min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2), merge.attrib["ref"]))
            if merges.attrib.get("count") not in (None, str(len(boxes))):
                problems.append(f"{sheet_name}: mergeCells count {merges.attrib.get('count')} != {len(boxes)}")
            boxes.sort()
            for i, (r1, r2, c1, c2, ref) in enumerate(boxes):
                for s1, s2, d1, d2, other in boxes[i + 1:]:
                    if s1 > r2:
                        break
                    if not (d2 < c1 or d1 > c2):
                        problems.append(f"{sheet_name}: overlapping merges {ref} and {other}")
        sheet_rels = _rels(parts, part)
        for el in root.iter():
            rid = el.attrib.get(f"{{{REL_NS}}}id")
            if rid and rid not in sheet_rels:
                problems.append(f"{sheet_name}: <{_local(el.tag)}> r:id {rid} missing from rels")

    # Drawings
    for name in names:
        if not re.fullmatch(r"xl/drawings/drawing\d+\.xml", name):
            continue
        root = trees[name]
        drawing_rels = _rels(parts, name)
        ids = []
        for anchor in root:
            kind = _local(anchor.tag)
            children = [_local(c.tag) for c in anchor]
            if kind == "oneCellAnchor":
                ok = len(children) == 4 and children[:2] == ["from", "ext"] and children[2] in ANCHOR_OBJECTS and children[3] == "clientData"
            elif kind == "twoCellAnchor":
                ok = len(children) == 4 and children[:2] == ["from", "to"] and children[2] in ANCHOR_OBJECTS and children[3] == "clientData"
            elif kind == "absoluteAnchor":
                ok = len(children) == 4 and children[:2] == ["pos", "ext"] and children[2] in ANCHOR_OBJECTS and children[3] == "clientData"
            else:
                ok = False
            if not ok:
                problems.append(f"{name}: bad anchor {kind} children {children}")
        for el in root.iter():
            if _local(el.tag) == "cNvPr":
                ids.append(el.attrib.get("id"))
            for attr in (f"{{{REL_NS}}}id", f"{{{REL_NS}}}embed", f"{{{REL_NS}}}link"):
                rid = el.attrib.get(attr)
                if rid and rid not in drawing_rels:
                    problems.append(f"{name}: {attr.rsplit('}', 1)[1]} {rid} missing from rels")
        dupes = {i for i in ids if ids.count(i) > 1}
        if dupes:
            problems.append(f"{name}: duplicate cNvPr ids {sorted(dupes)[:5]}")

    # Charts reference existing sheets
    known = {n.lower() for n in sheet_names}
    for name in names:
        if not re.fullmatch(r"xl/charts/chart\d+\.xml", name):
            continue
        for el in trees[name].iter(f"{{{CHART}}}f"):
            formula = el.text or ""
            match = re.match(r"'?((?:[^'!]|'')+)'?!", formula)
            if match and match.group(1).replace("''", "'").lower() not in known:
                problems.append(f"{name}: formula references missing sheet {formula}")
    return problems


def _check_styles(styles: ET.Element) -> list[str]:
    problems = []
    counts = {}
    for group in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs"):
        node = styles.find(f"{{{MAIN}}}{group}")
        if node is None:
            continue
        counts[group] = len(list(node))
        if node.attrib.get("count") not in (None, str(counts[group])):
            problems.append(f"styles {group} count {node.attrib.get('count')} != {counts[group]}")
    custom_fmts = {
        int(fmt.attrib["numFmtId"])
        for fmt in styles.iterfind(f"{{{MAIN}}}numFmts/{{{MAIN}}}numFmt")
    }
    for group in ("cellStyleXfs", "cellXfs"):
        for index, xf in enumerate(styles.iterfind(f"{{{MAIN}}}{group}/{{{MAIN}}}xf")):
            for attr, target in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                value = int(xf.attrib.get(attr, "0"))
                if value >= counts.get(target, 0):
                    problems.append(f"styles {group}[{index}] {attr}={value} >= {target} {counts.get(target, 0)}")
            fmt = int(xf.attrib.get("numFmtId", "0"))
            if fmt >= 164 and fmt not in custom_fmts:
                problems.append(f"styles {group}[{index}] numFmtId={fmt} not defined")
            if group == "cellXfs" and int(xf.attrib.get("xfId", "0")) >= counts.get("cellStyleXfs", 1):
                problems.append(f"styles cellXfs[{index}] xfId out of range")
    return problems


if __name__ == "__main__":
    status = 0
    for arg in sys.argv[1:]:
        found = check(Path(arg))
        print(f"{arg}: {'OK' if not found else f'{len(found)} problem(s)'}")
        for line in found[:60]:
            print("  -", line)
        if len(found) > 60:
            print(f"  ... {len(found) - 60} more")
        status |= bool(found)
    sys.exit(status)
