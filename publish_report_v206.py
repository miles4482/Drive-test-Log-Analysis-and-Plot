#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v2.06 — Excel-openable Active Session-ID workbook.

v1.98–v2.05 are not well-formed OpenXML, and Excel reports them as corrupt:

* ``xl/workbook.xml`` — the fresh openpyxl workbook declares ``xmlns:r`` on each
  ``<sheet>`` tag, not on ``<workbook>``. ``_ensure_ep_sheet`` inserted
  ``<sheet name="EP Details" … r:id="rId17"/>`` with no ``xmlns:r`` in scope.
* ``xl/drawings/drawing6.xml`` (Bad Spot RSRP_Active) — the fresh openpyxl
  drawing root declares only the default namespace; ``embed_circle_shapes``
  appended 122 ``<sp>`` circles that use ``a:`` with no ``xmlns:a`` in scope.

v1.94 (from v1.90) is the last package Excel opened. Its package rules are kept:
``_rels/.rels``, ``docProps/app.xml``, the theme and the ``workbook.xml``
frame come from v1.94; ``workbook.xml`` and every drawing declare ``xmlns:r`` /
``xmlns:a`` once on the root, like v1.94. Worksheets, charts, media and styles
come from v2.05 (Active Session-ID content, 97 Active bad spots + Scanner
section, EP Details, native legend-coloured Cover charts) with these repairs:

* rows with ``customHeight="52"`` (a regex hit the tail of ``customHeight``) get
  ``customHeight="1"``
* empty ``t="inlineStr"`` cells with no ``<is>`` become plain styled blanks
* ``<font>`` children are written in the order the schema lists them
* Cover: the Active Dataset / Radio KPI tables move up beside the Active charts
  (v1.98 moved the charts and titles but left Idle numbers under "Active Mode")

Chart XML and media bytes are unchanged from v2.05.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

SRC_CONTENT = Path("output/Bogura_DriveTest_Report_v2.05.xlsx")
SRC_PACKAGE = Path("output/Bogura_DriveTest_Report_v1.94.xlsx")
DEST = Path("output/Bogura_DriveTest_Report_v2.06.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

FROM_PACKAGE = ("_rels/.rels", "docProps/app.xml", "xl/theme/theme1.xml")

CONTENT_TYPES = (
    (r"xl/workbook\.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
    (r"xl/worksheets/sheet\d+\.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
    (r"xl/drawings/drawing\d+\.xml", "application/vnd.openxmlformats-officedocument.drawing+xml"),
    (r"xl/charts/chart\d+\.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
    (r"xl/styles\.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
    (r"xl/theme/theme\d+\.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
    (r"docProps/core\.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    (r"docProps/app\.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
)

FONT_ORDER = (
    "b", "i", "strike", "condense", "extend", "outline", "shadow", "u",
    "vertAlign", "sz", "color", "name", "family", "charset", "scheme",
)


def _hoist(xml: bytes, nsmap: dict[str, str]) -> bytes:
    """Declare the given prefixes once on the root and drop the local copies."""
    root = etree.fromstring(xml)
    etree.cleanup_namespaces(root, top_nsmap=nsmap)
    return etree.tostring(root, xml_declaration=False, encoding="UTF-8")


def _declare_on_root(xml: bytes, root_tag: str, nsmap: dict[str, str]) -> bytes:
    """Bind prefixes that the part uses without declaring, so it can be parsed."""
    text = xml.decode("utf-8")
    match = re.match(rf"(<{root_tag}\b)([^>]*)>", text)
    if match is None:
        raise SystemExit(f"<{root_tag}> root not found")
    attrs = match.group(2)
    extra = "".join(
        f' xmlns:{prefix}="{uri}"' for prefix, uri in nsmap.items() if f"xmlns:{prefix}=" not in attrs
    )
    return (text[: match.end(2)] + extra + text[match.end(2):]).encode("utf-8")


def _same_tree(a: etree._Element, b: etree._Element) -> bool:
    if a.tag != b.tag or dict(a.attrib) != dict(b.attrib):
        return False
    if (a.text or "") != (b.text or "") or (a.tail or "") != (b.tail or ""):
        return False
    if len(a) != len(b):
        return False
    return all(_same_tree(x, y) for x, y in zip(a, b))


def build_workbook(content_wb: bytes, package_wb: bytes) -> bytes:
    """v1.94 workbook frame with the v2.05 sheet list, r: declared on the root."""
    content = etree.fromstring(_declare_on_root(content_wb, "workbook", {"r": REL}))
    sheets = []
    for sheet in content.find(f"{{{MAIN}}}sheets"):
        sheets.append(
            f'<sheet name="{_attr(sheet.attrib["name"])}" sheetId="{sheet.attrib["sheetId"]}" '
            f'state="{sheet.attrib.get("state", "visible")}" r:id="{sheet.attrib[f"{{{REL}}}id"]}" />'
        )
    frame = package_wb.decode("utf-8")
    if not frame.startswith(f'<workbook xmlns:r="{REL}" xmlns="{MAIN}">'):
        raise SystemExit("v1.94 workbook frame changed; refusing to graft")
    body = re.sub(r"<sheets>.*?</sheets>", "<sheets>" + "".join(sheets) + "</sheets>", frame, count=1, flags=re.S)
    return body.encode("utf-8")


def _attr(value: str) -> str:
    return value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")


def repair_worksheet(xml: bytes) -> tuple[bytes, int, int]:
    text = xml.decode("utf-8")
    text, heights = re.subn(r'\bcustomHeight="(?!0"|1"|true"|false")[^"]*"', 'customHeight="1"', text)
    text, blanks = re.subn(
        r'<c (r="[A-Z]+\d+")( s="\d+")? t="inlineStr"\s*(?:/>|></c>)',
        lambda m: f"<c {m.group(1)}{m.group(2) or ''} />",
        text,
    )
    return text.encode("utf-8"), heights, blanks


def order_fonts(styles: bytes) -> bytes:
    root = etree.fromstring(styles)
    rank = {f"{{{MAIN}}}{name}": index for index, name in enumerate(FONT_ORDER)}
    for font in root.iter(f"{{{MAIN}}}font"):
        children = list(font)
        ordered = sorted(children, key=lambda el: rank.get(el.tag, len(rank)))
        if ordered != children:
            for el in children:
                font.remove(el)
            font.extend(ordered)
    return etree.tostring(root, xml_declaration=False, encoding="UTF-8")


def _col_number(letters: str) -> int:
    value = 0
    for ch in letters:
        value = value * 26 + ord(ch) - 64
    return value


def _set_cell_text(xml: str, ref: str, text: str) -> str:
    pattern = rf'(<c r="{ref}"[^>]*t="inlineStr"[^>]*><is><t[^>]*>)[^<]*(</t></is></c>)'
    xml, count = re.subn(pattern, lambda m: m.group(1) + text + m.group(2), xml, count=1)
    if count != 1:
        raise SystemExit(f"Cover cell {ref} not found for relabel")
    return xml


def fix_cover_blocks(xml: bytes) -> bytes:
    """Put the Active dataset/KPI tables beside the Active charts at the top of Cover.

    v1.98 swapped the Cover chart anchors and titles (Active charts to rows 5/22)
    but left the Dataset / Radio KPI tables where they were, so "Active Mode"
    listed the Idle scanner archive and "Scanner (Idle Mode)" listed Active.
    Columns A–K of rows 5–17 and 44–56 swap; chart headings in L+ stay put.
    """
    text = xml.decode("utf-8")
    top, bottom, offset, last_col = range(5, 18), range(44, 57), 39, _col_number("K")
    if "Data_active_session ID_Bogra.xlsb" not in re.search(r'<row r="47".*?</row>', text, flags=re.S).group(0):
        raise SystemExit("Cover lower block no longer holds the Active dataset; refusing to swap")

    rows: dict[int, re.Match] = {}
    for match in re.finditer(r'<row r="(\d+)"([^>]*)>(.*?)</row>', text, flags=re.S):
        number = int(match.group(1))
        if number in top or number in bottom:
            rows[number] = match
    if sorted(rows) != list(top) + list(bottom):
        raise SystemExit("Cover dataset rows are not all present")

    def split(number: int) -> tuple[list[tuple[int, str]], list[tuple[int, str]]]:
        left, right = [], []
        for cell in re.findall(r'<c r="[A-Z]+\d+"[^>]*?(?:/>|>.*?</c>)', rows[number].group(3), flags=re.S):
            letters = re.match(r'<c r="([A-Z]+)', cell).group(1)
            (left if _col_number(letters) <= last_col else right).append((_col_number(letters), cell))
        return left, right

    def move(cells: list[tuple[int, str]], src: int, dst: int) -> list[tuple[int, str]]:
        return [(col, re.sub(rf'<c r="([A-Z]+){src}"', rf'<c r="\g<1>{dst}"', cell, count=1)) for col, cell in cells]

    new_rows: dict[int, str] = {}
    for upper in top:
        lower = upper + offset
        up_left, up_right = split(upper)
        low_left, low_right = split(lower)
        for number, keep, incoming, src in (
            (upper, up_right, low_left, lower),
            (lower, low_right, up_left, upper),
        ):
            cells = sorted(move(incoming, src, number) + keep)
            match = rows[number]
            new_rows[number] = f'<row r="{number}"{match.group(2)}>' + "".join(c for _, c in cells) + "</row>"

    pieces, last = [], 0
    for number in sorted(rows):
        match = rows[number]
        pieces.append(text[last:match.start()])
        pieces.append(new_rows[number])
        last = match.end()
    pieces.append(text[last:])
    text = "".join(pieces)

    def shift_merge(match: re.Match) -> str:
        a_col, a_row, b_col, b_row = match.group(1), int(match.group(2)), match.group(3), int(match.group(4))
        if _col_number(b_col) > last_col:
            return match.group(0)
        if a_row in top and b_row in top:
            a_row, b_row = a_row + offset, b_row + offset
        elif a_row in bottom and b_row in bottom:
            a_row, b_row = a_row - offset, b_row - offset
        return f'<mergeCell ref="{a_col}{a_row}:{b_col}{b_row}"'

    text = re.sub(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', shift_merge, text)

    text = _set_cell_text(text, "B7", "Active (UE)")
    text = _set_cell_text(text, "B46", "Scanner (Idle)")
    text = _set_cell_text(text, "D44", "Radio KPIs (all samples)")
    text = _set_cell_text(text, "L44", "Scanner (Idle) histograms (native Excel charts with legends)")
    return text.encode("utf-8")


def content_types(names: list[str], package_ct: bytes) -> bytes:
    defaults = re.findall(r"<Default [^>]*/>", package_ct.decode("utf-8"))
    overrides = []
    for name in names:
        for pattern, kind in CONTENT_TYPES:
            if re.fullmatch(pattern, name):
                overrides.append(f'<Override PartName="/{name}" ContentType="{kind}" />')
                break
    return (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + "".join(defaults)
        + "".join(overrides)
        + "</Types>"
    ).encode("utf-8")


def publish() -> Path:
    for path in (SRC_CONTENT, SRC_PACKAGE):
        if not path.exists():
            raise SystemExit(f"Missing {path}")
    with zipfile.ZipFile(SRC_CONTENT) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    with zipfile.ZipFile(SRC_PACKAGE) as zf:
        package = {name: zf.read(name) for name in (*FROM_PACKAGE, "xl/workbook.xml", "[Content_Types].xml")}

    for name in FROM_PACKAGE:
        parts[name] = package[name]

    parts["xl/workbook.xml"] = build_workbook(parts["xl/workbook.xml"], package["xl/workbook.xml"])

    for name in names:
        if re.fullmatch(r"xl/drawings/drawing\d+\.xml", name):
            bound = _declare_on_root(parts[name], "wsDr", {"a": DML, "r": REL})
            hoisted = _hoist(bound, {None: XDR, "a": DML, "r": REL})
            if not _same_tree(etree.fromstring(bound), etree.fromstring(hoisted)):
                raise SystemExit(f"{name}: namespace hoist changed the drawing")
            parts[name] = hoisted

    parts["docProps/core.xml"] = _hoist(
        parts["docProps/core.xml"],
        {
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "dc": "http://purl.org/dc/elements/1.1/",
            "dcterms": "http://purl.org/dc/terms/",
            "xsi": "http://www.w3.org/2001/XMLSchema-instance",
        },
    )

    fixed_heights = fixed_blanks = 0
    for name in names:
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
            parts[name], heights, blanks = repair_worksheet(parts[name])
            fixed_heights += heights
            fixed_blanks += blanks

    parts["xl/worksheets/sheet1.xml"] = fix_cover_blocks(parts["xl/worksheets/sheet1.xml"])
    parts["xl/styles.xml"] = order_fonts(parts["xl/styles.xml"])
    parts["[Content_Types].xml"] = content_types(names, package["[Content_Types].xml"])

    DEST.parent.mkdir(parents=True, exist_ok=True)
    tmp = DEST.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name], compress_type=zipfile.ZIP_DEFLATED)
    tmp.replace(DEST)
    print(
        f"Wrote {DEST} ({DEST.stat().st_size / 1e6:.1f} MB): "
        f"{fixed_heights} customHeight rows, {fixed_blanks} empty inlineStr cells repaired",
        flush=True,
    )
    return DEST


def verify(path: Path) -> list[str]:
    from validate_xlsx_package import check

    problems = check(path)
    with zipfile.ZipFile(path) as zf, zipfile.ZipFile(SRC_CONTENT) as content, zipfile.ZipFile(SRC_PACKAGE) as package:
        if sorted(zf.namelist()) != sorted(content.namelist()):
            problems.append("part names differ from v2.05")
        for name in content.namelist():
            if name.startswith(("xl/charts/", "xl/media/")) and zf.read(name) != content.read(name):
                problems.append(f"chart/media mutated: {name}")
            if name.startswith(("xl/drawings/_rels/", "xl/worksheets/_rels/")) and zf.read(name) != content.read(name):
                problems.append(f"rels mutated: {name}")
        for name in FROM_PACKAGE:
            if zf.read(name) != package.read(name):
                problems.append(f"{name} is not the v1.94 part")

        workbook = zf.read("xl/workbook.xml").decode("utf-8")
        if not workbook.startswith(f'<workbook xmlns:r="{REL}" xmlns="{MAIN}">'):
            problems.append("workbook root is not the v1.94 form")
        names = re.findall(r'<sheet name="([^"]+)"', workbook)
        expected = [
            "Cover", "RSRP", "RSRQ", "SINR", "Bad Spot RSRP_Active", "RSRP vs KPIs",
            "RSRP vs KPIs_Active", "Sample Log", "EP Details", "_ChartData",
        ]
        if names != expected:
            problems.append(f"sheet list {names}")
        for name in zf.namelist():
            if re.fullmatch(r"xl/drawings/drawing\d+\.xml", name):
                head = zf.read(name)[:400].decode("utf-8", errors="ignore")
                if f'xmlns:a="{DML}"' not in head.split(">", 1)[0] or f'xmlns:r="{REL}"' not in head.split(">", 1)[0]:
                    problems.append(f"{name}: root does not declare a: and r:")

        active = zf.read("xl/worksheets/sheet6.xml").decode("utf-8")
        active_titles = len(re.findall(r">Bad Spot \d+<", active))
        scanner_titles = len(re.findall(r">Scanner Bad Spot \d+<", active))
        if active_titles != 97:
            problems.append(f"Active bad spots {active_titles}, expected 97")
        if scanner_titles != 26:
            problems.append(f"Scanner bad spots {scanner_titles}, expected 26")
        if "Scanner Bad Spot<" not in active:
            problems.append("Scanner Bad Spot banner missing")
        drawing6 = zf.read("xl/drawings/drawing6.xml").decode("utf-8")
        if drawing6.count('prst="ellipse"') != 122:
            problems.append(f"Active circles {drawing6.count('prst=\"ellipse\"')}, expected 122")

        cover_sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        for row, needles in (
            (3, ("Active Mode",)),
            (7, ("Active (UE)",)),
            (8, ("Data_active_session ID_Bogra.xlsb",)),
            (42, ("Scanner (Idle Mode)",)),
            (46, ("Scanner (Idle)",)),
            (48, ("Borga_Idle_V2.0.7z",)),
        ):
            block = re.search(rf'<row r="{row}"[^>]*>.*?</row>', cover_sheet, flags=re.S)
            for needle in needles:
                if block is None or needle not in block.group(0):
                    problems.append(f"Cover row {row} missing {needle!r}")
        if '<mergeCell ref="D11:K12"/>' not in cover_sheet or "D50:K51" in cover_sheet:
            problems.append("Cover Active note merge not moved to D11:K12")

        cover = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
        if "<pic>" in cover and cover.count("<pic>") != 1:
            problems.append("Cover has extra pictures")
        for rid in ("rId5", "rId6", "rId7"):
            if f'r:id="{rid}"' not in cover:
                problems.append(f"Cover Active chart {rid} missing")
        for chart, series in (("chart5", 5), ("chart6", 3), ("chart7", 6)):
            xml = zf.read(f"xl/charts/{chart}.xml").decode("utf-8")
            if '<grouping val="stacked"/>' not in xml or xml.count("<ser>") != series:
                problems.append(f"{chart}: not a {series}-series stacked legend chart")
        if any(n.startswith("xl/media/image_cover_") for n in zf.namelist()):
            problems.append("Cover PNG histogram media present")
    return problems


if __name__ == "__main__":
    dest = publish()
    problems = verify(dest)
    if problems:
        for line in problems[:40]:
            print("  -", line)
        raise SystemExit(f"v2.06 verify failed ({len(problems)} problems)")
    shutil.copy2(dest, LATEST)
    print(f"v2.06 verified; copied to {LATEST}")
