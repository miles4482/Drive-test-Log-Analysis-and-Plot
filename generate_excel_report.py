#!/usr/bin/env python3
"""Build a Bogura drive-test Excel report with RSRP/RSRQ/SINR charts.

Worksheet rules for every sheet:
- freeze panes off
- gridlines off (screen and print)
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import hashlib
import math
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile

import cairosvg
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import MultipleLocator
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from openpyxl.chart.text import RichText
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from merge_and_analyze import (
    ACTIVE_CELL_TYPE,
    ACTIVE_ARCHIVE,
    ACTIVE_OPERATOR,
    OUTPUT_DIR,
    PLOTS_DIR,
    SITE_DB_PATH,
    MAP_RSRP_COLORS,
    MAP_RSRP_LABELS,
    MAP_RSRQ_COLORS,
    MAP_RSRQ_LABELS,
    MAP_SINR_COLORS,
    MAP_SINR_LABELS,
    REPORT_LAYERS,
    coverage_extent,
    downsample,
    earfcn_to_band,
    earfcn_to_layer,
    filter_by_layer,
    BAD_SPOT_MAX_ZOOMS,
    BAD_SPOT_RULES,
    consecutive_label,
    bad_spot_advice,
    find_bad_spots,
    logic_drift,
    LOGIC_LOCK,
    MAP_CIRCLE_PIXELS,
    min_consecutive_m,
    best_carrier,
    best_server,
    select_active_report,
    strongest_fix,
    strongest_location,
    _equirect_km,
    plot_rsrp_bad_spot_map,
    plot_road_map,
    plot_rsrp_coverage_map,
    plot_rsrq_bad_spot_map,
    plot_rsrq_coverage_map,
    plot_sinr_bad_spot_map,
    plot_sinr_coverage_map,
    plot_bad_spot_serial_maps,
    plot_bad_spot_zoom_grid,
    plot_stacked_map_hist,
    rsrp_map_class,
    rsrp_range_counts,
    rsrq_map_class,
    sinr_map_class,
)

# Raise this on every report write. The versioned filename is the name that
# gets downloaded, so a new Excel or web report must not reuse the previous number.
REPORT_VERSION = "1.61"
SVG_DIR = PLOTS_DIR / "svg"
SVG_FALLBACK_SCALE = 3.0
SVG_BLIP_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
SVG_BLIP_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
# PNG fallback -> SVG it was rendered from; these pictures get an SVG layer in the workbook.
SVG_LAYERS: dict[Path, Path] = {}
# Every poor run is a bad spot, so a dense KPI produces over a thousand of them.
# The map marks them all; the table lists the longest ones a team can act on.
BAD_SPOT_TABLE_MAX = 300
REPORT_XLSX = OUTPUT_DIR / "Bogura_DriveTest_Report.xlsx"
VERSIONED_XLSX = OUTPUT_DIR / f"Bogura_DriveTest_Report_v{REPORT_VERSION}.xlsx"

NAVY = "1B4F72"
WHITE = "FFFFFF"
LIGHT = "F4F6F7"
CARD = "EAF2F8"
YELLOW = "F4D03F"
RSRP_COLOR = "C0392B"
RSRQ_COLOR = "1F618D"
SINR_COLOR = "117A65"
VS_SINR_LINE = "48C9B0"
VS_RSRQ_LINE = "5DADE2"
THIN = Border(
    left=Side(style="thin", color="D5D8DC"),
    right=Side(style="thin", color="D5D8DC"),
    top=Side(style="thin", color="D5D8DC"),
    bottom=Side(style="thin", color="D5D8DC"),
)


def apply_sheet_view(ws: Worksheet) -> None:
    """No freeze panes, no gridlines (on-screen or printed)."""
    ws.freeze_panes = None
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.view = "normal"


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def font(size=11, bold=False, color=NAVY, name="Calibri") -> Font:
    return Font(name=name, size=size, bold=bold, color=color)


def write_cell(ws, row, col, value, *, size=11, bold=False, color=NAVY, fill_color=None, align="left", num_fmt=None, wrap=False):
    cell = ws.cell(row, col, value)
    cell.font = font(size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill_color:
        cell.fill = fill(fill_color)
    if num_fmt:
        cell.number_format = num_fmt
    return cell


def set_widths(ws, widths: dict[str, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def kpi_stats(series: pd.Series) -> dict[str, float]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return {k: np.nan for k in ("Count", "Min", "P5", "Median", "Mean", "P95", "Max", "Std")}
    return {
        "Count": float(len(s)),
        "Min": float(s.min()),
        "P5": float(s.quantile(0.05)),
        "Median": float(s.median()),
        "Mean": float(s.mean()),
        "P95": float(s.quantile(0.95)),
        "Max": float(s.max()),
        "Std": float(s.std()),
    }


def map_range_counts(series: pd.Series, classifier, labels) -> pd.Series:
    return classifier(series).value_counts(dropna=True).reindex(labels, fill_value=0)


def add_image(ws: Worksheet, path: Path, anchor: str, width: int, height: int) -> tuple[int, int]:
    """Place the picture inside a width x height box without stretching it; returns its size."""
    img = XLImage(str(path))
    fit = min(width / img.width, height / img.height)
    img.width = round(img.width * fit)
    img.height = round(img.height * fit)
    ws.add_image(img, anchor)
    return img.width, img.height


def svg_with_fallback(svg_path: Path | None, png_path: Path) -> Path | None:
    """Render the PNG that viewers without SVG support show, and pair it with its SVG."""
    if svg_path is None:
        return None
    cairosvg.svg2png(
        url=str(svg_path), write_to=str(png_path), scale=SVG_FALLBACK_SCALE, background_color="white"
    )
    SVG_LAYERS[png_path.resolve()] = svg_path
    return png_path


def embed_svg_layers(path: Path) -> int:
    """Attach each paired SVG to its picture as an Office svgBlip.

    Excel 365 then draws the vector map; older Excel and other viewers keep the PNG.
    Returns the number of pictures that received an SVG layer.
    """
    by_digest = {
        hashlib.sha1(png.read_bytes()).hexdigest(): svg for png, svg in SVG_LAYERS.items() if png.exists()
    }
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    svg_for_media: dict[str, str] = {}
    for name in names:
        if name.startswith("xl/media/") and name.lower().endswith(".png"):
            svg = by_digest.get(hashlib.sha1(parts[name]).hexdigest())
            if svg is not None:
                svg_name = name[:-4] + ".svg"
                parts[svg_name] = svg.read_bytes()
                svg_for_media[name] = svg_name
    embedded = 0
    for rels_name in [n for n in names if re.fullmatch(r"xl/drawings/_rels/drawing\d+\.xml\.rels", n)]:
        drawing_name = "xl/drawings/" + rels_name.rsplit("/", 1)[1][: -len(".rels")]
        rels = parts[rels_name].decode("utf-8")
        drawing = parts[drawing_name].decode("utf-8")
        next_id = max((int(i) for i in re.findall(r'Id="rId(\d+)"', rels)), default=0) + 1
        added = []
        for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
            rid = re.search(r'Id="([^"]+)"', rel).group(1)
            target = re.search(r'Target="([^"]+)"', rel).group(1)
            media = target.lstrip("/") if target.startswith("/") else "xl/media/" + target.rsplit("/", 1)[1]
            svg_name = svg_for_media.get(media)
            if svg_name is None:
                continue
            svg_rid = f"rId{next_id}"
            next_id += 1
            svg_target = "/" + svg_name if target.startswith("/") else "../media/" + svg_name.rsplit("/", 1)[1]
            added.append(f'<Relationship Type="{IMAGE_REL_TYPE}" Target="{svg_target}" Id="{svg_rid}" />')
            blip = re.compile(r'<a:blip\b([^>]*\br:embed="' + re.escape(rid) + r'"[^>]*?)\s*/>')
            drawing, hits = blip.subn(
                lambda m: (
                    f"<a:blip{m.group(1)}><a:extLst><a:ext uri=\"{SVG_BLIP_EXT_URI}\">"
                    f"<asvg:svgBlip xmlns:asvg=\"{SVG_BLIP_NS}\" r:embed=\"{svg_rid}\" />"
                    "</a:ext></a:extLst></a:blip>"
                ),
                drawing,
            )
            embedded += hits
        if added:
            parts[rels_name] = rels.replace("</Relationships>", "".join(added) + "</Relationships>").encode("utf-8")
            parts[drawing_name] = drawing.encode("utf-8")
    types = parts["[Content_Types].xml"].decode("utf-8")
    if svg_for_media and 'Extension="svg"' not in types:
        types = types.replace("<Override ", '<Default Extension="svg" ContentType="image/svg+xml" /><Override ', 1)
        parts["[Content_Types].xml"] = types.encode("utf-8")
    order = names + [n for n in parts if n not in names]
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            zf.writestr(name, parts[name])
    tmp.replace(path)
    return embedded


def _sheet_drawings(parts: dict[str, bytes]) -> dict[str, str]:
    """Map a worksheet name to its drawing part, when the sheet has one."""
    main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    wb = ET.fromstring(parts["xl/workbook.xml"])
    book_rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rid_target = {rel.attrib["Id"]: rel.attrib["Target"].lstrip("/") for rel in book_rels}
    found = {}
    for sheet in wb.find(f"{main}sheets"):
        sheet_path = rid_target[sheet.attrib[f"{rel_ns}id"]]
        rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if rels_path not in parts:
            continue
        for rel in ET.fromstring(parts[rels_path]):
            target = rel.attrib["Target"].lstrip("/")
            if "drawings/drawing" in target:
                found[sheet.attrib["name"]] = target
    return found


def embed_circle_shapes(path: Path) -> int:
    """Add one Excel oval per bad-spot circle, anchored on top of its map picture.

    Each oval is its own drawing object, so it can be dragged without moving the
    map or the other circles. Ovals are added only on Bad Spot sheets.
    """
    if not MAP_CIRCLE_PIXELS:
        return 0
    by_digest = {}
    for png, geom in MAP_CIRCLE_PIXELS.items():
        file = Path(png)
        if file.exists() and geom.get("circles"):
            by_digest[hashlib.sha1(file.read_bytes()).hexdigest()] = geom
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    media_digest = {
        name: hashlib.sha1(parts[name]).hexdigest()
        for name in names
        if name.startswith("xl/media/") and name.lower().endswith(".png")
    }
    placed = 0
    bad_spot_drawings = {
        drawing for sheet, drawing in _sheet_drawings(parts).items() if sheet.startswith("Bad Spot")
    }
    for rels_name in [n for n in names if re.fullmatch(r"xl/drawings/_rels/drawing\d+\.xml\.rels", n)]:
        drawing_name = "xl/drawings/" + rels_name.rsplit("/", 1)[1][: -len(".rels")]
        if drawing_name not in bad_spot_drawings:
            continue
        rels = parts[rels_name].decode("utf-8")
        drawing = parts[drawing_name].decode("utf-8")
        rid_media = {}
        for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
            rid = re.search(r'Id="([^"]+)"', rel).group(1)
            target = re.search(r'Target="([^"]+)"', rel).group(1)
            media = target.lstrip("/") if target.startswith("/") else "xl/media/" + target.rsplit("/", 1)[1]
            rid_media[rid] = media
        next_shape = max((int(i) for i in re.findall(r'<cNvPr id="(\d+)"', drawing)), default=0) + 1
        extras = []
        for block in re.findall(r"<oneCellAnchor>.*?</oneCellAnchor>", drawing, flags=re.S):
            embed = re.search(r'r:embed="([^"]+)"', block)
            if embed is None:
                continue
            digest = media_digest.get(rid_media.get(embed.group(1), ""))
            geom = by_digest.get(digest or "")
            if geom is None:
                continue
            img_w, img_h = geom["size"]
            origin = re.search(
                r"<from><col>(\d+)</col><colOff>(\d+)</colOff><row>(\d+)</row><rowOff>(\d+)</rowOff></from><ext cx=\"(\d+)\" cy=\"(\d+)\"",
                block,
            )
            if origin is None or img_w <= 0 or img_h <= 0:
                continue
            col, col_off, row, row_off, pic_cx, pic_cy = (int(v) for v in origin.groups())
            for circle in geom["circles"]:
                extras.append(_circle_anchor(circle, next_shape, col, col_off, row, row_off, pic_cx, pic_cy, img_w, img_h))
                next_shape += 1
                placed += 1
        if extras:
            drawing = drawing.replace("</wsDr>", "".join(extras) + "</wsDr>")
            parts[drawing_name] = drawing.encode("utf-8")
    if placed:
        tmp = path.with_suffix(".tmp.xlsx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in names:
                zf.writestr(name, parts[name])
        tmp.replace(path)
    spread_circle_anchors(path)
    return placed


def _circle_anchor(circle, shape_id, col, col_off, row, row_off, pic_cx, pic_cy, img_w, img_h) -> str:
    """One movable oval, positioned as a fraction of the picture it sits on."""
    left = max(0.0, min(float(circle["x"]), img_w - 1))
    top = max(0.0, min(float(circle["y"]), img_h - 1))
    width = max(1.0, min(float(circle["w"]), img_w - left))
    height = max(1.0, min(float(circle["h"]), img_h - top))
    dash = "sysDot" if circle.get("dotted") else "dash"
    cx = max(1, int(width / img_w * pic_cx))
    cy = max(1, int(height / img_h * pic_cy))
    return (
        "<oneCellAnchor>"
        f"<from><col>{col}</col><colOff>{col_off + int(left / img_w * pic_cx)}</colOff>"
        f"<row>{row}</row><rowOff>{row_off + int(top / img_h * pic_cy)}</rowOff></from>"
        f"<ext cx=\"{cx}\" cy=\"{cy}\"/>"
        "<sp><nvSpPr>"
        f"<cNvPr id=\"{shape_id}\" name=\"{circle['name']}\"/><cNvSpPr/></nvSpPr>"
        "<spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"
        f"<a:ext cx=\"{cx}\" cy=\"{cy}\"/>"
        "</a:xfrm><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom><a:noFill/>"
        "<a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"FF6600\"/></a:solidFill>"
        f"<a:prstDash val=\"{dash}\"/></a:ln></spPr>"
        "<txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></txBody>"
        "</sp><clientData/></oneCellAnchor>"
    )


def _column_width_px(width: float) -> int:
    """Excel's display width for a stored column width, in pixels at 96 dpi."""
    if width < 1:
        return int(width * 12 + 0.5)
    return int(width * 7 + 0.5) + 5


def _sheet_metric(sheet_xml: bytes):
    """Column widths and row heights for one worksheet, in EMUs."""
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    root = ET.fromstring(sheet_xml)
    fmt = root.find(f"{ns}sheetFormatPr")
    default_pt = float(fmt.attrib.get("defaultRowHeight", "15")) if fmt is not None else 15.0
    default_width = 8.43
    if fmt is not None and "defaultColWidth" in fmt.attrib:
        default_width = float(fmt.attrib["defaultColWidth"])
    col_px: dict[int, int] = {}
    cols = root.find(f"{ns}cols")
    if cols is not None:
        for col in cols:
            px = _column_width_px(float(col.attrib.get("width", default_width)))
            for index in range(int(col.attrib["min"]), int(col.attrib["max"]) + 1):
                col_px[index] = px
    row_pt: dict[int, float] = {}
    for row in root.findall(f"{ns}sheetData/{ns}row"):
        if "ht" in row.attrib:
            row_pt[int(row.attrib["r"])] = float(row.attrib["ht"])
    return col_px, default_width, row_pt, default_pt


def _spread_drawing_anchors(drawing: str, sheet_xml: bytes) -> str:
    """Move each circle's anchor into the cell under its bad spot.

    A from-offset larger than the cell is ignored by Excel, which stacks every
    circle on the picture's first cell. The offset is split across the columns
    and rows the picture covers, and each circle keeps its own cell.
    """
    col_px, default_width, row_pt, default_pt = _sheet_metric(sheet_xml)
    px_emu = 9525
    pt_emu = 12700

    def col_emu(index0: int) -> int:
        return col_px.get(index0 + 1, _column_width_px(default_width)) * px_emu

    def row_emu(index0: int) -> int:
        return int(round(row_pt.get(index0 + 1, default_pt) * pt_emu))

    def prefix(count: int, size_of) -> list[int]:
        total = [0]
        for i in range(count):
            total.append(total[-1] + size_of(i))
        return total

    col_at = prefix(50, col_emu)
    row_at = prefix(2500, row_emu)

    def locate(at: list[int], absolute: int) -> tuple[int, int]:
        absolute = max(0, absolute)
        lo, hi = 0, len(at) - 2
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if at[mid] <= absolute:
                lo = mid
            else:
                hi = mid - 1
        offset = absolute - at[lo]
        room = at[lo + 1] - at[lo]
        if offset >= room:
            offset = max(0, room - 1)
        return lo, int(offset)

    def place(match: re.Match) -> str:
        block = match.group(0)
        if 'prst="ellipse"' not in block:
            return block
        found = re.search(
            r"<from><col>(\d+)</col><colOff>(\d+)</colOff><row>(\d+)</row><rowOff>(\d+)</rowOff></from>",
            block,
        )
        if found is None:
            return block
        col, col_off, row, row_off = (int(v) for v in found.groups())
        if col >= len(col_at) - 1 or row >= len(row_at) - 1:
            return block
        new_col, new_col_off = locate(col_at, col_at[col] + col_off)
        new_row, new_row_off = locate(row_at, row_at[row] + row_off)
        updated = (
            f"<from><col>{new_col}</col><colOff>{new_col_off}</colOff>"
            f"<row>{new_row}</row><rowOff>{new_row_off}</rowOff></from>"
        )
        return block.replace(found.group(0), updated, 1)

    return re.sub(r"<oneCellAnchor>.*?</oneCellAnchor>", place, drawing, flags=re.S)


def _worksheet_xml(parts: dict[str, bytes]) -> dict[str, bytes]:
    """Map a worksheet name to its worksheet XML."""
    main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    wb = ET.fromstring(parts["xl/workbook.xml"])
    book_rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rid_target = {rel.attrib["Id"]: rel.attrib["Target"].lstrip("/") for rel in book_rels}
    found = {}
    for node in wb.find(f"{main}sheets"):
        found[node.attrib["name"]] = parts[rid_target[node.attrib[f"{rel_ns}id"]]]
    return found


def _emu_prefix(sheet_xml: bytes):
    col_px, default_width, row_pt, default_pt = _sheet_metric(sheet_xml)

    def col_emu(index0: int) -> int:
        return col_px.get(index0 + 1, _column_width_px(default_width)) * 9525

    def row_emu(index0: int) -> int:
        return int(round(row_pt.get(index0 + 1, default_pt) * 12700))

    def prefix(count: int, size_of) -> list[int]:
        total = [0]
        for i in range(count):
            total.append(total[-1] + size_of(i))
        return total

    return prefix(60, col_emu), prefix(2500, row_emu)


def _drop_ellipses_on_picture(drawing: str, sheet_xml: bytes, picture_block: str, keep_last: int = 0) -> str:
    """Remove circle shapes that sit on this picture, optionally keeping the newest ones."""
    origin = re.search(
        r"<from><col>(\d+)</col><colOff>(\d+)</colOff><row>(\d+)</row><rowOff>(\d+)</rowOff></from><ext cx=\"(\d+)\" cy=\"(\d+)\"",
        picture_block,
    )
    if origin is None:
        return drawing
    col_at, row_at = _emu_prefix(sheet_xml)
    pcol, pcoff, prow, proff, pcx, pcy = (int(v) for v in origin.groups())
    if pcol >= len(col_at) - 1 or prow >= len(row_at) - 1:
        return drawing
    left, top = col_at[pcol] + pcoff, row_at[prow] + proff
    right, bottom = left + pcx, top + pcy
    blocks = list(re.finditer(r"<oneCellAnchor>.*?</oneCellAnchor>", drawing, flags=re.S))
    on_picture = []
    for match in blocks:
        block = match.group(0)
        if 'prst="ellipse"' not in block:
            continue
        found = re.search(
            r"<from><col>(\d+)</col><colOff>(\d+)</colOff><row>(\d+)</row><rowOff>(\d+)</rowOff></from><ext cx=\"(\d+)\" cy=\"(\d+)\"",
            block,
        )
        if found is None:
            continue
        col, coff, row, roff, cx, cy = (int(v) for v in found.groups())
        if col >= len(col_at) - 1 or row >= len(row_at) - 1:
            continue
        cx0 = col_at[col] + coff + cx / 2.0
        cy0 = row_at[row] + roff + cy / 2.0
        if left - 20000 <= cx0 <= right + 20000 and top - 20000 <= cy0 <= bottom + 20000:
            on_picture.append(match)
    drop = set(id(m) for m in on_picture[: max(0, len(on_picture) - keep_last)])
    if not drop:
        return drawing
    pieces = []
    cursor = 0
    for match in blocks:
        if id(match) in drop:
            pieces.append(drawing[cursor:match.start()])
            cursor = match.end()
    pieces.append(drawing[cursor:])
    return "".join(pieces)


def spread_circle_anchors(path: Path) -> int:
    """Re-anchor every bad-spot circle so Excel draws it on its own bad spot."""
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    moved = 0
    drawings = _sheet_drawings(parts)
    main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    wb = ET.fromstring(parts["xl/workbook.xml"])
    book_rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rid_target = {rel.attrib["Id"]: rel.attrib["Target"].lstrip("/") for rel in book_rels}
    sheet_part = {}
    for node in wb.find(f"{main}sheets"):
        sheet_part[node.attrib["name"]] = rid_target[node.attrib[f"{rel_ns}id"]]
    changed = False
    for sheet, drawing_name in drawings.items():
        if not sheet.startswith("Bad Spot"):
            continue
        original = parts[drawing_name].decode("utf-8")
        updated = _spread_drawing_anchors(original, parts[sheet_part[sheet]])
        if updated != original:
            parts[drawing_name] = updated.encode("utf-8")
            moved += original.count('prst="ellipse"')
            changed = True
    if changed:
        tmp = path.with_suffix(".tmp.xlsx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in names:
                zf.writestr(name, parts[name])
        tmp.replace(path)
    return moved


_CIRCLE_NOTE_OLD = (
    "Each circle is a separate orange shape on the satellite map, named "
    "Circle 1, Circle 2, and so on, in the same order as the table, so a circle can be dragged "
    "without moving the others."
)
_CIRCLE_NOTE_NEW = (
    "Each circle is a separate orange shape on the combined map, named "
    "Circle 1, Circle 2, and so on, in the same order as the table. Each zoom panel has that "
    "same circle as its own shape, not painted into the picture, so you can drag it. The RSRP, "
    "RSRQ and SINR sheets do not repeat these circles."
)
_PIE_NOTE_OLD = "The RSRP, RSRQ and SINR sheets do not repeat these circles."
_PIE_NOTE_NEW = (
    "The RSRP, RSRQ and SINR sheets do not repeat these circles. "
    "Each zoom panel draws a sector pie for every nearby site from System DUMP.xlsb. "
    "The label is the sector name. Where a sector shoots toward that bad spot, a yellow line shows the distance."
)


def refresh_zoom_circles(src: Path | None = None) -> Path:
    """Update an existing report without redrawing the sheets that stay the same.

    Removes the circle shapes from the RSRP, RSRQ and SINR sheets, redraws only
    the zoom pictures without a painted circle, and adds one movable circle on
    each zoom panel.
    """
    import io
    from PIL import Image

    from merge_and_analyze import ACTIVE_MERGED, MERGED_CSV, load_active_merged, load_merged_csv

    src = src or REPORT_XLSX
    print("Loading drive logs", flush=True)
    idle = load_merged_csv() if MERGED_CSV.exists() else None
    active = load_active_merged() if ACTIVE_MERGED.exists() else None
    if idle is None or active is None:
        raise SystemExit("Merged drive logs are missing; a full report build is required.")
    active = select_active_report(active)
    print(f"Idle rows {len(idle):,}  Active {ACTIVE_OPERATOR} {ACTIVE_CELL_TYPE} rows {len(active):,}", flush=True)
    idle_track = best_server(idle)
    active_track = best_carrier(active)
    MAP_CIRCLE_PIXELS.clear()
    zoom_for_sheet = {}
    for tag, frame in (("", idle_track), ("active", active_track)):
        suffix = f"_{tag}" if tag else ""
        zoom_bit = "Active" if tag else "first all-band map"
        for kpi, classify, labels, colors in (
            ("RSRP", rsrp_map_class, MAP_RSRP_LABELS, MAP_RSRP_COLORS),
            ("RSRQ", rsrq_map_class, MAP_RSRQ_LABELS, MAP_RSRQ_COLORS),
            ("SINR", sinr_map_class, MAP_SINR_LABELS, MAP_SINR_COLORS),
        ):
            print(f"Drawing {kpi} zooms ({tag or 'idle'})", flush=True)
            spots = find_bad_spots(frame, kpi)
            path = PLOTS_DIR / f"excel_bad_spots_{kpi.lower()}_zooms{suffix}.png"
            plot_bad_spot_zoom_grid(
                frame, spots, path, kpi, classify(frame[kpi]), labels, colors,
                f"{kpi} zoomed locations ({zoom_bit})", shapes=True,
            )
            sheet = f"Bad Spot {kpi}" + ("_Active" if tag else "")
            zoom_for_sheet[sheet] = path
            recorded = len(MAP_CIRCLE_PIXELS[str(path.resolve())]["circles"])
            print(f"  {sheet}: {recorded} movable circles", flush=True)

    print("Updating the workbook", flush=True)
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    drawings = _sheet_drawings(parts)
    sheet_xml = _worksheet_xml(parts)
    for sheet, drawing_name in drawings.items():
        if sheet.startswith("Bad Spot"):
            continue
        drawing = parts[drawing_name].decode("utf-8")
        stripped = re.sub(
            r'<oneCellAnchor>(?:(?!</oneCellAnchor>).)*prst="ellipse".*?</oneCellAnchor>',
            "",
            drawing,
            flags=re.S,
        )
        parts[drawing_name] = stripped.encode("utf-8")

    for sheet, png in zoom_for_sheet.items():
        drawing_name = drawings[sheet]
        rels_name = drawing_name.replace("drawings/", "drawings/_rels/") + ".rels"
        rels = parts[rels_name].decode("utf-8")
        drawing = parts[drawing_name].decode("utf-8")
        rid_media = {}
        for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
            rid = re.search(r'Id="([^"]+)"', rel).group(1)
            target = re.search(r'Target="([^"]+)"', rel).group(1)
            media = target.lstrip("/") if target.startswith("/") else "xl/media/" + target.rsplit("/", 1)[1]
            rid_media[rid] = media
        pictures = []
        for block in re.findall(r"<oneCellAnchor>.*?</oneCellAnchor>", drawing, flags=re.S):
            if "<pic>" not in block:
                continue
            row = int(re.search(r"<row>(\d+)</row>", block).group(1))
            pictures.append((row, block))
        _row, zoom_block = max(pictures)
        embed = re.search(r'r:embed="([^"]+)"', zoom_block).group(1)
        media = rid_media[embed]
        image = Image.open(png).convert("RGB")
        buf = io.BytesIO()
        image.save(buf, "JPEG", quality=74, optimize=True)
        parts[media] = buf.getvalue()
        geom = MAP_CIRCLE_PIXELS[str(Path(png).resolve())]
        img_w, img_h = geom["size"]
        origin = re.search(
            r"<from><col>(\d+)</col><colOff>(\d+)</colOff><row>(\d+)</row><rowOff>(\d+)</rowOff></from><ext cx=\"(\d+)\" cy=\"(\d+)\"",
            zoom_block,
        )
        col, col_off, row, row_off, pic_cx, pic_cy = (int(v) for v in origin.groups())
        drawing = _drop_ellipses_on_picture(drawing, sheet_xml[sheet], zoom_block, keep_last=0)
        next_shape = max(int(i) for i in re.findall(r'<cNvPr id="(\d+)"', drawing)) + 1
        extras = []
        for circle in geom["circles"]:
            extras.append(_circle_anchor(circle, next_shape, col, col_off, row, row_off, pic_cx, pic_cy, img_w, img_h))
            next_shape += 1
        parts[drawing_name] = drawing.replace("</wsDr>", "".join(extras) + "</wsDr>").encode("utf-8")
        print(f"  {sheet}: replaced zoom picture, added {len(extras)} circles", flush=True)

    note_hits = 0
    for name, body in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        text = body.decode("utf-8")
        if _CIRCLE_NOTE_OLD not in text:
            continue
        parts[name] = text.replace(_CIRCLE_NOTE_OLD, _CIRCLE_NOTE_NEW).encode("utf-8")
        note_hits += text.count(_CIRCLE_NOTE_OLD)
    pie_hits = 0
    for name, body in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        text = body.decode("utf-8")
        changed = False
        if _PIE_NOTE_OLD in text and "sector pie" not in text:
            text = text.replace(_PIE_NOTE_OLD, _PIE_NOTE_NEW)
            pie_hits += 1
            changed = True
        if "Physical_Site_Database_V2.0.xlsx" in text or "Physical_Site_Database_V1.0.xlsb" in text:
            text = text.replace("Physical_Site_Database_V2.0.xlsx", "System DUMP.xlsb")
            text = text.replace("Physical_Site_Database_V1.0.xlsb", "System DUMP.xlsb")
            changed = True
        if changed:
            parts[name] = text.encode("utf-8")
    print(f"Updated circle note in {note_hits} place(s) and sector-pie note in {pie_hits}", flush=True)

    order = [n for n in names if n in parts]
    tmp = src.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            zf.writestr(name, parts[name])
    tmp.replace(VERSIONED_XLSX)
    if VERSIONED_XLSX.resolve() != REPORT_XLSX.resolve():
        shutil.copy2(VERSIONED_XLSX, REPORT_XLSX)
    spread_circle_anchors(VERSIONED_XLSX)
    if VERSIONED_XLSX.resolve() != REPORT_XLSX.resolve():
        shutil.copy2(VERSIONED_XLSX, REPORT_XLSX)
    print(f"Wrote {VERSIONED_XLSX}", flush=True)
    return VERSIONED_XLSX


def compress_workbook_images(path: Path, quality: int = 74) -> int:
    """Store the satellite maps as JPEG so the workbook stays a normal download.

    Picture size and the circle shapes are unchanged; only the photo bytes shrink.
    """
    from PIL import Image
    import io

    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    renamed = {}
    for name in list(parts):
        if not (name.startswith("xl/media/") and name.lower().endswith(".png")):
            continue
        if len(parts[name]) < 400_000:
            continue
        image = Image.open(io.BytesIO(parts[name])).convert("RGB")
        buf = io.BytesIO()
        image.save(buf, "JPEG", quality=quality, optimize=True)
        jpg = name[:-4] + ".jpg"
        parts[jpg] = buf.getvalue()
        del parts[name]
        renamed[name] = jpg
    if not renamed:
        return 0
    for name, body in list(parts.items()):
        if not name.endswith(".rels"):
            continue
        text = body.decode("utf-8")
        for old, new in renamed.items():
            text = text.replace(old, new)
            text = text.replace(old.rsplit("/", 1)[1], new.rsplit("/", 1)[1])
        parts[name] = text.encode("utf-8")
    types = parts["[Content_Types].xml"].decode("utf-8")
    if 'Extension="jpg"' not in types:
        types = types.replace(
            "<Default ",
            '<Default Extension="jpg" ContentType="image/jpeg" /><Default ',
            1,
        )
        parts["[Content_Types].xml"] = types.encode("utf-8")
    order = [n for n in names if n in parts] + [n for n in parts if n not in names]
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            zf.writestr(name, parts[name])
    tmp.replace(path)
    return len(renamed)


def banner(ws: Worksheet, title: str, subtitle: str, last_col: int = 12) -> None:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    write_cell(ws, 1, 1, title, size=20, bold=True, color=WHITE, fill_color=NAVY, align="left")
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    write_cell(ws, 2, 1, subtitle, size=12, color=WHITE, fill_color=NAVY)
    ws.row_dimensions[2].height = 20
    for col in range(1, last_col + 1):
        ws.cell(1, col).fill = fill(NAVY)
        ws.cell(2, col).fill = fill(NAVY)
        ws.cell(1, col).font = font(20, bold=True, color=WHITE)
        ws.cell(2, col).font = font(12, color=WHITE)


def section_bar(ws: Worksheet, row: int, text: str, start_col: int, end_col: int, fill_color=YELLOW) -> None:
    """Yellow section header like IDLE Mode / Active Mode."""
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    write_cell(ws, row, start_col, text, size=13, bold=True, color=NAVY, fill_color=fill_color, align="center")
    ws.row_dimensions[row].height = 22
    for col in range(start_col, end_col + 1):
        ws.cell(row, col).fill = fill(fill_color)
        ws.cell(row, col).font = font(13, bold=True, color=NAVY)
        ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")


def save_active_placeholder(path: Path) -> Path:
    """Blue placeholder shown until Active Mode logs are provided."""
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(7.4, 8.2))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.add_patch(plt.Rectangle((0.12, 0.38), 0.76, 0.24, facecolor="#2E86C1", edgecolor="none", linewidth=0))
    ax.text(
        0.5,
        0.50,
        "Later I will give you data,\nthen plot here",
        ha="center",
        va="center",
        color="white",
        fontsize=14,
        fontweight="bold",
    )
    fig.savefig(path, dpi=140, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return path


def write_table(ws, start_row: int, start_col: int, headers: list[str], rows: list[list], header_fill=NAVY, num_formats: dict[int, str] | None = None):
    for i, h in enumerate(headers):
        cell = write_cell(ws, start_row, start_col + i, h, size=11, bold=True, color=WHITE, fill_color=header_fill, align="center")
        cell.border = THIN
    for r_i, row in enumerate(rows, start=start_row + 1):
        bg = LIGHT if (r_i - start_row) % 2 == 0 else WHITE
        for c_i, value in enumerate(row):
            fmt = (num_formats or {}).get(c_i)
            cell = write_cell(ws, r_i, start_col + c_i, value, fill_color=bg, align="center" if c_i else "left", num_fmt=fmt)
            cell.border = THIN
    return start_row + len(rows)


def _title_outside(title_obj) -> None:
    if title_obj is None:
        return
    try:
        title_obj.overlay = False
    except Exception:
        pass


def reserve_plot_area(chart, left=0.14, top=0.16, width=0.78, height=0.60) -> None:
    """Keep plot area inside the frame so axis titles and tick levels stay visible."""
    chart.layout = Layout(
        manualLayout=ManualLayout(
            xMode="edge",
            yMode="edge",
            x=left,
            y=top,
            w=width,
            h=height,
        )
    )


def apply_xy_levels(
    chart,
    x_title: str,
    y_title: str,
    y_min=None,
    y_max=None,
    show_all_x_ticks=True,
    x_tick_skip=None,
    rotate_x=False,
    y_major_unit=None,
    y_ax_pos="l",
    layout=True,
    y_num_fmt=None,
) -> None:
    """Put category levels on the bottom (X) and value levels on the left (Y).

    openpyxl defaults both axes to axPos='l', which Excel then draws without
    horizontal tick labels or a usable vertical scale — the bug in the v1.3 charts.
    """
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    _title_outside(getattr(chart, "title", None))
    _title_outside(chart.x_axis.title)
    _title_outside(chart.y_axis.title)
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = y_ax_pos
    chart.x_axis.tickLblPos = "nextTo"
    chart.y_axis.tickLblPos = "nextTo"
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"
    chart.x_axis.minorTickMark = "none"
    chart.y_axis.minorTickMark = "none"
    if y_num_fmt:
        chart.y_axis.numFmt = y_num_fmt
    if x_tick_skip is not None:
        try:
            chart.x_axis.tickLblSkip = int(x_tick_skip)
            chart.x_axis.tickMarkSkip = int(x_tick_skip)
        except Exception:
            pass
    elif show_all_x_ticks:
        try:
            chart.x_axis.tickLblSkip = 1
            chart.x_axis.tickMarkSkip = 1
        except Exception:
            pass
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    if y_max is not None:
        chart.y_axis.scaling.max = y_max
    if y_major_unit is not None:
        try:
            chart.y_axis.majorUnit = y_major_unit
        except Exception:
            pass
    if rotate_x:
        # Rotation only — do not add an empty text run, which blanks tick labels.
        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(rot="-5400000", wrap="square", anchor="ctr"),
            p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=700)))],
        )
    if layout:
        reserve_plot_area(chart)


def style_chart(chart, color: str, title: str, y_title: str, x_title: str, width=14, height=8, show_legend=False) -> None:
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = x_title
    chart.style = 10
    chart.width = width
    chart.height = height
    if show_legend:
        chart.legend.position = "b"
    else:
        chart.legend = None
    if chart.series:
        chart.series[0].graphicalProperties.solidFill = color
        chart.series[0].graphicalProperties.line.solidFill = color
    apply_xy_levels(chart, x_title, y_title, show_all_x_ticks=False, rotate_x=False)


def add_col_chart(ws_data, ws_dest, anchor, title, color, cat_col, data_col, min_row, max_row, y_title, x_title, width=14, height=8, show_legend=False):
    chart = BarChart()
    chart.type = "col"
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_col, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    style_chart(chart, color, title, y_title, x_title, width=width, height=height, show_legend=show_legend)
    chart.shape = 4
    ws_dest.add_chart(chart, anchor)
    return chart


def add_stacked_col_chart(ws_data, ws_dest, anchor, title, cat_col, data_min, data_max, min_row, max_row, y_title, x_title, colors, width=14, height=8):
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_min, max_col=data_max, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = x_title
    chart.style = 10
    chart.width = width
    chart.height = height
    chart.legend.position = "b"
    apply_xy_levels(chart, x_title, y_title, show_all_x_ticks=False, rotate_x=False)
    for i, color in enumerate(colors):
        if i < len(chart.series):
            hexcol = color.lstrip("#")
            chart.series[i].graphicalProperties.solidFill = hexcol
            chart.series[i].graphicalProperties.line.solidFill = hexcol
    ws_dest.add_chart(chart, anchor)
    return chart


def add_line_chart(ws_data, ws_dest, anchor, title, color, cat_col, data_col, min_row, max_row, y_title, x_title, width=14, height=8, legend=False):
    chart = LineChart()
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_col, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    style_chart(chart, color, title, y_title, x_title, width=width, height=height)
    if legend:
        chart.legend.position = "b"
    chart.series[0].graphicalProperties.line.solidFill = color
    chart.series[0].graphicalProperties.line.width = 18000
    ws_dest.add_chart(chart, anchor)
    return chart


def write_stacked_hist(ws, start_col, values, fine_bins, class_series, class_labels):
    values = pd.to_numeric(values, errors="coerce")
    mask = values.notna()
    values = values[mask]
    classes = class_series.reindex(values.index).astype(str)
    ws.cell(1, start_col, "bin")
    for j, lab in enumerate(class_labels):
        ws.cell(1, start_col + 1 + j, str(lab))
    lefts = fine_bins[:-1]
    for i, left in enumerate(lefts, start=2):
        ws.cell(i, start_col, float(left))
    for j, lab in enumerate(class_labels):
        subset = values[classes == str(lab)]
        counts, _ = np.histogram(subset.to_numpy(), bins=fine_bins)
        for i, n in enumerate(counts, start=2):
            ws.cell(i, start_col + 1 + j, int(n))
    return {
        "cat_col": start_col,
        "data_min": start_col + 1,
        "data_max": start_col + len(class_labels),
        "n": len(lefts),
    }


def write_chart_data(ws: Worksheet, df: pd.DataFrame) -> dict:
    """Hidden sheet holding all series used by Excel charts."""
    rsrp_bins = np.arange(-140, -48, 2)
    rsrq_bins = np.arange(-24, -2.5, 0.5)
    sinr_bins = np.arange(-15, 31, 1)

    rsrp_cls = rsrp_map_class(df["RSRP"])
    rsrq_cls = rsrq_map_class(df["RSRQ"])
    sinr_cls = sinr_map_class(df["SINR"])

    blocks = {
        "RSRP": write_stacked_hist(ws, 26, df["RSRP"], rsrp_bins, rsrp_cls, MAP_RSRP_LABELS),
        "RSRQ": write_stacked_hist(ws, 35, df["RSRQ"], rsrq_bins, rsrq_cls, MAP_RSRQ_LABELS),
        "SINR": write_stacked_hist(ws, 41, df["SINR"], sinr_bins, sinr_cls, MAP_SINR_LABELS),
    }

    percentiles = list(range(0, 101, 2))
    ws.cell(1, 7, "Percentile")
    ws.cell(1, 8, "RSRP_CDF")
    ws.cell(1, 9, "RSRQ_CDF")
    ws.cell(1, 10, "SINR_CDF")
    rsrp = df["RSRP"].dropna()
    rsrq = df["RSRQ"].dropna()
    sinr = df["SINR"].dropna()
    rsrp_p = np.nanpercentile(rsrp, percentiles)
    rsrq_p = np.nanpercentile(rsrq, percentiles)
    sinr_p = np.nanpercentile(sinr, percentiles)
    for i, p in enumerate(percentiles, start=2):
        ws.cell(i, 7, p)
        ws.cell(i, 8, float(rsrp_p[i - 2]))
        ws.cell(i, 9, float(rsrq_p[i - 2]))
        ws.cell(i, 10, float(sinr_p[i - 2]))
    blocks["cdf_n"] = len(percentiles)

    ws.cell(1, 12, "RSRP_quality")
    ws.cell(1, 13, "RSRP_quality_n")
    rsrp_q = rsrp_range_counts(df["RSRP"])
    for i, (label, n) in enumerate(rsrp_q.items(), start=2):
        ws.cell(i, 12, str(label))
        ws.cell(i, 13, int(n))
    blocks["rsrp_q_n"] = len(rsrp_q)
    ws.cell(1, 15, "RSRQ_quality")
    ws.cell(1, 16, "RSRQ_quality_n")
    for i, (label, n) in enumerate(map_range_counts(df["RSRQ"], rsrq_map_class, MAP_RSRQ_LABELS).items(), start=2):
        ws.cell(i, 15, str(label))
        ws.cell(i, 16, int(n))
    ws.cell(1, 18, "SINR_quality")
    ws.cell(1, 19, "SINR_quality_n")
    for i, (label, n) in enumerate(map_range_counts(df["SINR"], sinr_map_class, MAP_SINR_LABELS).items(), start=2):
        ws.cell(i, 18, str(label))
        ws.cell(i, 19, int(n))

    ts = (
        df.dropna(subset=["Time"])
        .set_index("Time")[["RSRP", "RSRQ", "SINR"]]
        .resample("10min")
        .mean()
        .dropna(how="all")
    )
    if len(ts) > 900:
        ts = ts.iloc[:: max(1, len(ts) // 900)]
    ws.cell(1, 21, "Time")
    ws.cell(1, 22, "RSRP")
    ws.cell(1, 23, "RSRQ")
    ws.cell(1, 24, "SINR")
    for i, (t, row) in enumerate(ts.iterrows(), start=2):
        cell = ws.cell(i, 21, pd.Timestamp(t).to_pydatetime())
        cell.number_format = "yyyy-mm-dd hh:mm"
        ws.cell(i, 22, None if pd.isna(row["RSRP"]) else float(row["RSRP"]))
        ws.cell(i, 23, None if pd.isna(row["RSRQ"]) else float(row["RSRQ"]))
        ws.cell(i, 24, None if pd.isna(row["SINR"]) else float(row["SINR"]))
    blocks["ts_n"] = len(ts)

    rsrp_vs_bins = list(range(-80, -126, -1))
    pair = df.copy()
    pair["rsrp_bin"] = pd.to_numeric(pair["RSRP"], errors="coerce").round().astype("Int64")
    ws.cell(1, 50, "RSRP_dBm")
    ws.cell(1, 51, "SINR")
    ws.cell(1, 52, "RSRQ")
    all_mean = pair.groupby("rsrp_bin")[["SINR", "RSRQ"]].mean()
    for i, b in enumerate(rsrp_vs_bins, start=2):
        ws.cell(i, 50, int(b))
        if b in all_mean.index:
            ws.cell(i, 51, None if pd.isna(all_mean.loc[b, "SINR"]) else float(all_mean.loc[b, "SINR"]))
            ws.cell(i, 52, None if pd.isna(all_mean.loc[b, "RSRQ"]) else float(all_mean.loc[b, "RSRQ"]))
    blocks["vs_n"] = len(rsrp_vs_bins)
    return blocks


def write_vs_block(ws: Worksheet, df: pd.DataFrame, cat_col: int, sinr_col: int, rsrq_col: int) -> int:
    """Mean SINR and RSRQ for each 1 dB RSRP bin, used by the RSRP vs KPIs charts."""
    rsrp_vs_bins = list(range(-80, -126, -1))
    pair = df.copy()
    pair["rsrp_bin"] = pd.to_numeric(pair["RSRP"], errors="coerce").round().astype("Int64")
    ws.cell(1, cat_col, "RSRP_dBm")
    ws.cell(1, sinr_col, "SINR")
    ws.cell(1, rsrq_col, "RSRQ")
    all_mean = pair.groupby("rsrp_bin")[["SINR", "RSRQ"]].mean()
    for i, b in enumerate(rsrp_vs_bins, start=2):
        ws.cell(i, cat_col, int(b))
        if b in all_mean.index:
            ws.cell(i, sinr_col, None if pd.isna(all_mean.loc[b, "SINR"]) else float(all_mean.loc[b, "SINR"]))
            ws.cell(i, rsrq_col, None if pd.isna(all_mean.loc[b, "RSRQ"]) else float(all_mean.loc[b, "RSRQ"]))
    return len(rsrp_vs_bins)


def _campaign_label(series: pd.Series) -> str:
    """Collapse SYS suffixes so AUTO VOICE_SYS01 and AUTO VOICE_SYS02 read as AUTO VOICE."""
    values = series.dropna().astype(str).str.strip()
    values = values[values.ne("") & values.str.lower().ne("nan")]
    if values.empty:
        return ""
    upper = values.str.upper()
    if upper.str.contains("AUTO VOICE").all():
        return "AUTO VOICE"
    if upper.str.contains("AUTO DATA").all():
        return "AUTO DATA"
    return str(values.mode().iloc[0])


def build_cover(
    ws: Worksheet,
    df: pd.DataFrame,
    active: pd.DataFrame | None = None,
    active_raw_n: int = 0,
    active_operator_n: int = 0,
) -> None:
    last_col = 20
    banner(
        ws,
        "  Bogura Drive-Test Report",
        f"  Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  IDLE + Active Mode  |  Gridlines off  |  Freeze panes off",
        last_col=last_col,
    )
    set_widths(ws, {get_column_letter(i): 14 for i in range(1, last_col + 1)})
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["E"].width = 20
    section_bar(ws, 3, "IDLE Mode", 1, last_col)

    write_cell(ws, 5, 1, "Dataset", size=14, bold=True)
    info_rows = [
        ["Mode", "IDLE (scanner)"],
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Archive", "Borga_Idle_V2.0.7z"],
        ["CSV files", int(df["Source"].nunique())],
        ["Merged samples", int(len(df))],
        ["Start time", df["Time"].min()],
        ["End time", df["Time"].max()],
        ["Unique Cell Ids", int(df["Cell Id"].nunique())],
        ["Longitude range", f"{df['Longitude'].min():.5f} – {df['Longitude'].max():.5f}"],
        ["Latitude range", f"{df['Latitude'].min():.5f} – {df['Latitude'].max():.5f}"],
        ["Site database", SITE_DB_PATH.name if SITE_DB_PATH.exists() else "not provided"],
    ]
    write_table(ws, 6, 1, ["Item", "Value"], info_rows, num_formats={1: "#,##0"})
    ws.cell(12, 2).number_format = "yyyy-mm-dd hh:mm:ss"
    ws.cell(13, 2).number_format = "yyyy-mm-dd hh:mm:ss"

    write_cell(ws, 5, 4, "Radio KPIs (all samples)", size=14, bold=True)
    stats_headers = ["KPI", "Count", "Min", "P5", "Median", "Mean", "P95", "Max"]
    stats_rows = []
    for name in ("RSRP", "RSRQ", "SINR"):
        st = kpi_stats(df[name])
        stats_rows.append(
            [name, st["Count"], st["Min"], st["P5"], st["Median"], st["Mean"], st["P95"], st["Max"]]
        )
    write_table(
        ws,
        6,
        4,
        stats_headers,
        stats_rows,
        num_formats={1: "#,##0", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00", 6: "0.00", 7: "0.00"},
    )

    write_cell(ws, 18, 1, "How to read this workbook", size=14, bold=True)
    notes = [
        "Cover \u2014 Idle dataset and plots on top. Active Mode dataset, histograms and the DT route are below.",
        "RSRP / RSRQ / SINR \u2014 Idle scanner on the left (combined all-band map with bad spots, then L900 / L1800 / L2100 / L2600). Active Mode Robi UE on the right, same layers. No site pies.",
        "Bad Spot RSRP / RSRQ / SINR are Idle. Bad Spot RSRP_Active / RSRQ_Active / SINR_Active apply the same poor-run rules to the Active Robi track. Each bad spot is its own large map, in order. The combined view has no circles.",
        "Poor is RSRP below -120 dBm (red only; -115 through -120 is light red and is not poor), RSRQ at or below -20 dB (inside red -30 to -17), SINR at or below 0 dB (orange/red); every poor run is circled, however short.",
        "RSRP vs KPIs is Idle. RSRP vs KPIs_Active is the same chart set for the Active Robi UE.",
        "Sample Log \u2014 evenly spaced subset of the Idle samples (full logs: output/Bogura_merged.csv.gz and output/Bogura_active_merged.csv.gz).",
    ]
    for i, text in enumerate(notes):
        ws.merge_cells(start_row=19 + i, start_column=1, end_row=19 + i, end_column=11)
        write_cell(ws, 19 + i, 1, text, wrap=True)
        ws.row_dimensions[19 + i].height = 18

    write_cell(ws, 25, 1, "Band mix (DL EARFCN)", size=14, bold=True)
    band = df["Band"].value_counts()
    band_rows = [[str(k), int(v), v / len(df)] for k, v in band.items()]
    write_table(ws, 26, 1, ["Band", "Samples", "Share"], band_rows, num_formats={1: "#,##0", 2: "0.0%"})

    write_cell(ws, 25, 5, "RSRP ranges (dBm)", size=14, bold=True)
    rsrp_q = rsrp_range_counts(df["RSRP"])
    rsrp_rows = [[str(label), int(n), n / len(df)] for label, n in rsrp_q.items()]
    write_table(
        ws,
        26,
        5,
        ["Range", "Samples", "Share"],
        rsrp_rows,
        num_formats={1: "#,##0", 2: "0.0%"},
    )
    write_cell(
        ws,
        36,
        5,
        "Ranges match the coverage-map legend: -95 to 0 (dark green), -110 to -95 (green), -115 to -110 (yellow), -120 to -115 (light red, from -115), -150 to -120 (red).",
        size=9,
        wrap=True,
    )
    ws.merge_cells(start_row=36, start_column=5, end_row=36, end_column=8)
    section_bar(ws, 42, "Active Mode", 1, last_col)
    if active is None or active.empty:
        write_cell(ws, 44, 1, "Active Mode log was not provided.", size=12, color="7F8C8D")
        return
    write_cell(ws, 44, 1, "Dataset", size=14, bold=True)
    plotted_n = int(len(active))
    operator_n = int(active_operator_n) if active_operator_n else plotted_n
    campaign = _campaign_label(active["Campaign"]) if "Campaign" in active.columns else ""
    voice = "VOICE" in campaign.upper()
    info_rows = [
        ["Mode", "Active (UE voice)" if voice else "Active (UE)"],
        ["Archive", ACTIVE_ARCHIVE.name],
        ["Campaign", campaign or "—"],
        ["Home Operator", ACTIVE_OPERATOR],
        ["Type", ACTIVE_CELL_TYPE],
        [f"{ACTIVE_OPERATOR} {ACTIVE_CELL_TYPE} rows", operator_n],
        ["Plotted samples", plotted_n],
        ["CSV files", int(active["Source"].nunique())],
        ["Start time", active["Time"].min()],
        ["End time", active["Time"].max()],
        ["Unique Cell Ids", int(active["Cell Id"].nunique())],
    ]
    write_table(ws, 45, 1, ["Item", "Value"], info_rows, num_formats={1: "#,##0"})
    ws.cell(54, 2).number_format = "yyyy-mm-dd hh:mm:ss"
    ws.cell(55, 2).number_format = "yyyy-mm-dd hh:mm:ss"
    write_cell(ws, 44, 4, f"Radio KPIs ({ACTIVE_OPERATOR}, plotted rows)", size=14, bold=True)
    stats_headers = ["KPI", "Count", "Min", "P5", "Median", "Mean", "P95", "Max"]
    stats_rows = []
    for name in ("RSRP", "RSRQ", "SINR"):
        st = kpi_stats(active[name])
        stats_rows.append(
            [name, st["Count"], st["Min"], st["P5"], st["Median"], st["Mean"], st["P95"], st["Max"]]
        )
    write_table(
        ws, 45, 4, stats_headers, stats_rows,
        num_formats={1: "#,##0", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00", 6: "0.00", 7: "0.00"},
    )
    ws.merge_cells(start_row=50, start_column=4, end_row=51, end_column=11)
    write_cell(
        ws, 50, 4,
        f"Maps and bad spots use Home Operator {ACTIVE_OPERATOR} and Type {ACTIVE_CELL_TYPE}. "
        f"Where several of those rows share a latitude and longitude, the strongest RSRP is kept. "
        "That row is one band, and its RSRQ and SINR are the values drawn. "
        "Cell Id is the Cell Id column, or the last number of the CGI when that column is empty. "
        f"The log is {ACTIVE_ARCHIVE.name}. Start and end are the timestamps in that log.",
        size=9, wrap=True,
    )


def build_kpi_sheet(
    ws,
    name: str,
    unit: str,
    map_path: Path | None,
    active_map: Path | None,
    band_maps: dict[str, Path],
    active_band_maps: dict[str, Path] | None = None,
):
    last_col = 18
    em = "\u2014"
    banner(ws, f"  {name} Report", f"  Unit: {unit}  |  IDLE and Active coverage maps  |  Gridlines off", last_col=last_col)
    set_widths(ws, {get_column_letter(i): 14 for i in range(1, last_col + 1)})
    section_bar(ws, 4, f"{name} coverage map {em} IDLE Mode  |  Combined view  |  bad spots", 1, 8)
    section_bar(ws, 4, f"{name} coverage map {em} Active Mode  |  {ACTIVE_OPERATOR}", 10, 18)
    if map_path and map_path.exists():
        add_image(ws, map_path, "A6", width=620, height=760)
    if active_map and active_map.exists():
        add_image(ws, active_map, "J6", width=620, height=760)

    start_row = 6 + 760 // 20 + 3
    block = 700 // 20 + 5
    active_band_maps = active_band_maps or {}
    for i, layer in enumerate(REPORT_LAYERS):
        row = start_row + i * block
        section_bar(ws, row, f"{name} coverage map {em} IDLE Mode_{layer}", 1, 8)
        section_bar(ws, row, f"{name} coverage map {em} Active Mode_{layer}", 10, 18)
        band_path = band_maps.get(layer)
        if band_path and band_path.exists():
            add_image(ws, band_path, f"A{row + 2}", width=560, height=700)
        active_path = active_band_maps.get(layer)
        if active_path and active_path.exists():
            add_image(ws, active_path, f"J{row + 2}", width=560, height=700)


def _kind_label(kind: object, kpi: object = None) -> str:
    text = str(kind or "")
    if text == "discrete":
        return "Discrete area \u2265 1 km\u00b2"
    return consecutive_label(str(kpi)) if kpi in BAD_SPOT_RULES else "Consecutive \u2265 200 m"


BAD_SPOT_KPI_NOTES = {
    "RSRP": (
        "Poor is RSRP below -120 dBm. Only the red bin is a bad spot. Light red, from -115 through -120, is not a bad spot. "
        "Every poor run is covered by a circle. Nearby runs join one circle, up to 10 km across, "
        "and a longer poor road is split into several circles. "
        "The combined view at the top is the all-band best-server map on satellite imagery, the same basemap as the "
        "RSRP sheet. The legend is on the left with the sample count of each colour, and there are no circles. "
        "Best server is the strongest band per scanner sweep, so a weak band is never mapped as the coverage of a road."
    ),
    "RSRQ": (
        "Poor is RSRQ at or below -20 dB. That sits inside the red bin (-30 to -17); red above -20 is not a bad spot. "
        "Every poor run is covered by a circle. "
        "Nearby runs join one circle, up to 10 km across, and separate pockets stay separate circles. "
        "The combined view at the top is the all-band best-server map on satellite imagery, the same basemap as the "
        "RSRP sheet. The legend is on the left with the sample count of each colour, and there are no circles."
    ),
    "SINR": (
        "Poor is SINR at or below 0 dB (orange and red), and every poor run is covered by a circle. "
        "Nearby runs join one circle, up to 10 km across. Where the poor road is spread out, "
        "several circles sit inside the same 10 km neighbourhood instead of one circle covering the good road between them. "
        "The combined view at the top is the all-band best-server map on satellite imagery, with no circles."
    ),
}


def place_solution_summary(ws, row: int, plans: list[dict]) -> int:
    """Count of tilt, 4th sector, and new site, just above the bad-spot table."""
    counts = {"tune": 0, "fourth": 0, "new": 0}
    for plan in plans:
        kind = str(plan.get("kind") or "")
        if kind in counts:
            counts[kind] += 1
    write_cell(ws, row, 1, "Summary of proposed solutions", size=14, bold=True, color=NAVY)
    write_table(
        ws,
        row + 2,
        1,
        ["Solution", "Count"],
        [
            ["Tilt and azimuth", counts["tune"]],
            ["4th sector, L900", counts["fourth"]],
            ["New 3-sector site", counts["new"]],
            ["All bad spots", len(plans)],
        ],
        num_formats={1: "#,##0"},
    )
    return row + 8


def _bad_spot_table_rows(spots: pd.DataFrame, *, include_kpi: bool = True, plans: list[dict] | None = None) -> list[list]:
    rows = []
    for n, (_, s) in enumerate(spots.iterrows(), start=1):
        top_cell = "" if pd.isna(s.get("top_cell", pd.NA)) else int(s["top_cell"])
        length_m = s.get("length_m", np.nan)
        area_km2 = s.get("area_km2", np.nan)
        row = [
            f"Circle {n}",
            _kind_label(s.get("kind"), s.get("kpi")),
        ]
        if include_kpi:
            row.append(str(s["kpi"]))
        row.extend(
            [
                float(s["lat"]),
                float(s["lon"]),
                int(s["n"]),
                int(s["n_poor"]),
                float(s["map_poor_share"]) if "map_poor_share" in s.index and pd.notna(s.get("map_poor_share")) else float(s["poor_pct"]),
                float(s["mean"]),
                float(s["mean_rsrp"]),
                float(s["mean_rsrq"]),
                float(s["mean_sinr"]),
                None if pd.isna(length_m) else float(length_m),
                None if pd.isna(area_km2) else float(area_km2),
                top_cell,
            ]
        )
        if plans is not None:
            plan = plans[n - 1]
            row.extend([plan.get("site", ""), plan.get("population", ""), plan.get("final", "")])
        rows.append(row)
    return rows


def _detection_summary(kpi: str, spots: pd.DataFrame) -> str:
    """State how many spots the rules found and how many of them are drawn."""
    if spots is None or spots.empty:
        return (
            f"The rules find 0 poor {kpi} circle(s) and 0 discrete area(s). "
            f"No {kpi} location meets either rule."
        )
    total_consec = int(spots.attrs.get("total_consecutive", int((spots["kind"] == "consecutive").sum())))
    total_discrete = int(spots.attrs.get("total_discrete", int((spots["kind"] == "discrete").sum())))
    shown_consec = int((spots["kind"] == "consecutive").sum())
    minimum = min_consecutive_m(kpi)
    if minimum <= 0:
        text = (
            f"The rules find {total_consec:,} poor {kpi} circle(s), each drawn up to 10 km and together "
            f"covering every poor run, and {total_discrete:,} discrete area(s) of at least 1 km\u00b2."
        )
    else:
        text = (
            f"The rules find {total_consec:,} consecutive {kpi} run(s) of at least {minimum:,.0f} m "
            f"and {total_discrete:,} discrete area(s) of at least 1 km\u00b2."
        )
    if total_consec > shown_consec:
        text += f" The {shown_consec:,} longest runs are mapped and listed below."
    longest_short = float(spots.attrs.get("longest_short_m", 0.0))
    if minimum > 0 and longest_short >= 50:
        text += (
            f" The longest poor run that misses the rule is {math.floor(longest_short):,} m, "
            "so shorter poor patches on the map are left unmarked."
        )
    return text


def _zoom_display_height(path: Path | None, spots: pd.DataFrame, width: int = 860) -> int:
    n = 0 if spots is None or spots.empty else int(min(len(spots), 16))
    rows = max(1, (n + 3) // 4)
    return int(min(280 * rows, 1120))


def place_bad_spot_advice(ws, header_row: int, plan: dict) -> None:
    """Suggestion table on the right of one bad-spot map.

    Row heights stay at the default, so the movable circle on that map stays put.
    """
    title = header_row + 2
    find = header_row + 4
    head = header_row + 10
    ws.merge_cells(start_row=title, start_column=9, end_row=title, end_column=16)
    write_cell(
        ws, title, 9, "Suggestions and solutions",
        size=11, bold=True, color=NAVY, fill_color=LIGHT, align="center",
    )
    for col in range(9, 17):
        ws.cell(title, col).fill = fill(LIGHT)
        ws.cell(title, col).font = font(11, bold=True, color=NAVY)
        ws.cell(title, col).alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=find, start_column=9, end_row=find + 4, end_column=9)
    ws.merge_cells(start_row=find, start_column=10, end_row=find + 4, end_column=16)
    write_cell(ws, find, 9, "Findings", size=11, bold=True, color=NAVY, align="center", wrap=True)
    write_cell(ws, find, 10, plan.get("finding", ""), size=11, color=NAVY, align="left", wrap=True)
    ws.cell(find, 9).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.cell(find, 10).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.merge_cells(start_row=head, start_column=10, end_row=head, end_column=11)
    ws.merge_cells(start_row=head, start_column=12, end_row=head, end_column=16)
    for col, text in ((9, "Step"), (10, "Action"), (12, "Proposal")):
        write_cell(ws, head, col, text, size=11, bold=True, color=NAVY, fill_color=LIGHT, align="center")
        ws.cell(head, col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(head, col).border = THIN
    for col in range(9, 17):
        ws.cell(head, col).fill = fill(LIGHT)
        ws.cell(head, col).border = THIN
        ws.cell(head, col).font = font(11, bold=True, color=NAVY)
    for index, step in enumerate(plan.get("steps") or []):
        row = head + 1 + index * 4
        ws.merge_cells(start_row=row, start_column=9, end_row=row + 3, end_column=9)
        ws.merge_cells(start_row=row, start_column=10, end_row=row + 3, end_column=11)
        ws.merge_cells(start_row=row, start_column=12, end_row=row + 3, end_column=16)
        write_cell(ws, row, 9, step.get("step", ""), size=12, bold=True, color=NAVY, align="center", wrap=True)
        write_cell(ws, row, 10, step.get("action", ""), size=11, bold=True, color=NAVY, align="center", wrap=True)
        write_cell(ws, row, 12, step.get("proposal", ""), size=11, color=NAVY, align="left", wrap=True)
        for col in (9, 10, 12):
            ws.cell(row, col).alignment = Alignment(
                horizontal="center" if col < 12 else "left", vertical="top", wrap_text=True,
            )
            ws.cell(row, col).border = THIN


def build_kpi_bad_spot_sheet(
    ws,
    kpi: str,
    spots: pd.DataFrame,
    combined_map: Path,
    serial_maps: list[Path] | None,
    *,
    mode: str = "IDLE Mode",
    samples: pd.DataFrame | None = None,
) -> None:
    """Combined route with no circles, then one large map per bad spot down the left side."""
    last_col = 17
    map_span = 8
    em = "\u2014"
    banner(
        ws,
        f"  Bad Spot Analysis  {em}  {kpi}",
        f"  {mode}  |  Combined view has no circles  |  Bad spots in order  |  No numbering",
        last_col=last_col,
    )
    set_widths(ws, {get_column_letter(i): 13 for i in range(1, last_col + 1)})
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["O"].width = 18
    ws.column_dimensions["P"].width = 16
    ws.column_dimensions["Q"].width = 64
    section_bar(ws, 3, mode, 1, last_col)
    ws.merge_cells(start_row=4, start_column=1, end_row=7, end_column=last_col)
    active_rule = ""
    if mode.startswith("Active"):
        active_rule = (
            f"Active Mode uses Home Operator {ACTIVE_OPERATOR} and Type {ACTIVE_CELL_TYPE}, "
            "so one band is plotted at a location: the strongest RSRP. "
            "The RSRQ and SINR maps use that same row. "
        )
    note = (
        active_rule
        + BAD_SPOT_KPI_NOTES[kpi]
        + " Maps and rules read the same 50 m road bins: a bin is coloured by the median of the "
        "samples measured in it and counts as poor only when that median is "
        + ("below -120 dBm" if kpi == "RSRP" else "at or below the threshold")
        + ", and poor bins are drawn on top of good ones. A circle can therefore only sit on "
        "colour you can see. The combined view has no circles. Each bad spot below is one large map, "
        "Bad Spot 1 then Bad Spot 2, and its orange circle is a separate shape named Circle 1, "
        "Circle 2, and so on, in the same order as the table, so you can drag it. The RSRP, RSRQ and "
        "SINR sheets do not repeat these circles. Each bad-spot map draws a transparent cyan outline "
        "for every sector of a site that falls in that view, from System DUMP.xlsb. No sector of "
        "that site is left out, and the outline is not filled. The label on each arm is the Sector "
        "name. A yellow dashed line starts at the end of the serving sector, not at the middle of the site, "
        "and runs to each red sample inside the orange circle. That cell is the Cell Id on the sample, "
        "on L900, L1800, L2100 or L2600, and it has to be in System DUMP.xlsb. Every serving cell is drawn. "
        "The length beside the red is the full red stretch, and a short gap of good coverage stays in that length. "
        "The columns to the right of each map are a suggestion table. Azimuth and tilt are used "
        "when a sector is within about 2.5 km, and downtilt is limited to 2°. A larger downtilt "
        "becomes a 4th L900 sector on a site within 1 km. A 4th sector also replaces a new site "
        "when a host inside 3 km can aim at the bad spot. A new 3-sector site is kept only past that. "
        "A blank with no plot inside the circle is counted "
        "with that bad spot. The list at the bottom has every bad spot, with Site, Population "
        "and Final solution on the right. "
        "Several circles can sit inside one 10 km neighbourhood when the poor "
        "road is not one stretch. A discrete patch of at least 1 km\u00b2 is a dotted circle. "
        + _detection_summary(kpi, spots)
        + " Numbers are not drawn on the maps."
    )
    if mode.startswith("Active"):
        note = note.replace(
            "Best server is the strongest band per scanner sweep, so a weak band is never mapped as the coverage of a road.",
            "That map keeps the strongest RSRP at each latitude and longitude, so only one Cell Id and one band are drawn there.",
        )
    write_cell(
        ws,
        4,
        1,
        note,
        size=11,
        wrap=True,
    )
    for row in range(4, 8):
        ws.row_dimensions[row].height = 36

    section_bar(ws, 9, f"Combined view {em} all-band {mode}  |  {kpi}  |  no circles", 1, last_col)
    map_h = 0
    if combined_map.exists():
        _, map_h = add_image(ws, combined_map, "A11", width=760, height=900)

    cursor = 11 + max(map_h, 660) // 20 + 3
    listed = spots if spots is None or len(spots) <= BAD_SPOT_TABLE_MAX else spots.head(BAD_SPOT_TABLE_MAX)
    plans = None
    if samples is not None and listed is not None and not listed.empty:
        plans = [bad_spot_advice(samples, kpi, listed.iloc[i]) for i in range(len(listed))]
    maps = [p for p in (serial_maps or []) if p is not None and p.exists()]
    for i, path in enumerate(maps, start=1):
        section_bar(ws, cursor, f"Bad Spot {i}", 1, map_span)
        if plans is not None and i <= len(plans):
            place_bad_spot_advice(ws, cursor, plans[i - 1])
        _, spot_h = add_image(ws, path, f"A{cursor + 2}", width=780, height=720)
        cursor = cursor + 2 + max(spot_h, 360) // 20 + 3
    table_row = cursor if maps else cursor + 1
    if plans:
        table_row = place_solution_summary(ws, table_row, plans)

    table_title = f"Identified {kpi} bad spots"
    if spots is not None and len(spots) > BAD_SPOT_TABLE_MAX:
        table_title = f"{table_title}  \u2014  {BAD_SPOT_TABLE_MAX:,} longest of {len(spots):,}"
    write_cell(ws, table_row, 1, table_title, size=14, bold=True)
    headers = [
        "Circle",
        "Type",
        "Latitude",
        "Longitude",
        "Samples",
        "Poor samples",
        "Poor share",
        "Mean KPI",
        "Mean RSRP",
        "Mean RSRQ",
        "Mean SINR",
        "Length (m)",
        "Area (km\u00b2)",
        "Top Cell Id",
        "Site",
        "Population",
        "Final solution",
    ]
    rows = _bad_spot_table_rows(listed, include_kpi=False, plans=plans)
    write_table(
        ws,
        table_row + 1,
        1,
        headers,
        rows,
        num_formats={
            2: "0.00000",
            3: "0.00000",
            4: "#,##0",
            5: "#,##0",
            6: "0.0%",
            7: "0.00",
            8: "0.00",
            9: "0.00",
            10: "0.00",
            11: "#,##0",
            12: "0.00",
        },
    )
    if plans is not None:
        for offset in range(len(rows)):
            cell = ws.cell(table_row + 2 + offset, 17)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[table_row + 2 + offset].height = 48
    note_row = table_row + 2 + len(rows)
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row + 1, end_column=13)
    if kpi == "SINR":
        combined_line = "The combined view is the all-band coverage map on satellite imagery and has no circles. "
    else:
        combined_line = (
            "The combined view is the all-band coverage map on satellite imagery, the same basemap as the RSRP sheet. "
            "The legend is on the left with the sample count of each colour, and there are no circles. "
        )
    write_cell(
        ws,
        note_row,
        1,
        combined_line
        + "Each bad spot above is one large map. Its orange circle is a separate shape with the same "
        "name as the Circle column, so it can be dragged on its own. The space to the right of each "
        "map is that suggestion table. Site, Population and Final solution are filled for every bad spot "
        "in the table below. A blank with no plot inside the circle is counted with that bad spot. "
        "Poor share is the share of raw samples inside the spot box that are at or "
        "below the threshold. Length is the poor road inside the circle, and a single 50 m bin is "
        "reported as 50 m. Map numbers are omitted.",
        size=9,
        wrap=True,
    )
    ws.row_dimensions[note_row].height = 22
    ws.row_dimensions[note_row + 1].height = 22


def style_smooth_line(series, color: str, width=25000) -> None:
    hexcol = color.lstrip("#")
    series.graphicalProperties.line.solidFill = hexcol
    series.graphicalProperties.line.width = width
    series.smooth = True
    series.marker = Marker(symbol="none")


def add_vs_line_chart(ws_data, ws_dest, anchor, title, cat_col, data_min, data_max, min_row, max_row, y_title, colors, width=16, height=8):
    chart = LineChart()
    chart.style = 10
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = "RSRP (dBm)"
    chart.width = width
    chart.height = height
    chart.legend = None
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_min, max_col=data_max, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    for i, color in enumerate(colors):
        if i >= len(chart.series):
            break
        style_smooth_line(chart.series[i], color)
    if "RSRQ" in y_title:
        y_min, y_max, y_major, y_fmt = -20, 0, 5, "0.0"
    else:
        y_min, y_max, y_major, y_fmt = 0, 30, 5, "0"
    apply_xy_levels(
        chart, "RSRP (dBm)", y_title,
        y_min=y_min, y_max=y_max, show_all_x_ticks=False, x_tick_skip=5, rotate_x=False,
        y_major_unit=y_major, y_num_fmt=y_fmt,
    )
    reserve_plot_area(chart, left=0.12, top=0.14, width=0.82, height=0.70)
    ws_dest.add_chart(chart, anchor)
    return chart


def add_dual_axis_vs_chart(ws_data, ws_dest, anchor, n_rows, width=16, height=9, cat_col=50, sinr_col=51, rsrq_col=52):
    cats = Reference(ws_data, min_col=cat_col, min_row=2, max_row=1 + n_rows)
    sinr = LineChart()
    sinr.style = 10
    sinr.title = "RSRP vs SINR vs RSRQ"
    sinr.y_axis.title = "SINR (dB)"
    sinr.x_axis.title = "RSRP (dBm)"
    sinr.width = width
    sinr.height = height
    sinr.add_data(Reference(ws_data, min_col=sinr_col, min_row=1, max_row=1 + n_rows), titles_from_data=True)
    sinr.set_categories(cats)
    style_smooth_line(sinr.series[0], VS_SINR_LINE)
    sinr.legend.position = "b"

    rsrq = LineChart()
    rsrq.y_axis.axId = 200
    rsrq.y_axis.title = "RSRQ (dB)"
    rsrq.add_data(Reference(ws_data, min_col=rsrq_col, min_row=1, max_row=1 + n_rows), titles_from_data=True)
    style_smooth_line(rsrq.series[0], VS_RSRQ_LINE)
    sinr.y_axis.crosses = "min"
    rsrq.y_axis.crosses = "max"
    apply_xy_levels(
        sinr, "RSRP (dBm)", "SINR (dB)",
        y_min=0, y_max=30, show_all_x_ticks=False, x_tick_skip=5, rotate_x=False,
        y_major_unit=5, layout=False, y_num_fmt="0",
    )
    reserve_plot_area(sinr, left=0.12, top=0.14, width=0.74, height=0.58)
    rsrq.y_axis.delete = False
    rsrq.y_axis.axPos = "r"
    rsrq.y_axis.tickLblPos = "nextTo"
    rsrq.y_axis.majorTickMark = "out"
    rsrq.y_axis.minorTickMark = "none"
    rsrq.y_axis.title = "RSRQ (dB)"
    _title_outside(rsrq.y_axis.title)
    rsrq.y_axis.scaling.min = -20
    rsrq.y_axis.scaling.max = 0
    rsrq.y_axis.numFmt = "0.0"
    try:
        rsrq.y_axis.majorUnit = 10
    except Exception:
        pass
    sinr += rsrq
    ws_dest.add_chart(sinr, anchor)


def save_rsrp_scatter(df: pd.DataFrame, ycol: str, path: Path, ylabel: str, title: str, ylim=None) -> Path:
    sample = downsample(df.dropna(subset=["RSRP", ycol]), 18000)
    fig, ax = plt.subplots(figsize=(9.2, 6.2))
    ax.scatter(sample["RSRP"], sample[ycol], s=6, alpha=0.22, c="#1F618D", linewidths=0)
    x = sample["RSRP"].to_numpy()
    y = sample[ycol].to_numpy()
    if len(x) > 20:
        coeff = np.polyfit(x, y, 1)
        xs = np.linspace(np.nanmin(x), np.nanmax(x), 80)
        ax.plot(xs, np.polyval(coeff, xs), color="#17202A", lw=1.8, label="Trend")
        ax.legend(fontsize=8)
    ax.set_xlabel("RSRP (dBm)")
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.set_xlim(-50, -130)
    if ylim:
        ax.set_ylim(*ylim)
    ax.xaxis.set_major_locator(MultipleLocator(5))
    ax.yaxis.set_major_locator(MultipleLocator(5 if "SINR" in ylabel else 2))
    ax.tick_params(axis="both", labelsize=8)
    ax.grid(True, alpha=0.35)
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white")
    plt.close(fig)
    return path


def save_binned_vs_preview(df: pd.DataFrame, path: Path) -> Path:
    pair = df.dropna(subset=["RSRP", "SINR", "RSRQ"]).copy()
    pair["rsrp_bin"] = pair["RSRP"].round().astype(int)
    g = pair.groupby("rsrp_bin")[["SINR", "RSRQ"]].mean()
    g = g.loc[(g.index <= -80) & (g.index >= -125)].sort_index(ascending=False)
    fig, axes = plt.subplots(1, 3, figsize=(18.6, 5.2))

    ax = axes[0]
    ax.plot(g.index, g["SINR"], color=f"#{VS_SINR_LINE}", lw=2.4)
    ax.set_title("RSRP vs SINR")
    ax.set_xlabel("RSRP (dBm)")
    ax.set_ylabel("SINR (dB)")
    ax.set_xlim(-80, -125)
    ax.set_ylim(0, 30)
    ax.set_xticks(range(-80, -126, -5))
    ax.yaxis.set_major_locator(MultipleLocator(5))
    ax.grid(True, axis="y", color="#D5D8DC", lw=0.8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax = axes[1]
    ax.plot(g.index, g["RSRQ"], color=f"#{VS_RSRQ_LINE}", lw=2.4)
    ax.set_title("RSRP vs RSRQ")
    ax.set_xlabel("RSRP (dBm)")
    ax.set_ylabel("RSRQ (dB)")
    ax.set_xlim(-80, -125)
    ax.set_ylim(-20, 0)
    ax.set_xticks(range(-80, -126, -5))
    ax.yaxis.set_major_locator(MultipleLocator(5))
    ax.grid(True, axis="y", color="#D5D8DC", lw=0.8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax = axes[2]
    ax.plot(g.index, g["SINR"], color=f"#{VS_SINR_LINE}", lw=2.4, label="SINR")
    ax.set_ylim(0, 30)
    ax.set_ylabel("SINR (dB)")
    ax.yaxis.set_major_locator(MultipleLocator(5))
    ax2 = ax.twinx()
    ax2.plot(g.index, g["RSRQ"], color=f"#{VS_RSRQ_LINE}", lw=2.4, label="RSRQ")
    ax2.set_ylim(-20, 0)
    ax2.set_ylabel("RSRQ (dB)")
    ax2.yaxis.set_major_locator(MultipleLocator(10))
    ax.set_title("RSRP vs SINR vs RSRQ")
    ax.set_xlabel("RSRP (dBm)")
    ax.set_xlim(-80, -125)
    ax.set_xticks(range(-80, -126, -5))
    ax.grid(True, axis="y", color="#D5D8DC", lw=0.8)
    ax.spines["top"].set_visible(False)
    handles = ax.get_lines() + ax2.get_lines()
    labels = [h.get_label() for h in handles]
    ax.legend(handles, labels, loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=2, frameon=False)
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white")
    plt.close(fig)
    return path


def build_rsrp_vs_sheet(
    ws, data_ws, n_rows: int, scatter_sinr: Path, scatter_rsrq: Path,
    *,
    title: str = "  RSRP vs KPIs IDLE Mode",
    subtitle: str = "  Idle V2.0 scanner  |  Smooth mean lines and scatter with trend  |  Gridlines off",
    cat_col: int = 50,
    sinr_col: int = 51,
    rsrq_col: int = 52,
) -> None:
    banner(ws, title, subtitle, last_col=12)
    set_widths(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    add_vs_line_chart(
        data_ws, ws, "A4", "RSRP vs SINR",
        cat_col, sinr_col, sinr_col, 2, 1 + n_rows, "SINR (dB)",
        [VS_SINR_LINE], width=15, height=8,
    )
    add_vs_line_chart(
        data_ws, ws, "I4", "RSRP vs RSRQ",
        cat_col, rsrq_col, rsrq_col, 2, 1 + n_rows, "RSRQ (dB)",
        [VS_RSRQ_LINE], width=15, height=8,
    )
    add_dual_axis_vs_chart(
        data_ws, ws, "A22", n_rows, width=22, height=8,
        cat_col=cat_col, sinr_col=sinr_col, rsrq_col=rsrq_col,
    )
    write_cell(ws, 48, 1, "RSRP vs SINR scatter (trend line)", size=14, bold=True)
    write_cell(ws, 48, 8, "RSRP vs RSRQ scatter (trend line)", size=14, bold=True)
    if scatter_sinr.exists():
        add_image(ws, scatter_sinr, "A50", width=520, height=350)
    if scatter_rsrq.exists():
        add_image(ws, scatter_rsrq, "H50", width=520, height=350)


def build_sample_log(ws: Worksheet, df: pd.DataFrame, n: int = 8000) -> None:
    banner(ws, "  Sample Drive Log IDLE Mode", "  Evenly spaced subset for inspection  |  Full merged file is CSV  |  No freeze / no gridlines", last_col=10)
    sample = downsample(df, max_points=n).copy()
    cols = ["Time", "RSRP", "RSRQ", "SINR", "Longitude", "Latitude", "Cell Id", "DL EARFCN", "Source", "Band"]
    sample = sample[cols]
    for c, w in zip(cols, (22, 12, 12, 12, 14, 14, 14, 14, 10, 22)):
        ws.column_dimensions[get_column_letter(cols.index(c) + 1)].width = w
    headers = cols
    for i, h in enumerate(headers, start=1):
        write_cell(ws, 4, i, h, bold=True, color=WHITE, fill_color=NAVY, align="center")
    for r_i, row in enumerate(sample.itertuples(index=False), start=5):
        bg = LIGHT if r_i % 2 == 0 else WHITE
        values = list(row)
        for c_i, value in enumerate(values, start=1):
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            if hasattr(value, "item"):
                try:
                    value = value.item()
                except Exception:
                    pass
            if pd.isna(value):
                value = None
            cell = write_cell(ws, r_i, c_i, value, fill_color=bg, align="center")
            cell.border = THIN
            if c_i == 1 and value is not None:
                cell.number_format = "yyyy-mm-dd hh:mm:ss.000"
            if c_i in (2, 3, 4, 5, 6):
                cell.number_format = "0.000"


def _chart_title_text(title) -> str | None:
    if title is None:
        return None
    try:
        return str(title.tx.rich.p[0].r[0].t)
    except Exception:
        return str(title)


def verify_report(path: Path) -> dict:
    wb = load_workbook(path)
    info = {"sheets": {}, "chart_titles": [], "axis_check": [], "problems": []}
    for name in wb.sheetnames:
        ws = wb[name]
        freeze = ws.freeze_panes
        grid = ws.sheet_view.showGridLines
        print_grid = ws.print_options.gridLines
        n_charts = len(ws._charts)
        n_images = len(ws._images)
        info["sheets"][name] = {
            "freeze_panes": freeze,
            "showGridLines": grid,
            "print_gridLines": print_grid,
            "charts": n_charts,
            "images": n_images,
        }
        if freeze not in (None, "A1"):
            info["problems"].append(f"{name}: freeze_panes={freeze}")
        if grid is not False:
            info["problems"].append(f"{name}: showGridLines={grid}")
        if print_grid:
            info["problems"].append(f"{name}: print gridlines on")
        for ch in ws._charts:
            title = _chart_title_text(getattr(ch, "title", None))
            x_title = _chart_title_text(getattr(ch.x_axis, "title", None))
            y_title = _chart_title_text(getattr(ch.y_axis, "title", None))
            x_pos = getattr(ch.x_axis, "axPos", None)
            y_pos = getattr(ch.y_axis, "axPos", None)
            info["chart_titles"].append((name, title, type(ch).__name__))
            info["axis_check"].append((name, title, x_title, y_title, x_pos, y_pos))
            if x_pos != "b":
                info["problems"].append(f"{name}/{title}: x axPos={x_pos} (expected b)")
            if y_pos not in ("l", "r"):
                info["problems"].append(f"{name}/{title}: y axPos={y_pos} (expected l/r)")
            if not x_title:
                info["problems"].append(f"{name}/{title}: missing horizontal axis title")
            if not y_title:
                info["problems"].append(f"{name}/{title}: missing vertical axis title")
    for name in (
        "Bad Spot RSRP",
        "Bad Spot RSRP_Active",
    ):
        if name not in info["sheets"]:
            info["problems"].append(f"missing {name} sheet")
        elif info["sheets"][name]["images"] < 1:
            info["problems"].append(f"{name}: expected combined map image")
    if "RSRP vs KPIs_Active" not in info["sheets"]:
        info["problems"].append("missing RSRP vs KPIs_Active sheet")
    elif info["sheets"]["RSRP vs KPIs_Active"]["charts"] < 3:
        info["problems"].append("RSRP vs KPIs_Active: expected three charts")
    info["problems"].extend(logic_drift())
    count_re = re.compile(r"find ([\d,]+) .*? and ([\d,]+) discrete")
    for sheet, (runs, areas) in LOGIC_LOCK["counts"].items():
        text = str(wb[sheet].cell(4, 1).value or "") if sheet in wb.sheetnames else ""
        found = count_re.search(text)
        got = None if found is None else (int(found.group(1).replace(",", "")), int(found.group(2).replace(",", "")))
        if got != (runs, areas):
            info["problems"].append(f"{sheet}: bad-spot count {got}, locked {(runs, areas)}")
    wb.close()
    with zipfile.ZipFile(path) as zf:
        ellipses = sum(
            zf.read(n).decode("utf-8").count('prst="ellipse"')
            for n in zf.namelist()
            if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n)
        )
        info["circle_shapes"] = ellipses
        # One movable circle on each serial bad-spot map. The combined view has none.
        expected_shapes = 0
        for sheet, (runs, areas) in LOGIC_LOCK["counts"].items():
            n = runs + areas
            shown = min(BAD_SPOT_MAX_ZOOMS, n)
            expected_shapes += shown
            images = info["sheets"].get(sheet, {}).get("images", 0)
            if sheet in info["sheets"] and images != 1 + shown:
                info["problems"].append(f"{sheet}: {images} images, expected {1 + shown}")
        if ellipses != expected_shapes:
            info["problems"].append(f"{ellipses} circle shapes, expected {expected_shapes}")
        for item in zf.namelist():
            if not item.startswith("xl/charts/chart"):
                continue
            xml = zf.read(item).decode("utf-8")
            for cat in re.findall(r"<catAx>.*?</catAx>", xml):
                pos = re.search(r'<axPos val="([^"]+)"', cat)
                if not pos or pos.group(1) != "b":
                    info["problems"].append(f"{item}: category axis axPos is {pos.group(1) if pos else 'missing'} (need b)")
                if 'tickLblPos val="none"' in cat:
                    info["problems"].append(f"{item}: category tick labels hidden")
            for val in re.findall(r"<valAx>.*?</valAx>", xml):
                pos = re.search(r'<axPos val="([^"]+)"', val)
                if not pos or pos.group(1) not in ("l", "r"):
                    info["problems"].append(f"{item}: value axis axPos is {pos.group(1) if pos else 'missing'}")
                if 'tickLblPos val="none"' in val:
                    info["problems"].append(f"{item}: value tick labels hidden")
            if "<a:t></a:t>" in xml or "<a:t/>" in xml:
                info["problems"].append(f"{item}: empty text run would blank axis labels")
    return info


def route_length_km(df: pd.DataFrame) -> float:
    """Kilometres along the drive. A step longer than 500 m is a new pass, not road."""
    work = df.dropna(subset=["Time", "Latitude", "Longitude"]).copy()
    if "Time" in work.columns:
        work["Time"] = pd.to_datetime(work["Time"], errors="coerce")
        work = work.dropna(subset=["Time"])
    if len(work) < 2:
        return 0.0
    sort_cols = ["Source", "Time"] if "Source" in work.columns else ["Time"]
    work = work.sort_values(sort_cols)
    groups = work.groupby("Source", sort=False) if "Source" in work.columns else [(None, work)]
    total = 0.0
    for _, part in groups:
        if len(part) < 2:
            continue
        lat = part["Latitude"].to_numpy(dtype=float)
        lon = part["Longitude"].to_numpy(dtype=float)
        step = np.asarray(_equirect_km(lat[:-1], lon[:-1], lat[1:], lon[1:]), dtype=float)
        total += float(step[step < 0.5].sum())
    return total


def write_active_histograms(ws: Worksheet, df: pd.DataFrame) -> dict:
    """Legend-coloured histogram series for the plotted Active rows."""
    rsrp_bins = np.arange(-140, -48, 2)
    rsrq_bins = np.arange(-24, -2.5, 0.5)
    sinr_bins = np.arange(-15, 31, 1)
    return {
        "RSRP": write_stacked_hist(ws, 60, df["RSRP"], rsrp_bins, rsrp_map_class(df["RSRP"]), MAP_RSRP_LABELS),
        "RSRQ": write_stacked_hist(ws, 68, df["RSRQ"], rsrq_bins, rsrq_map_class(df["RSRQ"]), MAP_RSRQ_LABELS),
        "SINR": write_stacked_hist(ws, 76, df["SINR"], sinr_bins, sinr_map_class(df["SINR"]), MAP_SINR_LABELS),
    }


def add_active_cover_extras(cover, data_ws, blocks, route_path: Path, active: pd.DataFrame) -> None:
    """Active histograms and the DT route. Idle charts and the other sheets stay put."""
    write_cell(cover, 44, 12, "Active histograms (plotted Robi rows)", size=14, bold=True)
    cover.merge_cells(start_row=44, start_column=12, end_row=44, end_column=20)
    add_stacked_col_chart(
        data_ws, cover, "L46", "RSRP histogram — Active",
        blocks["RSRP"]["cat_col"], blocks["RSRP"]["data_min"], blocks["RSRP"]["data_max"],
        2, 1 + blocks["RSRP"]["n"],
        "Samples", "RSRP (dBm)", MAP_RSRP_COLORS, width=9, height=6,
    )
    add_stacked_col_chart(
        data_ws, cover, "R46", "RSRQ histogram — Active",
        blocks["RSRQ"]["cat_col"], blocks["RSRQ"]["data_min"], blocks["RSRQ"]["data_max"],
        2, 1 + blocks["RSRQ"]["n"],
        "Samples", "RSRQ (dB)", MAP_RSRQ_COLORS, width=9, height=6,
    )
    add_stacked_col_chart(
        data_ws, cover, "L58", "SINR histogram — Active",
        blocks["SINR"]["cat_col"], blocks["SINR"]["data_min"], blocks["SINR"]["data_max"],
        2, 1 + blocks["SINR"]["n"],
        "Samples", "SINR (dB)", MAP_SINR_COLORS, width=9, height=6,
    )

    section_bar(cover, 72, "DT Route", 1, 20)
    write_cell(cover, 74, 1, "Drive details", size=14, bold=True)
    write_cell(cover, 74, 12, "DT route plot", size=14, bold=True)
    if route_path.exists():
        add_image(cover, route_path, "L76", width=520, height=640)
    ordered = active.dropna(subset=["Time", "Latitude", "Longitude"]).copy()
    ordered["Time"] = pd.to_datetime(ordered["Time"], errors="coerce")
    ordered = ordered.dropna(subset=["Time"]).sort_values("Time")
    start = ordered.iloc[0]
    end = ordered.iloc[-1]
    span = end["Time"] - start["Time"]
    hours = span.total_seconds() / 3600.0
    layers = pd.to_numeric(active["DL EARFCN"], errors="coerce").map(earfcn_to_layer)
    layer_bits = ", ".join(f"{name} {int(n):,}" for name, n in layers.value_counts().items() if name)
    detail_rows = [
        ["Mode", "Active (UE)"],
        ["Campaign", _campaign_label(active["Campaign"]) if "Campaign" in active.columns else "—"],
        ["Operator", ACTIVE_OPERATOR],
        ["Plotted samples", int(len(active))],
        ["Route length (km)", round(route_length_km(active), 1)],
        ["Duration (hours)", round(hours, 1)],
        ["Start time", start["Time"]],
        ["End time", end["Time"]],
        ["Start latitude", round(float(start["Latitude"]), 5)],
        ["Start longitude", round(float(start["Longitude"]), 5)],
        ["End latitude", round(float(end["Latitude"]), 5)],
        ["End longitude", round(float(end["Longitude"]), 5)],
        ["Unique Cell Ids", int(active["Cell Id"].nunique())],
        ["CSV files", int(active["Source"].nunique()) if "Source" in active.columns else "—"],
    ]
    write_table(cover, 75, 1, ["Item", "Value"], detail_rows)
    cover.cell(80, 2).number_format = "#,##0.0"
    cover.cell(81, 2).number_format = "#,##0.0"
    cover.cell(82, 2).number_format = "yyyy-mm-dd hh:mm:ss"
    cover.cell(83, 2).number_format = "yyyy-mm-dd hh:mm:ss"
    cover.merge_cells(start_row=91, start_column=1, end_row=93, end_column=11)
    write_cell(
        cover, 91, 1,
        "The route is the plotted Robi M-PCell track: one band at each latitude and longitude, "
        "the strongest RSRP. RSRQ and SINR on the histograms are the values from that same row. "
        f"Source is {ACTIVE_ARCHIVE.name}. "
        f"Plotted band mix: {layer_bits}.",
        size=9, wrap=True,
    )
    write_cell(
        cover, 19, 1,
        "Cover — Idle dataset and plots on top. Active Mode dataset, histograms and the DT route are below.",
        wrap=True,
    )


def add_cover_charts(cover, data_ws, blocks) -> None:
    write_cell(cover, 5, 12, "RSRP, RSRQ and SINR plots (native Excel charts with legends)", size=14, bold=True)
    cover.merge_cells(start_row=5, start_column=12, end_row=5, end_column=20)
    add_stacked_col_chart(
        data_ws, cover, "L6", "RSRP histogram",
        blocks["RSRP"]["cat_col"], blocks["RSRP"]["data_min"], blocks["RSRP"]["data_max"],
        2, 1 + blocks["RSRP"]["n"],
        "Samples", "RSRP (dBm)", MAP_RSRP_COLORS, width=11, height=7,
    )
    add_stacked_col_chart(
        data_ws, cover, "R6", "RSRQ histogram",
        blocks["RSRQ"]["cat_col"], blocks["RSRQ"]["data_min"], blocks["RSRQ"]["data_max"],
        2, 1 + blocks["RSRQ"]["n"],
        "Samples", "RSRQ (dB)", MAP_RSRQ_COLORS, width=11, height=7,
    )
    add_stacked_col_chart(
        data_ws, cover, "L23", "SINR histogram",
        blocks["SINR"]["cat_col"], blocks["SINR"]["data_min"], blocks["SINR"]["data_max"],
        2, 1 + blocks["SINR"]["n"],
        "Samples", "SINR (dB)", MAP_SINR_COLORS, width=11, height=7,
    )
    add_col_chart(
        data_ws, cover, "R23", "RSRP ranges (dBm)", RSRP_COLOR,
        12, 13, 2, 1 + blocks["rsrp_q_n"],
        "Samples", "RSRP range", width=11, height=7, show_legend=True,
    )


def save_report_preview(df: pd.DataFrame, path: Path) -> Path:
    fig, axes = plt.subplots(1, 3, figsize=(16, 5.2))
    plot_stacked_map_hist(
        axes[0], df["RSRP"], rsrp_map_class, MAP_RSRP_LABELS, MAP_RSRP_COLORS,
        np.arange(-140, -48, 2), "RSRP histogram", "RSRP (dBm)", "RSRP (dBm)", (-140, -50),
    )
    plot_stacked_map_hist(
        axes[1], df["RSRQ"], rsrq_map_class, MAP_RSRQ_LABELS, MAP_RSRQ_COLORS,
        np.arange(-24, -2.5, 0.5), "RSRQ histogram", "RSRQ (dB)", "RSRQ (dB)", (-24, -3),
    )
    plot_stacked_map_hist(
        axes[2], df["SINR"], sinr_map_class, MAP_SINR_LABELS, MAP_SINR_COLORS,
        np.arange(-15, 31, 1), "SINR histogram", "SINR (dB)", "SINR (dB)", (-15, 30),
    )
    for ax in axes:
        ax.grid(False)
        ax.set_facecolor("white")
    fig.suptitle("Bogura Drive-Test Report — RSRP / RSRQ / SINR", fontsize=13, fontweight="bold")
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white")
    plt.close(fig)
    return path


def _with_layers(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Band" not in out.columns:
        out["Band"] = out["DL EARFCN"].map(earfcn_to_band)
    if "Layer" not in out.columns:
        out["Layer"] = pd.to_numeric(out["DL EARFCN"], errors="coerce").map(earfcn_to_layer)
    return out


def _render_mode(df: pd.DataFrame, tag: str, extent, mode_label: str | None) -> dict:
    """Coverage maps, bad-spot maps and spot tables for one mode.

    tag is "" for Idle (existing file names) and "active" for the Active Mode copies.
    Idle rows that share a latitude, longitude and timestamp collapse to the
    strongest RSRP, then the strongest band in each scanner sweep is kept.
    Active auto mode keeps one Cell Id and one band per latitude and longitude,
    the strongest RSRP, and RSRQ and SINR stay on that row.
    """
    suffix = f"_{tag}" if tag else ""
    if tag == "active":
        base = strongest_location(df)
        map_df = base
    else:
        base = strongest_fix(df)
        map_df = best_server(base)
    plot_rsrp_coverage_map(map_df, PLOTS_DIR / f"excel_route_rsrp{suffix}.png", extent=extent, mode_label=mode_label)
    plot_rsrq_coverage_map(map_df, PLOTS_DIR / f"excel_route_rsrq{suffix}.png", extent=extent, mode_label=mode_label)
    plot_sinr_coverage_map(map_df, PLOTS_DIR / f"excel_route_sinr{suffix}.png", extent=extent, mode_label=mode_label)
    rsrp_band: dict[str, Path] = {}
    rsrq_band: dict[str, Path] = {}
    sinr_band: dict[str, Path] = {}
    for layer in REPORT_LAYERS:
        part = filter_by_layer(base, layer)
        if tag != "active":
            part = best_server(part)
        rsrp_band[layer] = plot_rsrp_coverage_map(
            part, PLOTS_DIR / f"excel_route_rsrp_{layer}{suffix}.png",
            layer=layer, extent=extent, mode_label=mode_label,
        )
        rsrq_band[layer] = plot_rsrq_coverage_map(
            part, PLOTS_DIR / f"excel_route_rsrq_{layer}{suffix}.png",
            layer=layer, extent=extent, mode_label=mode_label,
        )
        sinr_band[layer] = plot_sinr_coverage_map(
            part, PLOTS_DIR / f"excel_route_sinr_{layer}{suffix}.png",
            layer=layer, extent=extent, mode_label=mode_label,
        )
    rsrp_spots = find_bad_spots(map_df, "RSRP")
    rsrq_spots = find_bad_spots(map_df, "RSRQ")
    sinr_spots = find_bad_spots(map_df, "SINR")
    title_bit = f" — {mode_label}" if mode_label else ""
    bad_maps = {}
    zoom_maps = {}
    for kpi, plot_map, classify, labels, colors in (
        ("RSRP", plot_rsrp_bad_spot_map, rsrp_map_class, MAP_RSRP_LABELS, MAP_RSRP_COLORS),
        ("RSRQ", plot_rsrq_bad_spot_map, rsrq_map_class, MAP_RSRQ_LABELS, MAP_RSRQ_COLORS),
        ("SINR", plot_sinr_bad_spot_map, sinr_map_class, MAP_SINR_LABELS, MAP_SINR_COLORS),
    ):
        key = kpi.lower()
        spots = {"RSRP": rsrp_spots, "RSRQ": rsrq_spots, "SINR": sinr_spots}[kpi]
        total = int(spots.attrs.get("total_consecutive", len(spots))) if spots is not None else 0
        discrete = int(spots.attrs.get("total_discrete", 0)) if spots is not None else 0
        print(f"{kpi} {tag or 'idle'}: {total} poor runs, {discrete} discrete", flush=True)
        extra = {}
        if kpi != "SINR":
            mode_name = "Active" if tag == "active" else "Idle"
            extra["legend_title"] = f"{mode_name} by {kpi}"
        bad_maps[kpi] = plot_map(
            map_df, spots, PLOTS_DIR / f"excel_bad_spots_{key}{suffix}.png",
            extent=extent, title=f"Combined view — {kpi}{title_bit}", shapes=False, outlines=False,
            **extra,
        )
        zoom_maps[kpi] = plot_bad_spot_serial_maps(
            map_df, spots, PLOTS_DIR, f"excel_bad_spots_{key}_spot{suffix}",
            kpi, classify(map_df[kpi]), labels, colors,
            basemap_style="satellite",
            route_casing=True,
        )
    rsrp_bad, rsrq_bad, sinr_bad = bad_maps["RSRP"], bad_maps["RSRQ"], bad_maps["SINR"]
    rsrp_zooms, rsrq_zooms, sinr_zooms = zoom_maps["RSRP"], zoom_maps["RSRQ"], zoom_maps["SINR"]
    return {
        "map_df": map_df,
        "rsrp_band": rsrp_band, "rsrq_band": rsrq_band, "sinr_band": sinr_band,
        "rsrp_spots": rsrp_spots, "rsrq_spots": rsrq_spots, "sinr_spots": sinr_spots,
        "rsrp_bad": rsrp_bad, "rsrq_bad": rsrq_bad, "sinr_bad": sinr_bad,
        "rsrp_zooms": rsrp_zooms, "rsrq_zooms": rsrq_zooms, "sinr_zooms": sinr_zooms,
    }


def build_report(df: pd.DataFrame, active: pd.DataFrame | None = None, out_path: Path = REPORT_XLSX) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    df = _with_layers(df)
    active_raw_n = 0
    active_robi = None
    if active is not None and not active.empty:
        active = _with_layers(active)
        active_raw_n = len(active)
        active_robi = select_active_report(active)
        print(
            f"Active Mode {ACTIVE_OPERATOR} {ACTIVE_CELL_TYPE}: "
            f"{len(active_robi):,} rows of {active_raw_n:,}"
        )

    idle_fixed = strongest_fix(df)
    print(f"Idle strongest row per lat/lon/time: {len(idle_fixed):,} of {len(df):,}", flush=True)
    idle_track = best_server(idle_fixed)
    extent = coverage_extent(idle_track)
    print(f"Idle extent {tuple(round(v, 5) for v in extent)}", flush=True)
    idle = _render_mode(df, "", extent, None)
    active_view = None
    active_plotted = None
    if active_robi is not None and not active_robi.empty:
        active_plotted = strongest_location(active_robi)
        print(
            f"Active plotted: {len(active_plotted):,} rows, one band per lat/lon",
            flush=True,
        )
        active_extent = coverage_extent(active_plotted)
        print(f"Active extent {tuple(round(v, 5) for v in active_extent)}", flush=True)
        active_view = _render_mode(active_robi, "active", active_extent, "Active")

    scatter_sinr = save_rsrp_scatter(idle_fixed, "SINR", PLOTS_DIR / "excel_scatter_rsrp_sinr.png", "SINR (dB)", "RSRP vs SINR", ylim=(-10, 32))
    scatter_rsrq = save_rsrp_scatter(idle_fixed, "RSRQ", PLOTS_DIR / "excel_scatter_rsrp_rsrq.png", "RSRQ (dB)", "RSRP vs RSRQ", ylim=(-24, 0))
    shutil.copy2(idle["rsrp_bad"], PLOTS_DIR / "08_bad_spots_rsrp.png")
    shutil.copy2(idle["rsrq_bad"], PLOTS_DIR / "08_bad_spots_rsrq.png")
    shutil.copy2(idle["sinr_bad"], PLOTS_DIR / "08_bad_spots_sinr.png")
    save_binned_vs_preview(idle_fixed, PLOTS_DIR / "excel_rsrp_vs_sinr_lines.png")
    save_report_preview(df, PLOTS_DIR / "excel_rsrp_rsrq_sinr_histograms.png")
    if active_view is not None:
        save_rsrp_scatter(
            active_plotted, "SINR", PLOTS_DIR / "excel_scatter_rsrp_sinr_active.png",
            "SINR (dB)", "RSRP vs SINR — Active", ylim=(-10, 32),
        )
        save_rsrp_scatter(
            active_plotted, "RSRQ", PLOTS_DIR / "excel_scatter_rsrp_rsrq_active.png",
            "RSRQ (dB)", "RSRP vs RSRQ — Active", ylim=(-24, 0),
        )

    wb = Workbook()
    cover = wb.active
    cover.title = "Cover"
    rsrp_ws = wb.create_sheet("RSRP")
    rsrq_ws = wb.create_sheet("RSRQ")
    sinr_ws = wb.create_sheet("SINR")
    bad_rsrp_ws = wb.create_sheet("Bad Spot RSRP")
    bad_rsrp_a = wb.create_sheet("Bad Spot RSRP_Active")
    time_ws = wb.create_sheet("RSRP vs KPIs")
    time_active_ws = wb.create_sheet("RSRP vs KPIs_Active")
    sample_ws = wb.create_sheet("Sample Log")
    data_ws = wb.create_sheet("_ChartData")
    data_ws.sheet_state = "hidden"

    blocks = write_chart_data(data_ws, df)
    blocks["vs_n"] = write_vs_block(data_ws, idle_fixed, 50, 51, 52)
    active_vs_n = blocks["vs_n"]
    if active_view is not None:
        active_vs_n = write_vs_block(data_ws, active_plotted, 54, 55, 56)
    build_cover(
        cover, df, active_plotted, active_raw_n,
        0 if active_robi is None else len(active_robi),
    )
    add_cover_charts(cover, data_ws, blocks)
    if active_plotted is not None and not active_plotted.empty:
        route_path = plot_road_map(
            active_plotted, PLOTS_DIR / "excel_cover_active_route.png",
        )
        add_active_cover_extras(
            cover, data_ws, write_active_histograms(data_ws, active_plotted), route_path, active_plotted,
        )
    empty_bands: dict[str, Path] = {}
    build_kpi_sheet(
        rsrp_ws, "RSRP", "dBm", idle["rsrp_bad"],
        None if active_view is None else active_view["rsrp_bad"],
        idle["rsrp_band"], empty_bands if active_view is None else active_view["rsrp_band"],
    )
    build_kpi_sheet(
        rsrq_ws, "RSRQ", "dB", idle["rsrq_bad"],
        None if active_view is None else active_view["rsrq_bad"],
        idle["rsrq_band"], empty_bands if active_view is None else active_view["rsrq_band"],
    )
    build_kpi_sheet(
        sinr_ws, "SINR", "dB", idle["sinr_bad"],
        None if active_view is None else active_view["sinr_bad"],
        idle["sinr_band"], empty_bands if active_view is None else active_view["sinr_band"],
    )
    build_kpi_bad_spot_sheet(
        bad_rsrp_ws, "RSRP", idle["rsrp_spots"], idle["rsrp_bad"], idle["rsrp_zooms"],
        samples=idle["map_df"],
    )
    if active_view is not None:
        build_kpi_bad_spot_sheet(
            bad_rsrp_a, "RSRP", active_view["rsrp_spots"], active_view["rsrp_bad"], active_view["rsrp_zooms"],
            mode="Active Mode", samples=active_view["map_df"],
        )
        build_rsrp_vs_sheet(
            time_active_ws, data_ws, active_vs_n,
            PLOTS_DIR / "excel_scatter_rsrp_sinr_active.png",
            PLOTS_DIR / "excel_scatter_rsrp_rsrq_active.png",
            title="  RSRP vs KPIs Active Mode",
            subtitle=f"  {ACTIVE_OPERATOR} UE  |  RSRQ and SINR from the strongest RSRP row  |  Gridlines off",
            cat_col=54, sinr_col=55, rsrq_col=56,
        )
    build_rsrp_vs_sheet(
        time_ws, data_ws, blocks["vs_n"], scatter_sinr, scatter_rsrq,
        subtitle="  Idle V2.0 scanner  |  RSRQ and SINR from the strongest RSRP row  |  Gridlines off",
    )
    build_sample_log(sample_ws, df)

    for ws in wb.worksheets:
        apply_sheet_view(ws)

    wb.properties.title = "Bogura Drive-Test Report IDLE and Active Mode"
    wb.properties.subject = "IDLE and Active Mode"
    wb.properties.description = (
        "Bogura drive-test report. Idle scanner on the left and on Bad Spot RSRP/RSRQ/SINR. "
        "Active Mode Robi UE on the right and on Bad Spot *_Active and RSRP vs KPIs_Active. "
        "Poor runs use the same thresholds in both modes."
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    n_svg = embed_svg_layers(out_path)
    n_circles = embed_circle_shapes(out_path)
    n_jpg = compress_workbook_images(out_path)
    print(f"Embedded {n_svg} SVG layers and {n_circles} separate bad-spot circles; compressed {n_jpg} maps")
    shutil.copy2(out_path, VERSIONED_XLSX)
    print(f"Wrote Excel report: {out_path}")
    print(f"Wrote versioned copy: {VERSIONED_XLSX}")
    return out_path


_AIM_SENTENCE_OLD = (
    "The label on each arm is the Sector name. A yellow line shows the distance only from a sector whose azimuth "
    "points at that bad spot. A sector that faces another way has no line."
)
_AIM_SENTENCE_NEW = (
    "The label on each arm is the Sector name. A yellow dashed line starts at the end of the serving sector, "
    "not at the middle of the site, and runs to each red sample inside the orange circle. That cell is the Cell Id "
    "on the sample, on L900, L1800, L2100 or L2600, and it has to be in System DUMP.xlsb. Every serving cell is drawn. "
    "The length beside the red is the full red stretch, and a short gap of good coverage stays in that length."
)
_AIM_SENTENCE_CELL = (
    "The label on each arm is the Sector name. A yellow line runs from the serving cell to each red sample inside "
    "the orange circle. That cell is the Cell Id on the sample, on L900, L1800, L2100 or L2600, and it has to be in "
    "System DUMP.xlsb. The length of each red run is written beside it in metres."
)
_PIE_SENTENCE_OLD = (
    "Each bad-spot map draws a sector pie for every nearby site from System DUMP.xlsb. "
    "The label is the sector name. Where a sector shoots toward that bad spot, a yellow line shows the distance."
)
_PIE_SENTENCE_NEW = (
    "Each bad-spot map draws a transparent cyan outline for every sector of a site that falls in that view, "
    "from System DUMP.xlsb. No sector of that site is left out, and the outline is not filled. "
    "Where a sector shoots toward that bad spot, a yellow line shows the distance."
)


def refresh_bad_spot_pies(src: Path | None = None) -> Path:
    """Redraw the serial bad-spot maps and swap them into the current workbook.

    Combined views, KPI sheets, and the circle count stay as they are. Each serial
    map keeps one movable circle, placed again from the new picture.
    """
    import io
    from PIL import Image

    from merge_and_analyze import ACTIVE_MERGED, MERGED_CSV, load_active_merged, load_merged_csv

    src = src or REPORT_XLSX
    if not src.exists():
        raise SystemExit(f"Missing workbook: {src}")
    print("Loading drive logs", flush=True)
    idle = load_merged_csv() if MERGED_CSV.exists() else None
    active = load_active_merged() if ACTIVE_MERGED.exists() else None
    if idle is None or active is None:
        raise SystemExit("Merged drive logs are missing; a full report build is required.")
    active = select_active_report(active)
    print(f"Idle rows {len(idle):,}  Active {ACTIVE_OPERATOR} {ACTIVE_CELL_TYPE} rows {len(active):,}", flush=True)
    idle_track = best_server(idle)
    active_track = best_carrier(active)
    MAP_CIRCLE_PIXELS.clear()
    zoom_for_sheet: dict[str, list[Path]] = {}
    for tag, frame in (("", idle_track), ("active", active_track)):
        suffix = f"_{tag}" if tag else ""
        for kpi, classify, labels, colors in (
            ("RSRP", rsrp_map_class, MAP_RSRP_LABELS, MAP_RSRP_COLORS),
        ):
            print(f"Drawing {kpi} bad spots ({tag or 'idle'})", flush=True)
            spots = find_bad_spots(frame, kpi)
            paths = plot_bad_spot_serial_maps(
                frame, spots, PLOTS_DIR, f"excel_bad_spots_{kpi.lower()}_spot{suffix}",
                kpi, classify(frame[kpi]), labels, colors,
                basemap_style="satellite",
                route_casing=True,
            )
            sheet = f"Bad Spot {kpi}" + ("_Active" if tag else "")
            zoom_for_sheet[sheet] = paths
            print(f"  {sheet}: {len(paths)} maps", flush=True)

    print("Updating the workbook", flush=True)
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    drawings = _sheet_drawings(parts)
    sheet_xml = _worksheet_xml(parts)
    for sheet, paths in zoom_for_sheet.items():
        drawing_name = drawings[sheet]
        rels_name = drawing_name.replace("drawings/", "drawings/_rels/") + ".rels"
        rels = parts[rels_name].decode("utf-8")
        drawing = parts[drawing_name].decode("utf-8")
        rid_media = {}
        for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
            rid = re.search(r'Id="([^"]+)"', rel).group(1)
            target = re.search(r'Target="([^"]+)"', rel).group(1)
            media = target.lstrip("/") if target.startswith("/") else "xl/media/" + target.rsplit("/", 1)[1]
            rid_media[rid] = media
        pictures = []
        for block in re.findall(r"<oneCellAnchor>.*?</oneCellAnchor>", drawing, flags=re.S):
            if "<pic>" not in block:
                continue
            row = int(re.search(r"<row>(\d+)</row>", block).group(1))
            pictures.append((row, block))
        pictures.sort()
        serial = pictures[1:]
        if len(serial) != len(paths):
            raise SystemExit(f"{sheet}: {len(serial)} pictures, {len(paths)} new maps")
        drawing = re.sub(
            r'<oneCellAnchor>(?:(?!</oneCellAnchor>).)*prst="ellipse".*?</oneCellAnchor>',
            "",
            drawing,
            flags=re.S,
        )
        next_shape = max((int(i) for i in re.findall(r'<cNvPr id="(\d+)"', drawing)), default=0) + 1
        extras = []
        for (_row, block), png in zip(serial, paths):
            embed = re.search(r'r:embed="([^"]+)"', block).group(1)
            media = rid_media[embed]
            image = Image.open(png).convert("RGB")
            buf = io.BytesIO()
            image.save(buf, "JPEG", quality=74, optimize=True)
            parts[media] = buf.getvalue()
            geom = MAP_CIRCLE_PIXELS[str(Path(png).resolve())]
            img_w, img_h = geom["size"]
            origin = re.search(
                r"<from><col>(\d+)</col><colOff>(\d+)</colOff><row>(\d+)</row><rowOff>(\d+)</rowOff></from><ext cx=\"(\d+)\" cy=\"(\d+)\"",
                block,
            )
            col, col_off, row, row_off, pic_cx, pic_cy = (int(v) for v in origin.groups())
            for circle in geom["circles"]:
                extras.append(_circle_anchor(circle, next_shape, col, col_off, row, row_off, pic_cx, pic_cy, img_w, img_h))
                next_shape += 1
        parts[drawing_name] = drawing.replace("</wsDr>", "".join(extras) + "</wsDr>").encode("utf-8")
        print(f"  {sheet}: replaced {len(paths)} maps, {len(extras)} circles", flush=True)

    note_hits = 0
    for name, body in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        text = body.decode("utf-8")
        replaced = text
        if _PIE_SENTENCE_OLD in replaced:
            note_hits += replaced.count(_PIE_SENTENCE_OLD)
            replaced = replaced.replace(_PIE_SENTENCE_OLD, _PIE_SENTENCE_NEW)
        if _AIM_SENTENCE_OLD in replaced:
            note_hits += replaced.count(_AIM_SENTENCE_OLD)
            replaced = replaced.replace(_AIM_SENTENCE_OLD, _AIM_SENTENCE_NEW)
        if _AIM_SENTENCE_CELL in replaced:
            note_hits += replaced.count(_AIM_SENTENCE_CELL)
            replaced = replaced.replace(_AIM_SENTENCE_CELL, _AIM_SENTENCE_NEW)
        if replaced == text:
            continue
        parts[name] = replaced.encode("utf-8")
    print(f"Updated the pie note in {note_hits} place(s)", flush=True)

    tmp = src.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name])
        for name, body in parts.items():
            if name not in names:
                zf.writestr(name, body)
    tmp.replace(VERSIONED_XLSX)
    spread_circle_anchors(VERSIONED_XLSX)
    if VERSIONED_XLSX.resolve() != REPORT_XLSX.resolve():
        shutil.copy2(VERSIONED_XLSX, REPORT_XLSX)
    print(f"Wrote {VERSIONED_XLSX}", flush=True)
    return VERSIONED_XLSX


def _jpeg_bytes(path: Path) -> bytes:
    import io
    from PIL import Image

    image = Image.open(path).convert("RGB")
    buf = io.BytesIO()
    image.save(buf, "JPEG", quality=74, optimize=True)
    return buf.getvalue()


def _sheet_pictures(parts: dict[str, bytes], sheet_name: str) -> list[tuple[int, int, str]]:
    """Pictures on a sheet as (row, col, media path), top to bottom and left to right."""
    drawings = _sheet_drawings(parts)
    drawing_name = drawings[sheet_name]
    rels_name = drawing_name.replace("drawings/", "drawings/_rels/") + ".rels"
    rels = parts[rels_name].decode("utf-8")
    rid_media = {}
    for rel in re.findall(r"<Relationship\b[^>]*/>", rels):
        rid = re.search(r'Id="([^"]+)"', rel).group(1)
        target = re.search(r'Target="([^"]+)"', rel).group(1)
        if target.startswith("/"):
            media = target.lstrip("/")
        elif target.startswith("../"):
            media = "xl/" + target[3:]
        else:
            media = "xl/media/" + target.rsplit("/", 1)[-1]
        rid_media[rid] = media
    drawing = parts[drawing_name].decode("utf-8")
    pics = []
    for block in re.findall(r"<oneCellAnchor>.*?</oneCellAnchor>", drawing, flags=re.S):
        if "<pic>" not in block:
            continue
        row = int(re.search(r"<row>(\d+)</row>", block).group(1))
        col = int(re.search(r"<col>(\d+)</col>", block).group(1))
        embed = re.search(r'r:embed="([^"]+)"', block).group(1)
        pics.append((row, col, rid_media[embed]))
    pics.sort()
    return pics


def refresh_sheet_maps(src: Path | None = None) -> Path:
    """Redraw coverage maps, combined views and the DT route with a scale bar.

    Serial bad-spot maps are left for refresh-pies. Pictures are swapped in the
    zip so the movable circles stay.
    """
    from merge_and_analyze import ACTIVE_MERGED, MERGED_CSV, load_active_merged, load_merged_csv

    src = src or REPORT_XLSX
    if not src.exists():
        raise SystemExit(f"Missing workbook: {src}")
    print("Loading drive logs", flush=True)
    idle_raw = load_merged_csv() if MERGED_CSV.exists() else None
    active_raw = load_active_merged() if ACTIVE_MERGED.exists() else None
    if idle_raw is None or active_raw is None:
        raise SystemExit("Merged drive logs are missing.")
    idle_raw = _with_layers(idle_raw)
    active_raw = select_active_report(_with_layers(active_raw))
    idle_fixed = strongest_fix(idle_raw)
    idle_track = best_server(idle_fixed)
    active_track = strongest_location(active_raw)
    idle_extent = coverage_extent(idle_track)
    active_extent = coverage_extent(active_track)
    sheet_images: dict[str, list[Path]] = {}
    combined: dict[tuple[str, str], Path] = {}
    for kpi, plot_cov, plot_bad, labels in (
        ("RSRP", plot_rsrp_coverage_map, plot_rsrp_bad_spot_map, MAP_RSRP_LABELS),
        ("RSRQ", plot_rsrq_coverage_map, plot_rsrq_bad_spot_map, MAP_RSRQ_LABELS),
        ("SINR", plot_sinr_coverage_map, plot_sinr_bad_spot_map, MAP_SINR_LABELS),
    ):
        del labels
        ordered: list[Path] = []
        for tag, base, track, extent, mode_label in (
            ("", idle_fixed, idle_track, idle_extent, None),
            ("active", active_raw, active_track, active_extent, "Active"),
        ):
            suffix = f"_{tag}" if tag else ""
            print(f"Drawing {kpi} combined ({tag or 'idle'})", flush=True)
            spots = find_bad_spots(track, kpi)
            mode_name = "Active" if tag == "active" else "Idle"
            title = f"Combined view — {kpi}" + (f" — {mode_label}" if mode_label else "")
            bad_kwargs = {"extent": extent, "title": title, "shapes": False, "outlines": False}
            if kpi != "SINR":
                bad_kwargs["legend_title"] = f"{mode_name} by {kpi}"
            path = plot_bad(
                track, spots, PLOTS_DIR / f"excel_bad_spots_{kpi.lower()}{suffix}.png",
                **bad_kwargs,
            )
            combined[(kpi, tag)] = path
            ordered.append(path)
            band_paths = []
            for layer in REPORT_LAYERS:
                print(f"Drawing {kpi} {layer} ({tag or 'idle'})", flush=True)
                part = filter_by_layer(base, layer)
                if tag != "active":
                    part = best_server(part)
                band_paths.append(plot_cov(
                    part, PLOTS_DIR / f"excel_route_{kpi.lower()}_{layer}{suffix}.png",
                    layer=layer, extent=extent, mode_label=mode_label,
                ))
            if tag == "":
                idle_bands = band_paths
            else:
                for left, right in zip(idle_bands, band_paths):
                    ordered.append(left)
                    ordered.append(right)
        sheet_images[kpi] = ordered
        print(f"  {kpi}: {len(ordered)} maps", flush=True)

    print("Drawing the DT route", flush=True)
    route = plot_road_map(active_track, PLOTS_DIR / "excel_cover_active_route.png")

    print("Updating the workbook", flush=True)
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    for sheet, paths in sheet_images.items():
        pics = _sheet_pictures(parts, sheet)
        if len(pics) != len(paths):
            raise SystemExit(f"{sheet}: {len(pics)} pictures, {len(paths)} new maps")
        for (_row, _col, media), path in zip(pics, paths):
            parts[media] = _jpeg_bytes(path)
        print(f"  {sheet}: replaced {len(paths)} maps", flush=True)
    for sheet, key in (("Bad Spot RSRP", ("RSRP", "")), ("Bad Spot RSRP_Active", ("RSRP", "active"))):
        pics = _sheet_pictures(parts, sheet)
        if not pics:
            raise SystemExit(f"{sheet}: no pictures")
        _row, _col, media = pics[0]
        parts[media] = _jpeg_bytes(combined[key])
        print(f"  {sheet}: replaced the combined view", flush=True)
    cover_pics = _sheet_pictures(parts, "Cover")
    route_pics = [pic for pic in cover_pics if pic[0] >= 70]
    if len(route_pics) != 1:
        raise SystemExit(f"Cover: expected one DT route picture, found {len(cover_pics)}")
    parts[route_pics[0][2]] = _jpeg_bytes(route)
    print("  Cover: replaced the DT route", flush=True)

    tmp = src.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name])
    tmp.replace(VERSIONED_XLSX)
    if VERSIONED_XLSX.resolve() != REPORT_XLSX.resolve():
        shutil.copy2(VERSIONED_XLSX, REPORT_XLSX)
    print(f"Wrote {VERSIONED_XLSX}", flush=True)
    return VERSIONED_XLSX


_ADVICE_NOTE_OLD = (
    "The columns to the right of each map are a suggestion table. Azimuth and tilt are used when a sector is within about 2.5 km, and downtilt is limited to 2°. A larger downtilt becomes a 4th L900 sector on a site within 1 km. A 4th sector also replaces a new site when a host inside 4 km can aim at the bad spot. A new 3-sector site is kept only when the nearest site is farther than that. A blank with no plot inside the circle is counted with that bad spot.",
    "The space to the right of each map is that suggestion table. Site, Population and Final solution are filled for every bad spot in the table below. A blank with no plot inside the circle is counted with that bad spot.",
)
_ADVICE_NOTE_NEW = (
    "The columns to the right of each map are a suggestion table. Azimuth and tilt are used when a sector is within about 2.5 km, and downtilt is limited to 2°. A larger downtilt becomes a 4th L900 sector on a site within 1 km. A 4th sector also replaces a new site when a host inside 3 km can aim at the bad spot. A new 3-sector site is kept only when the nearest site is farther than that. A blank with no plot inside the circle is counted with that bad spot.",
    "The space to the right of each map is that suggestion table. Site, Population and Final solution are filled for every bad spot in the table below. A blank with no plot inside the circle is counted with that bad spot.",
)
_DROP_BAD_SPOT_SHEETS = (
    "Bad Spot RSRQ",
    "Bad Spot SINR",
    "Bad Spot RSRQ_Active",
    "Bad Spot SINR_Active",
)


def _xml_text(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _advice_sheet_rows(header: int, plan: dict) -> tuple[dict[int, str], list[str]]:
    """Inline cells for the suggestion table beside one bad-spot map."""
    title = header + 2
    find = header + 4
    head = header + 10

    def cell(ref: str, style: int, text: str | None = None) -> str:
        if text is None:
            return f'<c r="{ref}" s="{style}"/>'
        return (
            f'<c r="{ref}" s="{style}" t="inlineStr"><is>'
            f'<t xml:space="preserve">{_xml_text(text)}</t></is></c>'
        )

    def styled_row(row: int, cells: list[str]) -> str:
        return f'<row r="{row}">{"".join(cells)}</row>'

    rows = {
        title: styled_row(title, [cell(f"I{title}", 32, "Suggestions and solutions")] + [
            cell(f"{get_column_letter(col)}{title}", 32) for col in range(10, 17)
        ]),
        find: styled_row(find, [
            cell(f"I{find}", 36, "Findings"),
            cell(f"J{find}", 37, plan.get("finding", "")),
        ]),
        head: styled_row(head, [
            cell(f"I{head}", 35, "Step"),
            cell(f"J{head}", 35, "Action"),
            cell(f"K{head}", 35),
            cell(f"L{head}", 35, "Proposal"),
            cell(f"M{head}", 35),
            cell(f"N{head}", 35),
            cell(f"O{head}", 35),
            cell(f"P{head}", 35),
        ]),
    }
    merges = [
        f"I{title}:P{title}",
        f"I{find}:I{find + 4}",
        f"J{find}:P{find + 4}",
        f"J{head}:K{head}",
        f"L{head}:P{head}",
    ]
    for index, step in enumerate(plan.get("steps") or []):
        row = head + 1 + index * 4
        rows[row] = styled_row(row, [
            cell(f"I{row}", 36, str(step.get("step", ""))),
            cell(f"J{row}", 36, str(step.get("action", ""))),
            cell(f"L{row}", 37, str(step.get("proposal", ""))),
        ])
        merges.extend((
            f"I{row}:I{row + 3}",
            f"J{row}:K{row + 3}",
            f"L{row}:P{row + 3}",
        ))
    return rows, merges


def _strip_advice_rows(xml: str, headers: list[tuple[int, int]]) -> str:
    """Remove a previous suggestion block so the new table can take its rows."""
    drop = {row for header, _number in headers for row in range(header + 2, header + 32)}

    def keep_row(match: re.Match) -> str:
        number = int(re.search(r'\br="(\d+)"', match.group(0)).group(1))
        return "" if number in drop else match.group(0)

    xml = re.sub(r"<row\b.*?</row>", keep_row, xml, flags=re.S)

    def keep_merge(match: re.Match) -> str:
        rows = [int(value) for value in re.findall(r"\d+", match.group(0))]
        return "" if any(row in drop for row in rows) else match.group(0)

    xml = re.sub(r"<mergeCell\b[^>]*/>", keep_merge, xml)
    count = len(re.findall(r"<mergeCell\b", xml))
    xml = re.sub(r'<mergeCells count="\d+"', f'<mergeCells count="{count}"', xml, count=1)
    return xml


def _insert_sheet_rows(xml: str, rows: dict[int, str], merges: list[str]) -> str:
    match = re.search(r"(<sheetData>)(.*?)(</sheetData>)", xml, flags=re.S)
    if match is None:
        raise SystemExit("Worksheet has no sheetData")
    existing = re.findall(r"<row\b.*?</row>", match.group(2), flags=re.S)
    by_row: dict[int, str] = {}
    for row in existing:
        number = int(re.search(r'\br="(\d+)"', row).group(1))
        by_row[number] = row
    for number, row in rows.items():
        if number in by_row:
            raise SystemExit(f"Row {number} already exists; the suggestions table would overwrite it")
        by_row[number] = row
    body = "".join(by_row[number] for number in sorted(by_row))
    xml = xml[: match.start(2)] + body + xml[match.end(2) :]
    merge = re.search(r'<mergeCells count="(\d+)">', xml)
    if merge is None:
        raise SystemExit("Worksheet has no mergeCells")
    count = int(merge.group(1)) + len(merges)
    extra = "".join(f'<mergeCell ref="{ref}"/>' for ref in merges)
    xml = xml.replace(merge.group(0), f'<mergeCells count="{count}">', 1)
    xml = xml.replace("</mergeCells>", extra + "</mergeCells>", 1)
    return xml


def _add_advice_styles(styles: str) -> str:
    if "<cellXfs count=\"38\">" in styles:
        return styles
    if "<cellXfs count=\"35\">" not in styles:
        if "<fonts count=\"8\">" not in styles or "<cellXfs count=\"32\">" not in styles:
            raise SystemExit("Workbook styles are not the expected set; refusing to append advice styles")
        styles = styles.replace("<fonts count=\"8\">", "<fonts count=\"9\">", 1)
        styles = styles.replace(
            "</fonts>",
            '<font><name val="Calibri"/><b val="1"/><color rgb="001B4F72"/><sz val="11"/></font></fonts>',
            1,
        )
        styles = styles.replace("<cellXfs count=\"32\">", "<cellXfs count=\"35\">", 1)
        styles = styles.replace(
            "</cellXfs>",
            '<xf numFmtId="0" fontId="8" fillId="5" borderId="0" applyAlignment="1" xfId="0">'
            '<alignment horizontal="center" vertical="center"/></xf>'
            '<xf numFmtId="0" fontId="8" fillId="0" borderId="0" applyAlignment="1" xfId="0">'
            '<alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
            '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" applyAlignment="1" xfId="0">'
            '<alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
            "</cellXfs>",
            1,
        )
    styles = styles.replace("<cellXfs count=\"35\">", "<cellXfs count=\"38\">", 1)
    added = (
        '<xf numFmtId="0" fontId="8" fillId="5" borderId="1" applyAlignment="1" xfId="0">'
        '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="8" fillId="0" borderId="1" applyAlignment="1" xfId="0">'
        '<alignment horizontal="center" vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="6" fillId="0" borderId="1" applyAlignment="1" xfId="0">'
        '<alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
    )
    return styles.replace("</cellXfs>", added + "</cellXfs>", 1)


def _inline_cell(ref: str, style: int, text: str) -> str:
    return (
        f'<c r="{ref}" s="{style}" t="inlineStr"><is>'
        f'<t xml:space="preserve">{_xml_text(text)}</t></is></c>'
    )


def _add_solution_columns(xml: str, plans: list[dict]) -> str:
    """Add Site, Population and Final solution to the right of the bad-spot table."""
    header = re.search(
        r'<row r="(\d+)"[^>]*>(?:(?!</row>).)*<t>Top Cell Id</t>(?:(?!</row>).)*</row>',
        xml,
        flags=re.S,
    )
    if header is None:
        raise SystemExit("Bad-spot table header was not found")
    header_row = header.group(1)
    if "Final solution" not in header.group(0):
        extra = (
            _inline_cell(f"O{header_row}", 7, "Site")
            + _inline_cell(f"P{header_row}", 7, "Population")
            + _inline_cell(f"Q{header_row}", 7, "Final solution")
        )
        xml = xml.replace(header.group(0), header.group(0).replace("</row>", extra + "</row>"), 1)
    found = []
    for match in re.finditer(r'<row r="(\d+)"([^>]*)>(.*?)</row>', xml, flags=re.S):
        row_n = match.group(1)
        if re.search(rf'<c r="A{row_n}"[^>]*>\s*<is><t>Circle \d+</t>', match.group(0)):
            found.append(match)
    if len(found) != len(plans):
        raise SystemExit(f"Bad-spot table has {len(found)} rows and {len(plans)} solutions")
    pieces = []
    last = 0
    for match, plan in zip(found, plans):
        row_n = match.group(1)
        attrs = match.group(2)
        body = match.group(3)
        pieces.append(xml[last:match.start()])
        updates = (
            (f"O{row_n}", 20, str(plan.get("site", ""))),
            (f"P{row_n}", 20, str(plan.get("population", ""))),
            (f"Q{row_n}", 37, str(plan.get("final", ""))),
        )
        if f'r="Q{row_n}"' not in body:
            if "ht=" not in attrs:
                attrs += ' ht="52" customHeight="1"'
            body += "".join(_inline_cell(ref, style, text) for ref, style, text in updates)
        else:
            for ref, style, text in updates:
                body = re.sub(
                    rf'<c r="{ref}"[^>]*>.*?</c>',
                    _inline_cell(ref, style, text),
                    body,
                    count=1,
                    flags=re.S,
                )
        pieces.append(f'<row r="{row_n}"{attrs}>{body}</row>')
        last = match.end()
    pieces.append(xml[last:])
    xml = "".join(pieces)
    xml = re.sub(
        r'<col width="13" customWidth="1" min="15" max="15" />',
        '<col width="18" customWidth="1" min="15" max="15" />',
        xml,
        count=1,
    )
    xml = re.sub(
        r'<col width="13" customWidth="1" min="16" max="16" />',
        '<col width="16" customWidth="1" min="16" max="16" />',
        xml,
        count=1,
    )
    if 'min="17"' not in xml:
        xml = xml.replace("</cols>", '<col width="64" customWidth="1" min="17" max="17" /></cols>', 1)
    xml = re.sub(r'(<dimension ref="A1:)[A-Z]+(\d+")', r"\1Q\2", xml, count=1)
    marker = "Site, Population and Final solution on the right are filled for every bad spot in this table."
    if marker not in xml and "Map numbers are omitted." in xml:
        xml = xml.replace("Map numbers are omitted.", "Map numbers are omitted. " + marker, 1)
    return xml


def _shift_rows_down(xml: str, start: int, count: int) -> str:
    """Move every row at or below start down by count. Drawings are left as they are."""
    def shift_r(match: re.Match) -> str:
        row = int(match.group(2))
        if row >= start:
            row += count
        return f'r="{match.group(1)}{row}"'

    def shift_ref(match: re.Match) -> str:
        r1 = int(match.group(2))
        r2 = int(match.group(4))
        if r1 >= start:
            r1 += count
        if r2 >= start:
            r2 += count
        return f'ref="{match.group(1)}{r1}:{match.group(3)}{r2}"'

    xml = re.sub(r'\br="([A-Z]*)(\d+)"', shift_r, xml)
    return re.sub(r'ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', shift_ref, xml)


def _replace_summary_counts(xml: str, counts: dict[str, int], total: int) -> str:
    """Refresh the four count cells when the summary table is already on the sheet."""
    values = {
        "Tilt and azimuth": f"{counts['tune']:,}",
        "4th sector, L900": f"{counts['fourth']:,}",
        "New 3-sector site": f"{counts['new']:,}",
        "All bad spots": f"{total:,}",
    }

    def repl_row(match: re.Match) -> str:
        row = match.group(0)
        for label, value in values.items():
            if f">{label}<" not in row:
                continue
            return re.sub(
                r'(<c r="B\d+"[^>]*>.*?<t[^>]*>)([^<]*)(</t>)',
                lambda found: found.group(1) + value + found.group(3),
                row,
                count=1,
                flags=re.S,
            )
        return row

    return re.sub(r"<row r=\"\d+\"[^>]*>(?:(?!</row>).)*</row>", repl_row, xml)


def _insert_solution_summary(xml: str, plans: list[dict]) -> str:
    """Put the solution counts between the last map and the identified-spots table."""
    counts = {"tune": 0, "fourth": 0, "new": 0}
    for plan in plans:
        kind = str(plan.get("kind") or "")
        if kind in counts:
            counts[kind] += 1
    if "Summary of proposed solutions" in xml:
        return _replace_summary_counts(xml, counts, len(plans))
    found = re.search(
        r'<row r="(\d+)"[^>]*>(?:(?!</row>).)*Identified [^<]*bad spots(?:(?!</row>).)*</row>',
        xml,
        flags=re.S,
    )
    if found is None:
        raise SystemExit("Identified bad-spot title was not found")
    start = int(found.group(1))
    xml = _shift_rows_down(xml, start, 8)
    lines = [
        (start, "Summary of proposed solutions", None, 6),
        (start + 2, "Solution", "Count", 7),
        (start + 3, "Tilt and azimuth", f"{counts['tune']:,}", 20),
        (start + 4, "4th sector, L900", f"{counts['fourth']:,}", 20),
        (start + 5, "New 3-sector site", f"{counts['new']:,}", 20),
        (start + 6, "All bad spots", f"{len(plans):,}", 7),
    ]
    rows = {}
    for row, left, right, style in lines:
        cells = [_inline_cell(f"A{row}", style, left)]
        if right is not None:
            cells.append(_inline_cell(f"B{row}", style, right))
        rows[row] = f'<row r="{row}" ht="20" customHeight="1">{"".join(cells)}</row>'
    return _insert_sheet_rows(xml, rows, [])


def _drop_workbook_sheets(parts: dict[str, bytes], sheet_names: tuple[str, ...]) -> None:
    """Remove named sheets, their drawings, and pictures nothing else uses."""
    wb_xml = parts["xl/workbook.xml"].decode("utf-8")
    rels_xml = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    worksheets = []
    for name in sheet_names:
        match = re.search(
            rf'<sheet name="{re.escape(name)}" sheetId="\d+" state="[^"]+" r:id="([^"]+)" />',
            wb_xml,
        )
        if match is None:
            continue
        rid = match.group(1)
        wb_xml = wb_xml.replace(match.group(0), "")
        rel = re.search(rf'<Relationship\b[^>]*Id="{rid}"[^>]*/>', rels_xml)
        if rel is None:
            raise SystemExit(f"Workbook relationship {rid} for {name} was not found")
        target = re.search(r'Target="([^"]+)"', rel.group(0)).group(1).lstrip("/")
        rels_xml = rels_xml.replace(rel.group(0), "")
        worksheets.append(target)
    parts["xl/workbook.xml"] = wb_xml.encode("utf-8")
    parts["xl/_rels/workbook.xml.rels"] = rels_xml.encode("utf-8")
    if not worksheets:
        return
    drop = set(worksheets)
    for sheet in worksheets:
        rels_name = sheet.replace("worksheets/", "worksheets/_rels/") + ".rels"
        rels = parts.get(rels_name, b"").decode("utf-8")
        for target in re.findall(r'Target="([^"]+)"', rels):
            path = target.lstrip("/")
            if path.startswith("xl/"):
                drop.add(path)
                drawing_rels = path.replace("drawings/", "drawings/_rels/") + ".rels"
                drawing = parts.get(drawing_rels, b"").decode("utf-8")
                for media in re.findall(r'Target="([^"]+)"', drawing):
                    drop.add(media.lstrip("/"))
                drop.add(drawing_rels)
        drop.add(rels_name)
    kept_text = []
    for name, body in parts.items():
        if name in drop or not name.endswith(".rels"):
            continue
        kept_text.append(body.decode("utf-8"))
    still_used = "\n".join(kept_text)
    for path in list(drop):
        if path.startswith("xl/media/") and path in still_used:
            drop.remove(path)
    content = parts["[Content_Types].xml"].decode("utf-8")
    for path in drop:
        content = re.sub(rf'<Override PartName="/{re.escape(path)}"[^>]*/>', "", content)
        parts.pop(path, None)
    parts["[Content_Types].xml"] = content.encode("utf-8")


_COVER_SUMMARY_ROW = 110
_EP_SHEET = "xl/worksheets/sheet15.xml"
_EP_RID = "rId17"
_SOLUTION_NAME = {
    "tune": "Tilt and azimuth",
    "fourth": "4th sector, L900",
    "new": "New 3-sector site",
}


def _drop_cover_summary(xml: str) -> str:
    """Remove a previous BAD Spot Summary so a second advise does not stack it."""
    if "BAD Spot Summary" not in xml:
        return xml

    def keep_row(match: re.Match) -> str:
        return "" if int(match.group(1)) >= _COVER_SUMMARY_ROW else match.group(0)

    def keep_merge(match: re.Match) -> str:
        rows = [int(number) for number in re.findall(r"\d+", match.group(1))]
        return "" if rows and max(rows) >= _COVER_SUMMARY_ROW else match.group(0)

    xml = re.sub(r'<row r="(\d+)"[^>]*>.*?</row>', keep_row, xml, flags=re.S)
    xml = re.sub(r'<mergeCell ref="([^"]+)" />', keep_merge, xml)
    count = len(re.findall(r"<mergeCell ", xml))
    return re.sub(r'<mergeCells count="\d+">', f'<mergeCells count="{count}">', xml, count=1)


def _graft_cover_summary(xml: str, idle_plans: list[dict], active_plans: list[dict]) -> str:
    """Idle and Active bad-spot counts and the final solution for every spot."""
    xml = _drop_cover_summary(xml)
    rows: list[str] = []
    merges: list[str] = []
    row = _COVER_SUMMARY_ROW
    bar = [_inline_cell(f"A{row}", 5, "BAD Spot Summary")]
    for col in range(2, 21):
        bar.append(f'<c r="{get_column_letter(col)}{row}" s="5" t="n" />')
    rows.append(f'<row r="{row}" ht="22" customHeight="1">{"".join(bar)}</row>')
    merges.append(f"A{row}:T{row}")
    row += 2

    def add_mode(title: str, plans: list[dict]) -> None:
        nonlocal row
        rows.append(
            f'<row r="{row}" ht="20" customHeight="1">{_inline_cell(f"A{row}", 6, title)}</row>'
        )
        row += 1
        counts = {"tune": 0, "fourth": 0, "new": 0}
        for plan in plans:
            kind = str(plan.get("kind") or "")
            if kind in counts:
                counts[kind] += 1
        rows.append(
            f'<row r="{row}" ht="18" customHeight="1">'
            f'{_inline_cell(f"A{row}", 7, "Solution")}{_inline_cell(f"B{row}", 7, "Count")}</row>'
        )
        row += 1
        for label, key in (
            ("Tilt and azimuth", "tune"),
            ("4th sector, L900", "fourth"),
            ("New 3-sector site", "new"),
            ("All bad spots", None),
        ):
            value = str(len(plans) if key is None else counts[key])
            style = 7 if key is None else 20
            rows.append(
                f'<row r="{row}" ht="18" customHeight="1">'
                f'{_inline_cell(f"A{row}", style, label)}{_inline_cell(f"B{row}", style, value)}</row>'
            )
            row += 1
        row += 1
        rows.append(
            f'<row r="{row}" ht="18" customHeight="1">'
            f'{_inline_cell(f"A{row}", 7, "Bad spot")}'
            f'{_inline_cell(f"B{row}", 7, "Site")}'
            f'{_inline_cell(f"C{row}", 7, "Population")}'
            f'{_inline_cell(f"D{row}", 7, "Solution")}'
            f'{_inline_cell(f"E{row}", 7, "Final solution")}</row>'
        )
        merges.append(f"E{row}:T{row}")
        row += 1
        for index, plan in enumerate(plans, 1):
            rows.append(
                f'<row r="{row}" ht="48" customHeight="1">'
                f'{_inline_cell(f"A{row}", 20, f"Bad spot {index}")}'
                f'{_inline_cell(f"B{row}", 20, str(plan.get("site") or ""))}'
                f'{_inline_cell(f"C{row}", 20, str(plan.get("population") or ""))}'
                f'{_inline_cell(f"D{row}", 20, _SOLUTION_NAME.get(str(plan.get("kind") or ""), ""))}'
                f'{_inline_cell(f"E{row}", 37, str(plan.get("final") or ""))}</row>'
            )
            merges.append(f"E{row}:T{row}")
            row += 1
        row += 2

    add_mode("Idle", idle_plans)
    add_mode("Active", active_plans)
    last = row - 1
    xml = xml.replace("</sheetData>", "".join(rows) + "</sheetData>", 1)
    xml = re.sub(
        r'(<mergeCells count=")(\d+)(">)',
        lambda match: match.group(1) + str(int(match.group(2)) + len(merges)) + match.group(3),
        xml,
        count=1,
    )
    xml = xml.replace("</mergeCells>", "".join(f'<mergeCell ref="{ref}"/>' for ref in merges) + "</mergeCells>", 1)
    xml = re.sub(r'<dimension ref="A1:[A-Z]+\d+"', f'<dimension ref="A1:T{last}"', xml, count=1)
    return xml


def _ep_text(value: object) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, float):
        if abs(value - round(value)) < 0.05:
            return str(int(round(value)))
        return f"{value:.1f}"
    return str(value)


def _ep_rows(mode: str, plans: list[dict]) -> list[dict]:
    rows = []
    for index, plan in enumerate(plans, 1):
        for action in plan.get("actions") or []:
            item = dict(action)
            item["mode"] = mode
            item["spot"] = f"Bad spot {index}"
            rows.append(item)
    return rows


def _ep_sheet_xml(actions: list[dict]) -> bytes:
    """Sector and cell engineering-parameter tracker. Existing plan is blank for a new sector."""
    rest = "".join(f'<c r="{col}{{0}}" s="{{1}}" t="n" />' for col in "BCDEFGHIJKLMN")
    rows = [
        '<row r="1" ht="24" customHeight="1">'
        + _inline_cell("A1", 6, "EP Details")
        + rest.format(1, 6)
        + "</row>",
        '<row r="2" ht="32" customHeight="1">'
        + _inline_cell(
            "A2",
            16,
            "Engineering Parameter action tracker. One row is one sector or cell change for one bad spot. "
            "Existing Plan is blank when the sector is new. Height is metres. Azimuth, ET and MT are degrees.",
        )
        + rest.format(2, 16)
        + "</row>",
    ]
    top = {
        "A": "Site",
        "B": "Sector",
        "C": "Serving cell name",
        "D": "Existing Plan",
        "H": "Proposed Plan",
        "L": "Mode",
        "M": "Bad spot",
        "N": "Action",
    }
    sub = {
        "D": "Height",
        "E": "Azimuth",
        "F": "ET",
        "G": "MT",
        "H": "Height",
        "I": "Azimuth",
        "J": "ET",
        "K": "MT",
    }
    row3 = "".join(_inline_cell(f"{col}3", 7, top.get(col, "")) for col in "ABCDEFGHIJKLMN")
    row4 = "".join(
        _inline_cell(f"{col}4", 7, sub[col]) if col in sub else f'<c r="{col}4" s="7" t="n" />'
        for col in "ABCDEFGHIJKLMN"
    )
    rows.append(f'<row r="3" ht="20" customHeight="1">{row3}</row>')
    rows.append(f'<row r="4" ht="20" customHeight="1">{row4}</row>')
    fields = (
        "site", "sector", "cell",
        "height", "azimuth", "et", "mt",
        "p_height", "p_azimuth", "p_et", "p_mt",
        "mode", "spot", "action",
    )
    for offset, action in enumerate(actions):
        number = 5 + offset
        cells = []
        for col, field in zip("ABCDEFGHIJKLMN", fields):
            style = 20 if col in "ABCLMN" else 36
            if col == "C":
                style = 37
            cells.append(_inline_cell(f"{col}{number}", style, _ep_text(action.get(field))))
        rows.append(f'<row r="{number}" ht="22" customHeight="1">{"".join(cells)}</row>')
    last = 4 + max(len(actions), 1)
    merges = (
        "A1:N1", "A2:N2",
        "A3:A4", "B3:B4", "C3:C4",
        "D3:G3", "H3:K3",
        "L3:L4", "M3:M4", "N3:N4",
    )
    cols = (
        '<col width="28" customWidth="1" min="1" max="1" />'
        '<col width="16" customWidth="1" min="2" max="2" />'
        '<col width="28" customWidth="1" min="3" max="3" />'
        + "".join(
            f'<col width="12" customWidth="1" min="{index}" max="{index}" />'
            for index in range(4, 12)
        )
        + '<col width="12" customWidth="1" min="12" max="12" />'
        '<col width="16" customWidth="1" min="13" max="13" />'
        '<col width="22" customWidth="1" min="14" max="14" />'
    )
    xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr fitToPage="1" /></sheetPr>'
        f'<dimension ref="A1:N{last}" />'
        '<sheetViews><sheetView showGridLines="0" view="normal" zoomScale="100" workbookViewId="0">'
        '<selection activeCell="A1" sqref="A1" /></sheetView></sheetViews>'
        '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />'
        f"<cols>{cols}</cols>"
        f'<sheetData>{"".join(rows)}</sheetData>'
        f'<mergeCells count="{len(merges)}">'
        + "".join(f'<mergeCell ref="{ref}"/>' for ref in merges)
        + "</mergeCells>"
        '<printOptions gridLines="0" />'
        '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3" />'
        '<pageSetup orientation="landscape" paperSize="9" fitToHeight="0" fitToWidth="1" />'
        "</worksheet>"
    )
    return xml.encode("utf-8")


def _ensure_ep_sheet(parts: dict[str, bytes]) -> None:
    workbook = parts["xl/workbook.xml"].decode("utf-8")
    if 'name="EP Details"' not in workbook:
        workbook = workbook.replace(
            '<sheet name="_ChartData"',
            f'<sheet name="EP Details" sheetId="15" state="visible" r:id="{_EP_RID}" />'
            '<sheet name="_ChartData"',
            1,
        )
        parts["xl/workbook.xml"] = workbook.encode("utf-8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if f'Id="{_EP_RID}"' not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="/xl/worksheets/sheet15.xml" Id="{_EP_RID}" /></Relationships>',
            1,
        )
        parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    content = parts["[Content_Types].xml"].decode("utf-8")
    if "/xl/worksheets/sheet15.xml" not in content:
        content = content.replace(
            "</Types>",
            '<Override PartName="/xl/worksheets/sheet15.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
            1,
        )
        parts["[Content_Types].xml"] = content.encode("utf-8")


def write_bad_spot_advice(src: Path | None = None) -> Path:
    """Graft the findings table onto the six bad-spot sheets without rewriting the file.

    openpyxl save drops the movable circles, so the cells are inserted into the
    worksheet XML. Row heights stay at the default, and the drawings are copied
    through unchanged.
    """
    from merge_and_analyze import ACTIVE_MERGED, MERGED_CSV, load_active_merged, load_merged_csv

    src = src or REPORT_XLSX
    if not src.exists():
        raise SystemExit(f"Missing workbook: {src}")
    print("Loading drive logs", flush=True)
    idle = load_merged_csv() if MERGED_CSV.exists() else None
    active = load_active_merged() if ACTIVE_MERGED.exists() else None
    if idle is None or active is None:
        raise SystemExit("Merged drive logs are missing; a full report build is required.")
    active = select_active_report(active)
    idle_track = best_server(idle)
    active_track = best_carrier(active)
    tracks = {
        "Bad Spot RSRP": (idle_track, "RSRP"),
        "Bad Spot RSRP_Active": (active_track, "RSRP"),
    }
    advice_for: dict[str, list[dict]] = {}
    for sheet, (frame, kpi) in tracks.items():
        print(f"Advice for {sheet}", flush=True)
        spots = find_bad_spots(frame, kpi)
        if len(spots) > BAD_SPOT_TABLE_MAX:
            spots = spots.head(BAD_SPOT_TABLE_MAX)
        advice_for[sheet] = [bad_spot_advice(frame, kpi, spots.iloc[i]) for i in range(len(spots))]
        kinds: dict[str, int] = {}
        for plan in advice_for[sheet]:
            kinds[plan["kind"]] = kinds.get(plan["kind"], 0) + 1
        print(f"  {len(advice_for[sheet])} spots {kinds}", flush=True)

    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    drawings_before = {
        name: parts[name] for name in names if name.startswith("xl/drawings/")
    }
    sheet_xml = _worksheet_xml(parts)
    for sheet, pairs in advice_for.items():
        xml = sheet_xml[sheet].decode("utf-8")
        headers = [
            (int(row), int(number))
            for row, number in re.findall(
                r'<c r="A(\d+)"[^>]*>\s*<is><t>Bad Spot (\d+)</t>',
                xml,
            )
        ]
        map_plans = pairs[:BAD_SPOT_MAX_ZOOMS]
        if len(headers) != len(map_plans):
            raise SystemExit(f"{sheet}: {len(headers)} headers, {len(map_plans)} advice rows")
        rows: dict[int, str] = {}
        merges: list[str] = []
        xml = _strip_advice_rows(xml, headers)
        for (header, number), plan in zip(headers, map_plans):
            if number < 1:
                raise SystemExit(f"{sheet}: unexpected header {number}")
            added_rows, added_merges = _advice_sheet_rows(header, plan)
            rows.update(added_rows)
            merges.extend(added_merges)
        xml = _insert_sheet_rows(xml, rows, merges)
        xml = _add_solution_columns(xml, pairs)
        xml = _insert_solution_summary(xml, pairs)
        for old, new in zip(_ADVICE_NOTE_OLD, _ADVICE_NOTE_NEW):
            if old not in xml:
                if new not in xml:
                    raise SystemExit(f"{sheet}: missing note sentence {old!r}")
                continue
            xml = xml.replace(old, new)
        sheet_xml[sheet] = xml.encode("utf-8")
    # write the updated sheet bytes back onto their zip paths
    main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    wb = ET.fromstring(parts["xl/workbook.xml"])
    book_rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rid_target = {rel.attrib["Id"]: rel.attrib["Target"].lstrip("/") for rel in book_rels}
    cover_path = None
    for node in wb.find(f"{main}sheets"):
        name = node.attrib["name"]
        path = rid_target[node.attrib[f"{rel_ns}id"]]
        if name == "Cover":
            cover_path = path
        if name in sheet_xml and isinstance(sheet_xml[name], (bytes, bytearray)):
            parts[path] = sheet_xml[name]
    if cover_path is None:
        raise SystemExit("Cover sheet was not found")
    parts[cover_path] = _graft_cover_summary(
        parts[cover_path].decode("utf-8"),
        advice_for["Bad Spot RSRP"],
        advice_for["Bad Spot RSRP_Active"],
    ).encode("utf-8")
    parts[_EP_SHEET] = _ep_sheet_xml(
        _ep_rows("Idle", advice_for["Bad Spot RSRP"])
        + _ep_rows("Active", advice_for["Bad Spot RSRP_Active"])
    )
    _ensure_ep_sheet(parts)
    parts["xl/styles.xml"] = _add_advice_styles(parts["xl/styles.xml"].decode("utf-8")).encode("utf-8")
    for name, body in drawings_before.items():
        if parts.get(name) != body:
            raise SystemExit(f"Drawing changed while adding advice: {name}")
    _drop_workbook_sheets(parts, _DROP_BAD_SPOT_SHEETS)
    tmp = src.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            if name in parts:
                zf.writestr(name, parts[name])
        for name, body in parts.items():
            if name not in names:
                zf.writestr(name, body)
    tmp.replace(VERSIONED_XLSX)
    if VERSIONED_XLSX.resolve() != REPORT_XLSX.resolve():
        shutil.copy2(VERSIONED_XLSX, REPORT_XLSX)
    print(f"Wrote {VERSIONED_XLSX}", flush=True)
    return VERSIONED_XLSX


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "refresh-circles":
        raise SystemExit(
            "The bad-spot sheets are a vertical series of maps. Run this script with no arguments "
            "so the workbook is rebuilt."
        )
    if len(sys.argv) > 1 and sys.argv[1] == "refresh-cover":
        raise SystemExit(
            "Do not openpyxl-save this workbook. That rewrite drops the movable bad-spot circles. "
            "A full rebuild (no arguments) writes the Cover histograms and DT route with every sheet. "
            "A Cover-only update has to be grafted into the xlsx zip."
        )
    if len(sys.argv) > 1 and sys.argv[1] == "advise":
        path = write_bad_spot_advice()
        result = verify_report(path)
        print(result["problems"], result.get("circle_shapes"))
        if result["problems"]:
            raise SystemExit("Excel view checks failed: " + "; ".join(result["problems"]))
        print("Verified: findings table added and the movable circles are still present.")
        raise SystemExit(0)
    if len(sys.argv) > 1 and sys.argv[1] == "refresh-maps":
        path = refresh_sheet_maps()
        result = verify_report(path)
        print(result["problems"], result.get("circle_shapes"))
        if result["problems"]:
            raise SystemExit("Excel view checks failed: " + "; ".join(result["problems"]))
        print("Verified: scale bars on the sheet maps, and the movable circles are still present.")
        raise SystemExit(0)
    if len(sys.argv) > 1 and sys.argv[1] == "refresh-pies":
        path = refresh_bad_spot_pies()
        result = verify_report(path)
        print(result["problems"], result.get("circle_shapes"))
        if result["problems"]:
            raise SystemExit("Excel view checks failed: " + "; ".join(result["problems"]))
        print("Verified: transparent sector outlines on the bad-spot maps.")
        raise SystemExit(0)
    from merge_and_analyze import (
        ACTIVE_ARCHIVE,
        ACTIVE_MERGED,
        MERGED_CSV,
        extract_archives,
        load_active_merged,
        load_merged_csv,
        merge_active,
        merge_parts,
    )

    if MERGED_CSV.exists():
        frame = load_merged_csv()
    else:
        extract_archives()
        frame = merge_parts()
    source_newer = (
        ACTIVE_ARCHIVE.exists()
        and ACTIVE_MERGED.exists()
        and ACTIVE_ARCHIVE.stat().st_mtime >= ACTIVE_MERGED.stat().st_mtime
    )
    if ACTIVE_ARCHIVE.exists() and (not ACTIVE_MERGED.exists() or source_newer):
        active_frame = merge_active()
    elif ACTIVE_MERGED.exists():
        active_frame = load_active_merged()
    else:
        active_frame = None
    path = build_report(frame, active_frame)
    result = verify_report(path)
    print(result)
    if result["problems"]:
        raise SystemExit("Excel view checks failed: " + "; ".join(result["problems"]))
    print("Verified: no freeze panes, no gridlines, Idle and Active sheets present.")
