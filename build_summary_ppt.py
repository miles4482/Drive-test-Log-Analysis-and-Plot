#!/usr/bin/env python3
"""Bogura summary deck: cover, KPI maps, bad-spot counts, and examples of each solution.

Proposed azimuth, E_tilt and M_tilt on these slides are the nearest whole degree.
Half a degree rounds up. A later workbook publish must use that same rounding on
the proposed value. The value stored in the dump stays as it was measured.
"""
from __future__ import annotations

import json
import math
import re
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
REPORT = OUTPUT / "Bogura_DriveTest_Report_v1.67.xlsx"
PLOTS = OUTPUT / "plots"
DEST = OUTPUT / "Bogura_DriveTest_Summary_v1.68.pptx"
CIRCLES = PLOTS / "bad_spot_circles.json"

NAVY = RGBColor(0x1B, 0x4F, 0x72)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5D, 0x6D, 0x7E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAPER = RGBColor(0xF7, 0xF9, 0xFB)
LINE = RGBColor(0xD5, 0xDD, 0xE4)
GOLD = RGBColor(0xB7, 0x86, 0x0B)
ORANGE = RGBColor(0xD3, 0x54, 0x00)
PURPLE = RGBColor(0x6C, 0x34, 0x83)


GREEN = RGBColor(0x1A, 0x7A, 0x4C)

SOLUTION_COLOR = {
    "Azimuth": GOLD,
    "Tilt": GREEN,
    "4th sector, L900": ORANGE,
    "New 3-sector site": PURPLE,
}


def _whole_degrees(text: str) -> str:
    """Nearest whole degree. Half a degree rounds up. A no-longer-visible change stays."""
    def repl(match: re.Match) -> str:
        whole = int(math.floor(float(match.group(1)) + 0.5))
        return f"{whole}°"

    text = re.sub(r"(\d+\.\d+)\s*°", repl, text)
    return re.sub(r"(E_tilt|M_tilt|azimuth) (\d+)° to \2°", r"\1 stays \2°", text)


# Idle has 4 azimuth turns and 3 new sites, so those two groups show every spot.
# Every other group shows 5. An azimuth example may also change tilt. A tilt
# example leaves the azimuth where it is.
GROUPS = [
    {
        "solution": "Azimuth",
        "mode": "Active",
        "spots": 25,
        "items": [
            {"n": 97, "why": "The tilt already fits 0.6 km. GBSDR37B turns from 140° to 132°. The dashed line leaves GBSDR37A, the cell that served the red samples."},
            {"n": 60, "why": "Tilt stays. GBSND16A turns from 20° to 7°, and GBSND21B turns from 240° to 216°."},
            {"n": 23, "why": "BOKHL09B turns from 110° to 135°. E_tilt goes from 2° to 1°."},
            {"n": 43, "why": "SGTRS13C turns from 255° to 275°. E_tilt goes from 4° to 1°."},
            {"n": 70, "why": "BOSDRC1B turns from 215° to 229°. E_tilt goes from 3° to 1°."},
        ],
    },
    {
        "solution": "Tilt",
        "mode": "Active",
        "spots": 17,
        "items": [
            {"n": 16, "why": "PBCTM14C is uptilted from 2° to 1°. The azimuth stays 240°."},
            {"n": 37, "why": "PBSTH22A is uptilted from 2° to 1°. The azimuth stays 60°."},
            {"n": 27, "why": "JYPBB01C is uptilted. The calculated 0.9° is shown as 1°. The azimuth stays 335°."},
            {"n": 34, "why": "BOSDR78B serves the road from 5.2 km and is downtilted from 4° to 6°. The nearer sector is uptilted, shown as 2° to 1°."},
            {"n": 7, "why": "Two far cells serve red samples in this circle. Both are downtilted. The nearer sector is uptilted, shown as 2° to 1°."},
        ],
    },
    {
        "solution": "4th sector, L900",
        "mode": "Active",
        "spots": 45,
        "items": [
            {"n": 4, "why": "The red road is 0.2 km from GBPLS02 and sits in the null, so a new L900 sector is aimed through it. Far cells on other red samples in the circle are downtilted too."},
            {"n": 1, "why": "BOSRK11 is 2.6 km away, inside 3 km, so a 4th L900 sector replaces a new site. It also covers the 1,301 m blank."},
            {"n": 5, "why": "GBGBD24 is 0.4 km away and the spot is in the null. GBSGT12C serves a red sample from 6.1 km and is downtilted from 3° to 5°."},
            {"n": 66, "why": "PBSTH27 is 0.2 km away. A new L900 sector at 255°, E_tilt 7°, fills the null. Existing sectors stay."},
            {"n": 14, "why": "GBSDR46 is 0.7 km away. A new L900 sector at 136°, E_tilt 3°, is aimed at the spot."},
        ],
    },
    {
        "solution": "New 3-sector site",
        "mode": "Active",
        "spots": 12,
        "items": [
            {"n": 2, "why": "The nearest existing site is 3.9 km away, past 3 km. The new site also covers the 2,749 m blank."},
            {"n": 10, "why": "The nearest existing site is 4.8 km away. No existing sector is close enough."},
            {"n": 8, "why": "The nearest existing site is 3.8 km away. The new site covers the red road and the 1,418 m blank."},
            {"n": 12, "why": "The nearest existing site is 3.3 km away, just past the 3 km line, so the plan is a new site."},
            {"n": 61, "why": "The nearest existing site is past 3 km. Each of the three L900 sectors has its own azimuth and tilt."},
        ],
    },
    {
        "solution": "Azimuth",
        "mode": "Idle",
        "spots": 4,
        "items": [
            {"n": 2, "why": "BOSDR52A turns from 355° to 9°. BOSDR76A keeps azimuth 40°, and its tilt is shown as 2° to 1°."},
            {"n": 13, "why": "SGTRS15C turns from 270° to 295°. E_tilt 4° to 0.9° is shown as 4° to 1°."},
            {"n": 17, "why": "The far cell is downtilted from 3° to 5°. SGROY05C also turns, from 350° to 328°, and its tilt is shown as 3° to 1°."},
            {"n": 26, "why": "DPGRT02B turns from 170° to 151°. The tilts are shown as 4° to 5° and 4° to 1°."},
        ],
    },
    {
        "solution": "Tilt",
        "mode": "Idle",
        "spots": 7,
        "items": [
            {"n": 20, "why": "DPFLB01B serves this short road from 7.3 km, so E_tilt goes from 2° to 4°. DPFLB03, 0.4 km away, is not moved."},
            {"n": 1, "why": "The far cell is downtilted from 3° to 5°. The nearer sector is uptilted, and 1.8° is shown as 2°."},
            {"n": 14, "why": "DPHKM11B is uptilted from 3° to 1°. The azimuth stays 120°."},
            {"n": 8, "why": "JYKTL02A is uptilted. The calculated 0.8° is shown as 1°. The azimuth stays 50°."},
            {"n": 3, "why": "BOSJP16B serves the road from 4.1 km and is downtilted from 2° to 4°. The nearer site is not moved."},
        ],
    },
    {
        "solution": "4th sector, L900",
        "mode": "Idle",
        "spots": 12,
        "items": [
            {"n": 12, "why": "The 4th L900 sector on BOSBG02 aims at the 123 m piece. BOSBG26B serves that piece from about 4.9 km. BOSBG03A serves the separate 50 m sample from about 6.2 km. Both are downtilted, E_tilt 2° to 4°."},
            {"n": 11, "why": "BOSRP22 is 1.1 km away and the spot is in the null. A new L900 sector at 359°, E_tilt 2°, is aimed through it."},
            {"n": 16, "why": "SGROY15 is 0.1 km from the spot. A new L900 sector at 118°, E_tilt 10°, covers it. Existing sectors stay."},
            {"n": 22, "why": "DPHKM12 is 1.7 km away. A new L900 sector at 213°, E_tilt 1°, fills the gap."},
            {"n": 24, "why": "JYKTL09 is 2.9 km away, inside 3 km, so a 4th sector is used instead of a new site. The far cell is downtilted as well."},
        ],
    },
    {
        "solution": "New 3-sector site",
        "mode": "Idle",
        "spots": 3,
        "items": [
            {"n": 4, "why": "The nearest existing site is 3.3 km away, just past 3 km, so the final step is a new site."},
            {"n": 10, "why": "The nearest existing site is 4.3 km away. The three L900 sectors are 180°, 300° and 60°."},
            {"n": 23, "why": "The nearest existing site is 3.3 km away. The new site also covers the 304 m blank."},
        ],
    },
]

def _cover_rows() -> dict[tuple[str, int], dict[str, str]]:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(REPORT) as book:
        root = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
    found: dict[tuple[str, int], dict[str, str]] = {}
    for row in root.findall("m:sheetData/m:row", ns):
        cells: dict[str, str] = {}
        for cell in row.findall("m:c", ns):
            ref = cell.get("r") or ""
            col = re.match(r"[A-Z]+", ref)
            text = cell.find("m:is/m:t", ns)
            if col and text is not None and text.text:
                cells[col.group(0)] = text.text
        mode = cells.get("A")
        if mode not in ("Idle", "Active"):
            continue
        number = re.search(r"(\d+)", cells.get("B", ""))
        if number is None:
            continue
        found[(mode, int(number.group(1)))] = cells
    return found


def _map_path(mode: str, number: int) -> Path:
    if mode == "Idle":
        return PLOTS / f"excel_bad_spots_rsrp_spot_{number:02d}.png"
    return PLOTS / f"excel_bad_spots_rsrp_spot_active_{number:02d}.png"


def _circle_box(path: Path) -> dict | None:
    if not CIRCLES.exists():
        return None
    boxes = json.loads(CIRCLES.read_text())
    return boxes.get(path.name)


def _paint_circle(image: Image.Image, box: dict) -> Image.Image:
    """The orange bad-spot oval the workbook keeps as a movable shape."""
    import math

    from PIL import ImageDraw

    recorded_w = float(box["width"]) or image.width
    recorded_h = float(box["height"]) or image.height
    sx = image.width / recorded_w
    sy = image.height / recorded_h
    left = float(box["x"]) * sx
    top = float(box["y"]) * sy
    right = left + float(box["w"]) * sx
    bottom = top + float(box["h"]) * sy
    cx = (left + right) / 2.0
    cy = (top + bottom) / 2.0
    rx = max(2.0, (right - left) / 2.0)
    ry = max(2.0, (bottom - top) / 2.0)
    overlay = image.copy()
    draw = ImageDraw.Draw(overlay)
    steps = 280
    points = []
    for index in range(steps):
        angle = 2.0 * math.pi * index / steps
        points.append((cx + rx * math.cos(angle), cy + ry * math.sin(angle)))
    points.append(points[0])

    def stroke(width: int, fill: tuple[int, int, int], dash: float, gap: float) -> None:
        travel = 0.0
        drawing = True
        for start, end in zip(points, points[1:]):
            length = math.hypot(end[0] - start[0], end[1] - start[1])
            if drawing:
                draw.line([start, end], fill=fill, width=width)
            travel += length
            limit = dash if drawing else gap
            if travel >= limit:
                drawing = not drawing
                travel = 0.0

    dash, gap = (8, 7) if box.get("dotted") else (22, 14)
    stroke(7, (27, 27, 27), dash, gap)
    stroke(4, (255, 102, 0), dash, gap)
    return overlay


def _jpeg(path: Path, width: int = 1400) -> BytesIO:
    image = Image.open(path).convert("RGB")
    box = _circle_box(path)
    if box:
        image = _paint_circle(image, box)
    if image.width > width:
        height = round(image.height * width / image.width)
        image = image.resize((width, height), Image.Resampling.LANCZOS)
    buf = BytesIO()
    image.save(buf, format="JPEG", quality=80, optimize=True)
    buf.seek(0)
    return buf


def _set_run(run, text: str, size: int, color: RGBColor, bold: bool = False) -> None:
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = False


def _box(slide, left, top, width, height, fill: RGBColor, line: RGBColor | None = None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    # Drop the shadow that PowerPoint adds on new shapes.
    effect = shape._element.find(qn("a:effectLst"))
    if effect is not None:
        effect.getparent().remove(effect)
    return shape


def _text(slide, left, top, width, height, text: str, size: int, color: RGBColor,
          bold: bool = False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    shape = slide.shapes.add_textbox(left, top, width, height)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.auto_size = None
    frame.margin_left = Emu(0)
    frame.margin_right = Emu(0)
    frame.margin_top = Emu(0)
    frame.margin_bottom = Emu(0)
    try:
        frame._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    _set_run(paragraph.add_run(), text, size, color, bold)
    return shape


def _footer(slide, page: int, pages: int) -> None:
    _text(
        slide, Inches(0.4), Inches(7.18), Inches(10), Inches(0.24),
        "Bogura drive test  ·  Summary v1.68  ·  Excel v1.67  ·  Web v1.67",
        11, MUTED,
    )
    _text(
        slide, Inches(11.4), Inches(7.18), Inches(1.5), Inches(0.24),
        f"{page}  /  {pages}", 11, MUTED, align=PP_ALIGN.RIGHT,
    )



def _heading(slide, title: str) -> None:
    _box(slide, 0, 0, Inches(13.333), Inches(0.62), NAVY)
    _text(slide, Inches(0.36), Inches(0.1), Inches(12.2), Inches(0.42), title, 24, WHITE, True)


def _stat(slide, left, top, width, label, value) -> None:
    _text(slide, Inches(left), Inches(top), Inches(2.35), Inches(0.28), label, 12, MUTED, True)
    _text(slide, Inches(left + 2.35), Inches(top), Inches(width - 2.35), Inches(0.28), value, 13, INK)


def _cover(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, prs.slide_height, NAVY)
    _box(slide, 0, 0, Inches(0.16), prs.slide_height, GOLD)
    _text(slide, Inches(0.55), Inches(0.28), Inches(8), Inches(0.28), "BOGURA DRIVE TEST", 14, RGBColor(0xF5, 0xC4, 0x00), True)
    _text(slide, Inches(0.55), Inches(0.55), Inches(8), Inches(0.55), "Summary report", 36, WHITE, True)
    _text(slide, Inches(8.4), Inches(0.42), Inches(4.5), Inches(0.5), "11–21 May 2026\nSummary v1.68", 16, RGBColor(0xD6, 0xE4, 0xF0), align=PP_ALIGN.RIGHT)

    cards = [
        ("Idle", "Borga_Idle_V2.0, scanner. Unchanged.\n60,957 plotted samples, best server.\nRSRP median −78.8 dBm."),
        ("Active", "Borga_Active_mode_11Nov26.\nRobi, M-PCell, one band per location.\n288,308 plotted samples. RSRP median −98.1 dBm."),
    ]
    for index, (title, body) in enumerate(cards):
        left = 0.55 + index * 6.3
        _box(slide, Inches(left), Inches(1.4), Inches(6.0), Inches(1.45), RGBColor(0x15, 0x3E, 0x5A))
        _text(slide, Inches(left + 0.22), Inches(1.5), Inches(5.5), Inches(0.3), title, 16, RGBColor(0xF5, 0xC4, 0x00), True)
        _text(slide, Inches(left + 0.22), Inches(1.84), Inches(5.55), Inches(0.9), body, 14, WHITE)

    _text(slide, Inches(0.55), Inches(3.05), Inches(12), Inches(0.3), "Bad spots: 26 idle and 99 active. A 50 m bin is poor when its median RSRP is below −120 dBm.", 15, RGBColor(0xD6, 0xE4, 0xF0))
    counts = [
        ("11 + 42", "Tilt and azimuth", "4 idle and 25 active turn an azimuth"),
        ("12 + 45", "4th sector, L900", "A host inside 3 km, aimed at the spot"),
        ("3 + 12", "New 3-sector site", "Only when the nearest site is past 3 km"),
    ]
    colors = [GOLD, ORANGE, PURPLE]
    for index, ((count, label, note), color) in enumerate(zip(counts, colors)):
        left = 0.55 + index * 4.2
        _box(slide, Inches(left), Inches(3.5), Inches(4.0), Inches(1.45), RGBColor(0x15, 0x3E, 0x5A))
        _box(slide, Inches(left), Inches(3.5), Inches(0.1), Inches(1.45), color)
        _text(slide, Inches(left + 0.25), Inches(3.6), Inches(3.6), Inches(0.4), count, 22, WHITE, True)
        _text(slide, Inches(left + 0.25), Inches(4.05), Inches(3.6), Inches(0.3), label, 14, WHITE, True)
        _text(slide, Inches(left + 0.25), Inches(4.4), Inches(3.6), Inches(0.42), note, 12, RGBColor(0xD6, 0xE4, 0xF0))

    _text(
        slide, Inches(0.55), Inches(5.15), Inches(12.2), Inches(0.85),
        "This deck has the RSRP, RSRQ and SINR plots, the bad-spot counts, then the examples. "
        "Active shows 5 azimuth, 5 tilt, 5 fourth-sector and 5 new-site maps. "
        "Idle shows all 4 azimuth turns, 5 tilts, 5 fourth sectors and all 3 new sites.",
        15, RGBColor(0xD6, 0xE4, 0xF0),
    )
    _text(
        slide, Inches(0.55), Inches(6.15), Inches(12.2), Inches(0.7),
        "Proposed azimuth, E_tilt and M_tilt are the nearest whole degree. Half a degree rounds up. "
        "Slides after the Annex hold the rules, the map key, and the check that every red sample inside the circle is included.",
        14, RGBColor(0xAE, 0xC6, 0xD6),
    )


def _kpi(prs: Presentation, spec: dict, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, spec["title"] + "  ·  plot and summary")
    _text(slide, Inches(0.36), Inches(0.66), Inches(12.6), Inches(0.42), spec["rule"], 13, MUTED)
    left = 0.28
    for path, caption in ((spec["idle"], "Idle"), (spec["active"], "Active")):
        with Image.open(path) as source:
            aspect = source.width / source.height
        height = 4.95
        width = height * aspect
        _text(slide, Inches(left), Inches(1.12), Inches(width), Inches(0.22), caption, 13, NAVY, True, PP_ALIGN.CENTER)
        slide.shapes.add_picture(_jpeg(path, width=900), Inches(left), Inches(1.36), Inches(width), Inches(height))
        left += width + 0.16
    panel = left
    panel_w = 13.05 - panel
    _text(slide, Inches(panel), Inches(1.12), Inches(panel_w), Inches(0.26), "IDLE", 12, NAVY, True)
    top = 1.38
    for label, value in spec["idle_stats"]:
        _stat(slide, panel, top, panel_w, label, value)
        top += 0.28
    top += 0.08
    _text(slide, Inches(panel), Inches(top), Inches(panel_w), Inches(0.26), "ACTIVE", 12, NAVY, True)
    top += 0.28
    for label, value in spec["active_stats"]:
        _stat(slide, panel, top, panel_w, label, value)
        top += 0.28
    top += 0.12
    _text(slide, Inches(panel), Inches(top), Inches(panel_w), Inches(0.24), "LEGEND", 12, NAVY, True)
    top += 0.28
    for hex_color, label in spec["legend"]:
        red = int(hex_color[1:3], 16)
        green = int(hex_color[3:5], 16)
        blue = int(hex_color[5:7], 16)
        _box(slide, Inches(panel), Inches(top + 0.04), Inches(0.22), Inches(0.18), RGBColor(red, green, blue), LINE)
        _text(slide, Inches(panel + 0.32), Inches(top), Inches(panel_w - 0.32), Inches(0.24), label, 12, INK)
        top += 0.26
    _footer(slide, page, pages)


def _bad_summary(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Bad spot summary")
    headers = ["Solution", "Idle spots", "Idle shown", "Active spots", "Active shown"]
    rows = [
        ("Azimuth", "4", "4", "25", "5", GOLD),
        ("Tilt, azimuth unchanged", "7", "5", "17", "5", GREEN),
        ("4th sector, L900", "12", "5", "45", "5", ORANGE),
        ("New 3-sector site", "3", "3", "12", "5", PURPLE),
        ("All bad spots", "26", "", "99", "", NAVY),
    ]
    widths = [4.3, 1.9, 1.9, 2.1, 2.1]
    left0 = 0.45
    top = 0.95
    x = left0
    for header, width in zip(headers, widths):
        _box(slide, Inches(x), Inches(top), Inches(width), Inches(0.48), NAVY)
        _text(slide, Inches(x), Inches(top + 0.08), Inches(width), Inches(0.32), header, 13, WHITE, True, PP_ALIGN.CENTER)
        x += width
    for row_index, (label, idle, idle_shown, active, active_shown, color) in enumerate(rows):
        y = top + 0.48 + row_index * 0.72
        values = [label, idle, idle_shown, active, active_shown]
        x = left0
        fill = RGBColor(0xEE, 0xF3, 0xF7) if row_index % 2 == 0 else WHITE
        if row_index == len(rows) - 1:
            fill = RGBColor(0xE8, 0xF0, 0xF6)
        for col, (value, width) in enumerate(zip(values, widths)):
            _box(slide, Inches(x), Inches(y), Inches(width), Inches(0.72), fill, LINE)
            if col == 0:
                _box(slide, Inches(x), Inches(y), Inches(0.1), Inches(0.72), color)
            _text(
                slide, Inches(x + (0.18 if col == 0 else 0)), Inches(y + 0.16), Inches(width), Inches(0.4),
                value, 16, INK, col == 0 or row_index == len(rows) - 1,
                PP_ALIGN.CENTER if col else PP_ALIGN.LEFT,
            )
            x += width
    _text(
        slide, Inches(0.45), Inches(5.2), Inches(12.4), Inches(1.7),
        "Tilt and azimuth together are 11 idle and 42 active. The table splits that group. "
        "An azimuth row is a spot where an existing azimuth is turned. Most of those also change tilt. "
        "A tilt row leaves the azimuth as it is.\n"
        "Idle has 4 azimuth turns and 3 new sites, so those examples are the full set. "
        "Every other group shows 5. The other spots stay in the Excel workbook and use the same rules.\n"
        "On the example slides, a tenth of a degree is shown as the nearest whole degree. 0.5 rounds up.",
        15, INK,
    )
    _footer(slide, page, pages)


def _example(prs: Presentation, group: dict, item: dict, record: dict[str, str], ordinal: int, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    color = SOLUTION_COLOR[group["solution"]]
    total = len(group["items"])
    _box(slide, 0, 0, prs.slide_width, Inches(0.56), NAVY)
    _box(slide, 0, 0, Inches(0.12), Inches(0.56), color)
    _text(
        slide, Inches(0.28), Inches(0.08), Inches(12.6), Inches(0.4),
        f"{group['solution']}    ·    {ordinal} of {total}    ·    {group['mode']} Bad Spot {item['n']}",
        20, WHITE, True,
    )
    picture = _map_path(group["mode"], item["n"])
    with Image.open(picture) as source:
        aspect = source.width / source.height
    image = _jpeg(picture)
    max_w, max_h = 6.7, 6.35
    width = max_w
    height = width / aspect
    if height > max_h:
        height = max_h
        width = height * aspect
    slide.shapes.add_picture(image, Inches(0.16), Inches(0.68), Inches(width), Inches(height))

    panel = 7.05
    panel_w = 5.95
    _text(slide, Inches(panel), Inches(0.66), Inches(panel_w), Inches(0.22), group["solution"].upper(), 12, color, True)
    _text(slide, Inches(panel), Inches(0.9), Inches(panel_w), Inches(1.05), item["why"], 13, INK)
    fields = [
        ("Mode", group["mode"]),
        ("Bad spot", str(item["n"])),
        ("Site", record.get("C", "")),
        ("Sector", record.get("D", "")),
        ("Serving cell", record.get("E", "")),
        ("Population", record.get("F", "")),
    ]
    top = 2.0
    for label, value in fields:
        _text(slide, Inches(panel), Inches(top), Inches(1.35), Inches(0.26), label, 11, MUTED, True)
        _text(slide, Inches(panel + 1.35), Inches(top), Inches(panel_w - 1.35), Inches(0.26), value, 12, INK)
        top += 0.26
    _text(slide, Inches(panel), Inches(top + 0.04), Inches(panel_w), Inches(0.22), "FINAL SOLUTION  ·  NEAREST DEGREE", 11, color, True)
    solution = _whole_degrees(record.get("H", ""))
    size = 11 if len(solution) > 420 else 12
    _box(slide, Inches(panel), Inches(top + 0.28), Inches(panel_w), Inches(7.05 - (top + 0.28)), PAPER, LINE)
    _text(slide, Inches(panel + 0.1), Inches(top + 0.34), Inches(panel_w - 0.2), Inches(6.95 - (top + 0.34)), solution, size, INK)
    _footer(slide, page, pages)


def _annex(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, prs.slide_height, NAVY)
    _box(slide, 0, 0, Inches(0.16), prs.slide_height, GOLD)
    _text(slide, Inches(0.7), Inches(2.15), Inches(12), Inches(0.9), "Annex", 60, WHITE, True)
    _text(
        slide, Inches(0.7), Inches(3.3), Inches(11.5), Inches(1.4),
        "The report ends on the previous slide. The slides after this one add the rounding rule, "
        "how a solution is chosen, how to read a map, and the check that every red sample inside the circle is included.",
        20, RGBColor(0xD6, 0xE4, 0xF0),
    )
    _text(slide, Inches(0.7), Inches(6.9), Inches(10), Inches(0.28), "Bogura drive test  ·  Summary v1.68", 14, RGBColor(0xAE, 0xC6, 0xD6))
    _text(slide, Inches(11.2), Inches(6.9), Inches(1.6), Inches(0.28), f"{page}  /  {pages}", 14, RGBColor(0xAE, 0xC6, 0xD6), align=PP_ALIGN.RIGHT)


def _annex_rounding(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  azimuth and tilt are whole degrees")
    rows = [
        ("What is rounded", "The proposed azimuth, E_tilt and M_tilt. The measured value in the dump is not changed."),
        ("Which way", "The nearest whole degree. 0.5 rounds up. 1.8° is shown as 2°. 0.9° is shown as 1°. 1.2° is shown as 1°."),
        ("A very small change", "If the nearest degree is the same as the degree it already has, the line says the tilt or azimuth stays."),
        ("Where it applies", "Every example in this deck. The next workbook publish uses the same rounding on the proposed value."),
        ("What does not move", "A 4th-sector azimuth and a new-site azimuth were already whole degrees. Downtilt of a far cell stays a step of 2°."),
    ]
    top = 0.95
    for label, body in rows:
        _box(slide, Inches(0.4), Inches(top), Inches(12.5), Inches(1.05), PAPER, LINE)
        _text(slide, Inches(0.58), Inches(top + 0.12), Inches(2.7), Inches(0.8), label, 15, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        _text(slide, Inches(3.4), Inches(top + 0.14), Inches(9.2), Inches(0.78), body, 15, INK, anchor=MSO_ANCHOR.MIDDLE)
        top += 1.15
    _footer(slide, page, pages)


def _annex_rules(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  how a solution is chosen")
    cards = [
        (GOLD, "Azimuth", "A sector inside 2.5 km can be turned when it does not already face the spot. The turn is kept when the customers on the new bearing are still served. A turn under 15° on the nearest sector is not added just because a far cell was downtilted."),
        (GREEN, "Tilt", "Downtilt of an existing sector is at most 2°. A cell that serves a red sample from at least 4 km, and much farther than the nearest site, is downtilted first. The nearest sector is uptilted only when its own beam falls short."),
        (ORANGE, "4th sector", "Used when the spot sits in the null inside 1.5 km, when one or two sites inside 2 km leave it in the null, when the required downtilt would pass 2°, or when a host inside 3 km can aim at the spot. Existing sectors are not moved for that new sector."),
        (PURPLE, "New site", "Used only when the nearest existing site is farther than 3 km. Three L900 sectors are drawn at the spot. The sector toward the red road is long enough to cross it, and every sector shows its azimuth."),
    ]
    for index, (color, title, body) in enumerate(cards):
        col = index % 2
        row = index // 2
        left = 0.35 + col * 6.5
        top = 0.9 + row * 3.0
        _box(slide, Inches(left), Inches(top), Inches(6.25), Inches(2.8), PAPER, LINE)
        _box(slide, Inches(left), Inches(top), Inches(6.25), Inches(0.1), color)
        _text(slide, Inches(left + 0.22), Inches(top + 0.25), Inches(5.8), Inches(0.4), title, 20, NAVY, True)
        _text(slide, Inches(left + 0.22), Inches(top + 0.75), Inches(5.8), Inches(1.85), body, 14, INK)
    _footer(slide, page, pages)


def _annex_read(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  how to read an example")
    items = [
        ("Red on the road", "The poor samples. The length beside the red is that red stretch, and it sits off the samples."),
        ("Orange circle", "The bad spot. A blank with no plot inside the circle is part of the same spot."),
        ("Yellow dashed line", "The cell that served the red samples. It starts at the end of that sector. It is not the proposed sector. The length is from the arm tip, so it is shorter than the site-centre distance."),
        ("Tilt on the dashed line", "When that cell is an overshooter, the line also carries the E_tilt change."),
        ("Orange arm", "A proposed 4th L900 sector, drawn from the existing site toward the spot. The words sit along the arm."),
        ("Purple arms", "A proposed new 3-sector site. Every arm is labeled with its azimuth."),
    ]
    for index, (title, body) in enumerate(items):
        col = index % 2
        row = index // 2
        left = 0.35 + col * 6.5
        top = 0.88 + row * 2.0
        _box(slide, Inches(left), Inches(top), Inches(6.25), Inches(1.85), PAPER, LINE)
        _text(slide, Inches(left + 0.22), Inches(top + 0.14), Inches(5.85), Inches(0.36), title, 16, NAVY, True)
        _text(slide, Inches(left + 0.22), Inches(top + 0.55), Inches(5.85), Inches(1.15), body, 14, INK)
    _footer(slide, page, pages)


def _annex_circle(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  every red sample inside the circle")
    _text(
        slide, Inches(0.45), Inches(0.85), Inches(12.4), Inches(0.7),
        "A far cell is downtilted when it serves any red sample inside the orange circle, including a short piece that is not the longest run. Later report builds keep this check.",
        16, INK,
    )
    rows = [
        ("Idle 12", "BOSBG03A L2100 serves the separate 50 m sample from 6,238 m. E_tilt 2° to 4°. The map line reads 5,853 m because it starts at the arm tip. BOSBG26B still covers the 123 m piece."),
        ("Idle 5", "BOSJP30B L2600 from 5,720 m, E_tilt 4° to 6°."),
        ("Active 4, 5, 6", "GBPLS12A, GBSGT12C, and on Active 6 the cells BODNT08B L1800, BODNT08B L900 and SGKZP18C."),
        ("Active 7, 9, 19, 25, 31", "JYKLI02C, both DPHKM12A layers, DPBRM24C, BOSRP16A and JYAKL20A. Each is downtilted, and the dashed line carries that change."),
    ]
    top = 1.7
    for label, body in rows:
        _box(slide, Inches(0.45), Inches(top), Inches(12.4), Inches(1.15), PAPER, LINE)
        _text(slide, Inches(0.62), Inches(top + 0.14), Inches(2.8), Inches(0.85), label, 15, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        _text(slide, Inches(3.5), Inches(top + 0.12), Inches(9.1), Inches(0.9), body, 14, INK, anchor=MSO_ANCHOR.MIDDLE)
        top += 1.25
    _footer(slide, page, pages)


def _annex_workbook(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  where the full set is")
    rows = [
        ("Excel v1.67", "A map, a suggestion table and a movable circle for all 26 idle spots and all 99 active spots. The Cover sheet lists every spot. EP Details has the existing and proposed height, azimuth, E_tilt and M_tilt."),
        ("This deck", "The examples above. Idle azimuth and Idle new site are complete. The other groups are 5 examples each."),
        ("Web v1.67", "The live map. Site names are white with a dark edge. Sector names on that map stay small. With KPI set to Blank, Info on a sector or the site centre opens the cell table."),
        ("RSRQ and SINR", "The coverage maps are in this deck. The workbook no longer has separate RSRQ or SINR bad-spot sheets. Poor RSRQ is at or below −20 dB. Poor SINR is at or below 0 dB."),
    ]
    top = 0.95
    for label, body in rows:
        _box(slide, Inches(0.4), Inches(top), Inches(12.5), Inches(1.35), PAPER, LINE)
        _text(slide, Inches(0.58), Inches(top + 0.18), Inches(2.5), Inches(1.0), label, 16, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        _text(slide, Inches(3.2), Inches(top + 0.16), Inches(9.4), Inches(1.05), body, 15, INK, anchor=MSO_ANCHOR.MIDDLE)
        top += 1.48
    _footer(slide, page, pages)


KPI_SLIDES = [
    {
        "title": "RSRP",
        "rule": "Idle is the best server per sweep. Active keeps the strongest RSRP at each location. A bad spot is a 50 m median below −120 dBm.",
        "idle": PLOTS / "excel_route_rsrp.png",
        "active": PLOTS / "excel_route_rsrp_active.png",
        "idle_stats": [("Median", "−78.8 dBm"), ("−95 to 0", "90.2%"), ("−110 to −95", "9.0%"), ("Below −120", "0.2%, 142"), ("Bad spots", "26")],
        "active_stats": [("Median", "−98.1 dBm"), ("−95 to 0", "40.6%"), ("−110 to −95", "46.4%"), ("Below −120", "0.4%, 1,077"), ("Bad spots", "99")],
        "legend": [("#008000", "−95 to 0"), ("#00FF00", "−110 to −95"), ("#FFFF00", "−115 to −110"), ("#FF6666", "−120 to −115"), ("#FF0000", "below −120")],
    },
    {
        "title": "RSRQ",
        "rule": "RSRQ comes from the same row as the plotted RSRP. Poor is at or below −20 dB. The red colour on the map is the wider bin, −30 to −17.",
        "idle": PLOTS / "excel_route_rsrq.png",
        "active": PLOTS / "excel_route_rsrq_active.png",
        "idle_stats": [("Median", "−11.8 dB"), ("−10 to 0", "4.0%"), ("−17 to −10", "92.8%"), ("−30 to −17", "3.2%"), ("At or below −20", "1.7%")],
        "active_stats": [("Median", "−9.7 dB"), ("−10 to 0", "55.4%"), ("−17 to −10", "43.2%"), ("−30 to −17", "1.4%"), ("At or below −20", "0.2%")],
        "legend": [("#00FF00", "−10 to 0"), ("#FFFF00", "−17 to −10"), ("#FF0000", "−30 to −17")],
    },
    {
        "title": "SINR",
        "rule": "SINR comes from the same row as the plotted RSRP. Idle has a value on 60,874 samples. Active has a value on 277,699. Poor is at or below 0 dB.",
        "idle": PLOTS / "excel_route_sinr.png",
        "active": PLOTS / "excel_route_sinr_active.png",
        "idle_stats": [("Median", "9.9 dB"), ("15 and above", "27.4%"), ("5 to 15", "45.2%"), ("0 to 5", "17.5%"), ("At or below 0", "10.0%")],
        "active_stats": [("Median", "8.1 dB"), ("15 and above", "20.8%"), ("5 to 15", "44.7%"), ("0 to 5", "21.8%"), ("At or below 0", "13.1%")],
        "legend": [("#0070C0", "15 and above"), ("#00B050", "10 to 15"), ("#00B0F0", "5 to 10"), ("#FFFF00", "0 to 5"), ("#FF6600", "−5 to 0"), ("#FF0000", "below −5")],
    },
]


def build() -> Path:
    if not REPORT.exists():
        raise SystemExit(f"Missing {REPORT.name}")
    rows = _cover_rows()
    missing = []
    for group in GROUPS:
        for item in group["items"]:
            if (group["mode"], item["n"]) not in rows:
                missing.append(f"{group['mode']} {item['n']}")
            if not _map_path(group["mode"], item["n"]).exists():
                missing.append(f"map {group['mode']} {item['n']}")
    for spec in KPI_SLIDES:
        for key in ("idle", "active"):
            if not spec[key].exists():
                missing.append(spec[key].name)
    if missing:
        raise SystemExit("Missing example data: " + ", ".join(missing))

    example_count = sum(len(group["items"]) for group in GROUPS)
    # cover, 3 KPI, bad-spot summary, examples, annex, 5 annex detail slides
    pages = 1 + len(KPI_SLIDES) + 1 + example_count + 1 + 5
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Bogura Drive Test Summary v1.68"
    prs.core_properties.subject = "Summary report, KPI plots, and bad-spot examples"
    _cover(prs)
    page = 2
    for spec in KPI_SLIDES:
        _kpi(prs, spec, page, pages)
        page += 1
    _bad_summary(prs, page, pages)
    page += 1
    for group in GROUPS:
        for ordinal, item in enumerate(group["items"], start=1):
            _example(prs, group, item, rows[(group["mode"], item["n"])], ordinal, page, pages)
            page += 1
    _annex(prs, page, pages)
    page += 1
    _annex_rounding(prs, page, pages)
    page += 1
    _annex_rules(prs, page, pages)
    page += 1
    _annex_read(prs, page, pages)
    page += 1
    _annex_circle(prs, page, pages)
    page += 1
    _annex_workbook(prs, page, pages)
    if page != pages:
        raise SystemExit(f"Page count {page} does not match {pages}")
    OUTPUT.mkdir(parents=True, exist_ok=True)
    prs.save(DEST)
    return DEST


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path} ({path.stat().st_size / 1e6:.1f} MB, {len(GROUPS)} groups)")
