#!/usr/bin/env python3
"""One summary presentation: the bad-spot rules, the counts, and three examples of each solution."""
from __future__ import annotations

import json
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
DEST = OUTPUT / "Bogura_DriveTest_Summary_v1.67.pptx"
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

SOLUTION_COLOR = {
    "Tilt and azimuth": GOLD,
    "4th sector, L900": ORANGE,
    "New 3-sector site": PURPLE,
}

# Three examples of each final solution. The words on the slide come from the workbook.
EXAMPLES = [
    {
        "solution": "Tilt and azimuth",
        "mode": "Idle",
        "n": 1,
        "why": "BOSDR56A L2600 is serving this road from about 6.6 km, so it is downtilted from 3° to 5°. BOGBT18C still needs an uptilt, from 3° to 1.8°, because that nearer beam falls short.",
    },
    {
        "solution": "Tilt and azimuth",
        "mode": "Active",
        "n": 97,
        "why": "The tilt is already right for 0.6 km. The azimuth turns 8°, from 140° to 132°. The dashed line leaves GBSDR37A because that cell served the red samples. The plan turns GBSDR37B.",
    },
    {
        "solution": "Tilt and azimuth",
        "mode": "Idle",
        "n": 20,
        "why": "A far cell serves this short red road. That cell is downtilted so it stops overshooting. The nearest sector is not moved, because its own beam already fits the road.",
    },
    {
        "solution": "4th sector, L900",
        "mode": "Active",
        "n": 4,
        "why": "The red road is 0.2 km from GBPLS02 and sits in the null, so a new L900 sector is aimed through it. GBPLS12A, GBSDL14B and GBSDL27B also serve red samples in this circle, from 4.4 km to 6.5 km, so those beams are pulled back.",
    },
    {
        "solution": "4th sector, L900",
        "mode": "Idle",
        "n": 12,
        "why": "The 4th L900 sector on BOSBG02 is aimed at the 123 m red piece. BOSBG26B serves that piece from about 4.9 km and is downtilted. BOSBG03A serves the separate 50 m sample from about 6.2 km, so that beam is pulled back too, E_tilt 2° to 4°.",
    },
    {
        "solution": "4th sector, L900",
        "mode": "Active",
        "n": 1,
        "why": "BOSRK11 is 2.6 km away, inside 3 km, so a 4th L900 sector replaces a new site. It also covers the 1,301 m blank.",
    },
    {
        "solution": "New 3-sector site",
        "mode": "Idle",
        "n": 4,
        "why": "The nearest existing site is 3.3 km away, just past the 3 km line, so the final step is a new site.",
    },
    {
        "solution": "New 3-sector site",
        "mode": "Active",
        "n": 2,
        "why": "This spot includes the long blank. The nearest existing site is 3.9 km away, so the plan is a new site on the spot.",
    },
    {
        "solution": "New 3-sector site",
        "mode": "Active",
        "n": 10,
        "why": "The nearest existing site is 4.8 km away. No existing sector is close enough to cover this spot.",
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
        "Bogura drive test  ·  Summary v1.67  ·  Excel v1.67  ·  Web v1.67",
        11, MUTED,
    )
    _text(
        slide, Inches(11.4), Inches(7.18), Inches(1.5), Inches(0.24),
        f"{page}  /  {pages}", 11, MUTED, align=PP_ALIGN.RIGHT,
    )


def _title(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, prs.slide_height, NAVY)
    _box(slide, 0, 0, Inches(0.18), prs.slide_height, GOLD)
    _text(slide, Inches(0.7), Inches(1.35), Inches(11), Inches(0.35), "BOGURA DRIVE TEST", 16, RGBColor(0xF5, 0xC4, 0x00), True)
    _text(slide, Inches(0.7), Inches(1.8), Inches(11.5), Inches(1.3), "Bad spot summary", 44, WHITE, True)
    _text(
        slide, Inches(0.7), Inches(3.25), Inches(11), Inches(0.8),
        "Idle and Active. Three examples of each solution:\ntilt and azimuth, a 4th L900 sector, and a new 3-sector site.",
        20, RGBColor(0xD6, 0xE4, 0xF0),
    )
    cards = [
        ("11  +  42", "Tilt and azimuth", GOLD),
        ("12  +  45", "4th sector, L900", ORANGE),
        ("3  +  12", "New 3-sector site", PURPLE),
    ]
    for index, (count, label, color) in enumerate(cards):
        left = Inches(0.7 + index * 3.9)
        _box(slide, left, Inches(4.7), Inches(3.55), Inches(1.45), RGBColor(0x15, 0x3E, 0x5A))
        _box(slide, left, Inches(4.7), Inches(0.1), Inches(1.45), color)
        _text(slide, left + Inches(0.28), Inches(4.85), Inches(3.1), Inches(0.55), count, 26, WHITE, True)
        _text(slide, left + Inches(0.28), Inches(5.45), Inches(3.1), Inches(0.4), label, 14, RGBColor(0xD6, 0xE4, 0xF0))
    _text(slide, Inches(0.7), Inches(6.6), Inches(11), Inches(0.3), "Idle count  +  Active count.   26 idle bad spots and 99 active bad spots.", 14, RGBColor(0xAE, 0xC6, 0xD6))


def _measured(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, Inches(0.72), NAVY)
    _text(slide, Inches(0.4), Inches(0.16), Inches(12), Inches(0.42), "What was measured", 26, WHITE, True)
    rows = [
        ("Idle", "Borga_Idle_V2.0. This drive is unchanged."),
        ("Active", "Borga_Active_mode_11Nov26. 11–21 May 2026. Robi, M-PCell, one band at each location: the strongest RSRP."),
        ("Plotted Active rows", "288,308 of 553,201 Robi M-PCell rows. 2,185,081 rows after the duplicate drop."),
        ("Bad spot", "A 50 m bin is poor when its median RSRP is below −120 dBm. A blank with no plot inside the orange circle is counted with that spot."),
        ("This deck", "The rules, the counts, and three examples of each solution. Every bad spot stays in the Excel workbook."),
    ]
    top = 1.05
    for label, body in rows:
        _box(slide, Inches(0.4), Inches(top), Inches(12.5), Inches(1.05), PAPER, LINE)
        _text(slide, Inches(0.6), Inches(top + 0.12), Inches(2.4), Inches(0.8), label, 16, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        _text(slide, Inches(3.15), Inches(top + 0.14), Inches(9.4), Inches(0.78), body, 16, INK, anchor=MSO_ANCHOR.MIDDLE)
        top += 1.15
    _footer(slide, page, pages)


def _rules(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, Inches(0.72), NAVY)
    _text(slide, Inches(0.4), Inches(0.16), Inches(12), Inches(0.42), "How a solution is chosen", 26, WHITE, True)
    cards = [
        (GOLD, "1", "Tilt and azimuth",
         "A cell serving the red road from more than 4 km, and much farther than the nearest site, is downtilted first so it stops overshooting. The nearest sector stays unless its own beam misses the road. Otherwise a sector inside 2.5 km can be aimed at the spot. Downtilt of an existing sector does not increase by more than 2°."),
        (ORANGE, "2", "4th sector, L900",
         "Used when the spot sits in the null between sectors, when the required downtilt would be more than 2°, or when an existing site inside 3 km can aim at the spot. Existing sectors are not moved."),
        (PURPLE, "3", "New 3-sector site",
         "Used only when the nearest existing site is farther than 3 km. The site is drawn at the spot, with an azimuth, E_tilt and M_tilt for each of the three L900 sectors."),
    ]
    for index, (color, number, title, body) in enumerate(cards):
        left = Inches(0.35 + index * 4.3)
        _box(slide, left, Inches(1.15), Inches(4.05), Inches(5.55), PAPER, LINE)
        _box(slide, left, Inches(1.15), Inches(4.05), Inches(0.12), color)
        _text(slide, left + Inches(0.25), Inches(1.5), Inches(3.5), Inches(0.4), number, 20, color, True)
        _text(slide, left + Inches(0.25), Inches(2.05), Inches(3.55), Inches(0.8), title, 22, NAVY, True)
        _text(slide, left + Inches(0.25), Inches(3.05), Inches(3.55), Inches(3.2), body, 16, INK)
    _footer(slide, page, pages)


def _counts(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, Inches(0.72), NAVY)
    _text(slide, Inches(0.4), Inches(0.16), Inches(12), Inches(0.42), "Bad spot summary across mode", 26, WHITE, True)
    headers = ["Solution", "Idle", "Active", "Together"]
    data = [
        ("Tilt and azimuth", "11", "42", "53", GOLD),
        ("4th sector, L900", "12", "45", "57", ORANGE),
        ("New 3-sector site", "3", "12", "15", PURPLE),
        ("All bad spots", "26", "99", "125", NAVY),
    ]
    widths = [4.6, 2.2, 2.2, 2.4]
    left0 = 0.85
    top = 1.45
    x = left0
    for header, width in zip(headers, widths):
        _box(slide, Inches(x), Inches(top), Inches(width), Inches(0.62), NAVY)
        _text(slide, Inches(x), Inches(top + 0.1), Inches(width), Inches(0.42), header, 16, WHITE, True, PP_ALIGN.CENTER)
        x += width
    for row_index, (label, idle, active, total, color) in enumerate(data):
        y = top + 0.62 + row_index * 0.95
        values = [label, idle, active, total]
        x = left0
        fill = RGBColor(0xEE, 0xF3, 0xF7) if row_index % 2 == 0 else WHITE
        if row_index == len(data) - 1:
            fill = RGBColor(0xE8, 0xF0, 0xF6)
        for col, (value, width) in enumerate(zip(values, widths)):
            _box(slide, Inches(x), Inches(y), Inches(width), Inches(0.95), fill, LINE)
            if col == 0:
                _box(slide, Inches(x), Inches(y), Inches(0.1), Inches(0.95), color)
            _text(
                slide, Inches(x + (0.2 if col == 0 else 0)), Inches(y + 0.22), Inches(width), Inches(0.5),
                value, 20, INK, col == 0 or row_index == len(data) - 1, PP_ALIGN.CENTER if col else PP_ALIGN.LEFT,
            )
            x += width
    _text(
        slide, Inches(0.85), Inches(6.35), Inches(11.5), Inches(0.55),
        "The next slides show three examples of each solution. The other 116 bad spots use the same rules and are in the workbook.",
        16, MUTED,
    )
    _footer(slide, page, pages)


def _legend(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, Inches(0.72), NAVY)
    _text(slide, Inches(0.4), Inches(0.16), Inches(12), Inches(0.42), "How to read an example", 26, WHITE, True)
    items = [
        ("Red on the road", "The poor samples. The length beside the red is that red stretch."),
        ("Orange circle", "The bad spot. A blank with no plot inside the circle is part of the same spot."),
        ("Yellow dashed line", "The cell that served the red samples. It starts at the end of that sector. It is not the plan."),
        ("Light Y", "The existing sectors of a site in the view. The arms meet at the site centre."),
        ("Orange arm", "A proposed 4th L900 sector, drawn from the existing site toward the spot."),
        ("Purple arms", "A proposed new 3-sector site, drawn at the planned coordinates."),
    ]
    for index, (title, body) in enumerate(items):
        col = index % 2
        row = index // 2
        left = 0.4 + col * 6.45
        top = 1.15 + row * 1.85
        _box(slide, Inches(left), Inches(top), Inches(6.2), Inches(1.65), PAPER, LINE)
        _text(slide, Inches(left + 0.25), Inches(top + 0.2), Inches(5.7), Inches(0.4), title, 20, NAVY, True)
        _text(slide, Inches(left + 0.25), Inches(top + 0.7), Inches(5.7), Inches(0.75), body, 16, INK)
    _footer(slide, page, pages)


def _example(prs: Presentation, item: dict, record: dict[str, str], ordinal: int, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    color = SOLUTION_COLOR[item["solution"]]
    _box(slide, 0, 0, prs.slide_width, Inches(0.62), NAVY)
    _box(slide, 0, 0, Inches(0.12), Inches(0.62), color)
    _text(
        slide, Inches(0.32), Inches(0.12), Inches(10), Inches(0.4),
        f"{item['solution']}    ·    {ordinal} of 3    ·    {item['mode']} Bad Spot {item['n']}",
        20, WHITE, True,
    )
    picture = _map_path(item["mode"], item["n"])
    with Image.open(picture) as source:
        aspect = source.width / source.height
    image = _jpeg(picture)
    max_w, max_h = 6.9, 6.28
    width = max_w
    height = width / aspect
    if height > max_h:
        height = max_h
        width = height * aspect
    slide.shapes.add_picture(image, Inches(0.18), Inches(0.76), Inches(width), Inches(height))

    panel_left = 7.35
    _text(slide, Inches(panel_left), Inches(0.76), Inches(5.6), Inches(0.28), item["solution"].upper(), 12, color, True)
    _text(slide, Inches(panel_left), Inches(1.08), Inches(5.6), Inches(1.15), item["why"], 14, INK)

    fields = [
        ("Mode", item["mode"]),
        ("Bad spot", str(item["n"])),
        ("Site", record.get("C", "")),
        ("Sector", record.get("D", "")),
        ("Serving cell", record.get("E", "")),
        ("Population", record.get("F", "")),
    ]
    top = 2.32
    for label, value in fields:
        _text(slide, Inches(panel_left), Inches(top), Inches(1.45), Inches(0.32), label, 12, MUTED, True)
        _text(slide, Inches(panel_left + 1.5), Inches(top), Inches(3.8), Inches(0.32), value, 13, INK)
        top += 0.32
    _text(slide, Inches(panel_left), Inches(top + 0.08), Inches(5.2), Inches(0.28), "FINAL SOLUTION", 12, color, True)
    _box(slide, Inches(panel_left), Inches(top + 0.38), Inches(5.6), Inches(2.35), PAPER, LINE)
    _text(slide, Inches(panel_left + 0.14), Inches(top + 0.48), Inches(5.3), Inches(2.15), record.get("H", ""), 13, INK)
    _footer(slide, page, pages)


def _close(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, prs.slide_height, NAVY)
    _text(slide, Inches(0.7), Inches(1.5), Inches(12), Inches(0.8), "The full set is in the workbook", 32, WHITE, True)
    lines = [
        "Excel v1.67 has a map, a suggestion table and a movable circle for all 26 idle spots and all 99 active spots.",
        "The Cover sheet lists every spot, with the sector, the serving cell and the final solution.",
        "EP Details is the sector and cell tracker: existing height, azimuth, E_tilt and M_tilt beside the proposed values.",
        "The web map is v1.67. Site names are white with a dark edge. With KPI set to Blank, Info on a sector or the site centre opens the cell table. Sector names on the live map stay small. The larger sector names are on the bad-spot maps only.",
    ]
    top = 2.6
    for line in lines:
        _text(slide, Inches(0.7), Inches(top), Inches(11.8), Inches(0.7), line, 18, RGBColor(0xD6, 0xE4, 0xF0))
        top += 0.85


def build() -> Path:
    if not REPORT.exists():
        raise SystemExit(f"Missing {REPORT.name}")
    rows = _cover_rows()
    missing = []
    for item in EXAMPLES:
        if (item["mode"], item["n"]) not in rows:
            missing.append(f"{item['mode']} {item['n']}")
        if not _map_path(item["mode"], item["n"]).exists():
            missing.append(f"map {item['mode']} {item['n']}")
    if missing:
        raise SystemExit("Missing example data: " + ", ".join(missing))

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Bogura Drive Test Summary v1.67"
    prs.core_properties.subject = "Three examples of each bad-spot solution"
    pages = 5 + len(EXAMPLES) + 1
    _title(prs)
    _measured(prs, 2, pages)
    _rules(prs, 3, pages)
    _counts(prs, 4, pages)
    _legend(prs, 5, pages)
    seen: dict[str, int] = {}
    for item in EXAMPLES:
        seen[item["solution"]] = seen.get(item["solution"], 0) + 1
        _example(prs, item, rows[(item["mode"], item["n"])], seen[item["solution"]], 5 + seen_total(seen), pages)
    _close(prs, pages, pages)
    OUTPUT.mkdir(parents=True, exist_ok=True)
    prs.save(DEST)
    return DEST


def seen_total(seen: dict[str, int]) -> int:
    return sum(seen.values())


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path} ({path.stat().st_size / 1e6:.1f} MB)")
