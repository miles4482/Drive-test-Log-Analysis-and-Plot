#!/usr/bin/env python3
"""Bogura summary deck: route and radio KPIs, bad-spot counts, maps, and examples.

The first slide is the drive-test route and the radio KPI tables with a share chart
for RSRP, RSRQ and SINR. The bad-spot summary is the second slide. Proposed azimuth,
E_tilt and M_tilt are the nearest whole degree. Half a degree rounds up. The workbook
uses the same rounding. The value stored in the dump stays as it was measured.
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
REPORT = OUTPUT / "Bogura_DriveTest_Report_v1.97.xlsx"
IDLE_NOTES = OUTPUT / "idle_vs_active_notes.json"
PLOTS = OUTPUT / "plots"
DESCRIPTIONS = OUTPUT / "bad_spot_descriptions.json"
SUMMARY_VERSION = "1.98"
DEST = OUTPUT / f"Bogura_DriveTest_Summary_v{SUMMARY_VERSION}.pptx"
EXCEL_VERSION = "1.97"
EXEC_HIST = {
    "RSRP": PLOTS / "ppt_exec_v198" / "active_rsrp_hist.png",
    "RSRQ": PLOTS / "ppt_exec_v198" / "active_rsrq_hist.png",
    "SINR": PLOTS / "ppt_exec_v198" / "active_sinr_hist.png",
}
CIRCLES = PLOTS / "bad_spot_circles.json"

# Neon navy to match Excel v1.91+ restyle
NAVY = RGBColor(0x0F, 0x29, 0x42)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5D, 0x6D, 0x7E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAPER = RGBColor(0xF7, 0xF9, 0xFB)
LINE = RGBColor(0xD5, 0xDD, 0xE4)
GOLD = RGBColor(0xB7, 0x86, 0x0B)
ORANGE = RGBColor(0xD3, 0x54, 0x00)
PURPLE = RGBColor(0x6C, 0x34, 0x83)


GREEN = RGBColor(0x1A, 0x7A, 0x4C)
TEAL = RGBColor(0x0F, 0x6C, 0x8C)

SOLUTION_COLOR = {
    "Azimuth": GOLD,
    "Tilt": GREEN,
    "4th sector, L900": ORANGE,
    "New 3-sector site": PURPLE,
    "On site EP Check": TEAL,
}
SOLUTION_LABEL = {
    "Azimuth": "Azimuth",
    "Tilt": "Tilt",
    "4th sector, L900": "4th Sector, L900",
    "New 3-sector site": "New 3-sector site",
    "On site EP Check": "On site EP Check",
}
SOLUTION_ORDER = (
    "Azimuth",
    "Tilt",
    "4th sector, L900",
    "New 3-sector site",
    "On site EP Check",
)


def _whole_degrees(text: str) -> str:
    """Nearest whole degree. Half a degree rounds up. A no-longer-visible change stays."""
    def repl(match: re.Match) -> str:
        whole = int(math.floor(float(match.group(1)) + 0.5))
        return f"{whole}°"

    text = re.sub(r"(\d+\.\d+)\s*°", repl, text)
    return re.sub(r"(E_tilt|M_tilt|azimuth) (\d+)° to \2°", r"\1 stays \2°", text)


def _slide_solution(cells: dict[str, str]) -> str:
    """Cover column G, with tilt-and-azimuth split by whether an azimuth is turned."""
    kind = cells.get("G", "")
    if kind in ("4th sector, L900", "New 3-sector site", "On site EP Check"):
        return kind
    if re.search(r"azimuth \d+° to \d+°", cells.get("H", "")):
        return "Azimuth"
    return "Tilt"


def _lead_line(final: str) -> str:
    lines = [re.sub(r"^\d+\.\s*", "", line).strip() for line in final.split("\n") if line.strip()]
    if not lines:
        return ""
    lead = lines[0]
    if lead.lower().startswith("on site ep check") and len(lines) > 1:
        lead = lines[1]
    if len(lead) > 260:
        lead = lead[:257].rstrip() + "..."
    return lead


def _example_groups(rows: dict[tuple[str, int], dict[str, str]]) -> list[dict]:
    """Up to five examples of each solution, taken from the cover sheet in spot order."""
    buckets: dict[tuple[str, str], list[int]] = {}
    for (mode, number), cells in rows.items():
        buckets.setdefault((mode, _slide_solution(cells)), []).append(number)
    groups = []
    for mode in ("Active", "Idle"):
        for solution in SOLUTION_ORDER:
            spots = sorted(buckets.get((mode, solution), []))
            if not spots:
                continue
            shown = spots if len(spots) <= 5 else spots[:5]
            groups.append({
                "solution": solution,
                "mode": mode,
                "spots": len(spots),
                "items": [
                    {"n": number, "why": _lead_line(rows[(mode, number)].get("H", ""))}
                    for number in shown
                ],
            })
    return groups

def _cover_rows() -> dict[tuple[str, int], dict[str, str]]:
    """Bad-spot rows for example slides.

    v1.97 Cover keeps an executive summary only (no per-spot list), so rows come
    from bad_spot_descriptions.json. Shape matches the former Cover columns.
    """
    solution_name = {
        "tune": "Tilt and azimuth",
        "fourth": "4th sector, L900",
        "new": "New 3-sector site",
        "ep": "On site EP Check",
    }
    found: dict[tuple[str, int], dict[str, str]] = {}
    if not DESCRIPTIONS.exists():
        return found
    for item in json.loads(DESCRIPTIONS.read_text()):
        mode = str(item.get("mode") or "")
        if mode not in ("Idle", "Active"):
            continue
        number = int(item.get("n") or 0)
        if number < 1:
            continue
        kind = str(item.get("kind") or "")
        badge = str(item.get("badge") or "")
        # Match former Cover column G: tilt/azimuth split later by _slide_solution.
        if kind == "tune":
            g = "Tilt and azimuth"
        else:
            g = solution_name.get(kind, badge or kind)
        final = str(item.get("suggestions") or item.get("proposed") or "")
        serving = str(item.get("serving") or "")
        # Prefer a short serving cell id when the serving block is multi-line.
        serving_cell = serving.split(",")[0].split(".")[-1].strip() if serving else ""
        found[(mode, number)] = {
            "A": mode,
            "B": f"Bad spot {number}",
            "C": str((item.get("changes") or [{}])[0].get("site") or ""),
            "D": str((item.get("changes") or [{}])[0].get("sector") or ""),
            "E": serving_cell or serving[:80],
            "F": "",
            "G": g,
            "H": final,
        }
    return found


def _map_path(mode: str, number: int) -> Path:
    if mode == "Idle":
        return PLOTS / f"excel_bad_spots_rsrp_spot_{number:02d}.png"
    return PLOTS / f"excel_bad_spots_rsrp_spot_active_{number:02d}.png"


def _idle_area_path(number: int) -> Path | None:
    path = PLOTS / f"excel_bad_spots_rsrp_spot_active_idle_{number:02d}.png"
    return path if path.exists() else None


def _idle_note(number: int) -> dict | None:
    if not IDLE_NOTES.exists():
        return None
    if not hasattr(_idle_note, "rows"):
        rows = {}
        for item in json.loads(IDLE_NOTES.read_text()):
            rows[int(item["n"])] = item
        _idle_note.rows = rows
    return _idle_note.rows.get(number)


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
    lines = str(text).split("\n")
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    _set_run(paragraph.add_run(), lines[0], size, color, bold)
    for line in lines[1:]:
        paragraph = frame.add_paragraph()
        paragraph.alignment = align
        paragraph.space_before = Pt(3)
        paragraph.space_after = Pt(0)
        _set_run(paragraph.add_run(), line, size, color, bold)
    return shape


def _footer(slide, page: int, pages: int) -> None:
    _text(
        slide, Inches(0.4), Inches(7.18), Inches(10), Inches(0.24),
        f"Bogura drive test  ·  Summary v{SUMMARY_VERSION}  ·  Excel v{EXCEL_VERSION}  ·  Web v1.89",
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


def _share_bar(slide, left: float, top: float, width: float, height: float, parts: list[tuple[float, str]]) -> None:
    """Horizontal share chart. parts are (percent, hex colour)."""
    used = left
    for percent, hex_color in parts:
        span = width * percent / 100.0
        if span < 0.02:
            continue
        red = int(hex_color[1:3], 16)
        green = int(hex_color[3:5], 16)
        blue = int(hex_color[5:7], 16)
        _box(slide, Inches(used), Inches(top), Inches(span), Inches(height), RGBColor(red, green, blue))
        used += span


def _report(prs: Presentation, pages: int) -> None:
    """Executive Summary: Active Mode histogram charts, notes, band mix, RSRP ranges."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, prs.slide_height, WHITE)
    _heading(slide, "Executive Summary")
    _text(
        slide, Inches(8.4), Inches(0.12), Inches(4.6), Inches(0.38),
        f"11–21 May 2026    ·    v{SUMMARY_VERSION}", 14, WHITE, align=PP_ALIGN.RIGHT,
    )

    # Compact DT route — Active first, Scanner second.
    _text(slide, Inches(0.28), Inches(0.66), Inches(4), Inches(0.20), "DT route", 12, NAVY, True)
    headers = ["Mode", "Source", "Plotted", "What is drawn"]
    values = [
        ["Active", "Borga_Active_mode_11Nov26", "288,308", "Robi M-PCell, strongest RSRP per location."],
        ["Scanner", "Borga_Idle_V2.0", "60,957", "Idle scanner — comparison only, not actual UE coverage."],
    ]
    widths = [1.15, 3.15, 1.35, 6.95]
    top = 0.88
    x = 0.28
    for header, width in zip(headers, widths):
        _box(slide, Inches(x), Inches(top), Inches(width), Inches(0.24), NAVY)
        _text(slide, Inches(x + 0.05), Inches(top + 0.01), Inches(width - 0.08), Inches(0.22), header, 10, WHITE, True)
        x += width
    for row_index, row in enumerate(values):
        y = top + 0.24 + row_index * 0.24
        x = 0.28
        fill = PAPER if row_index % 2 == 0 else WHITE
        for value, width in zip(row, widths):
            _box(slide, Inches(x), Inches(y), Inches(width), Inches(0.24), fill, LINE)
            _text(slide, Inches(x + 0.05), Inches(y + 0.01), Inches(width - 0.08), Inches(0.22), value, 10, INK)
            x += width
    _text(
        slide, Inches(0.28), Inches(1.62), Inches(12.7), Inches(0.28),
        "Active Mode only on this slide. Scanner coverage does not reflect actual UE coverage. "
        "Excel Bad Spot RSRP_Active: 99 Active spots then 26 Scanner spots. A bad spot is a 50 m bin with median RSRP below −120 dBm.",
        11, MUTED,
    )

    # Three Active Mode histogram charts (real stacked histograms, not share bars).
    _text(
        slide, Inches(0.28), Inches(1.92), Inches(8), Inches(0.22),
        "Active Mode — RSRP / RSRQ / SINR histogram charts", 13, NAVY, True,
    )
    hist_notes = [
        (
            "RSRP",
            EXEC_HIST["RSRP"],
            "Median −98.1 dBm. Poor (< −120): 1,077 samples, 0.4%. "
            "Most samples sit in −110 to −95 (46.4%) and −95 to 0 (40.6%).",
        ),
        (
            "RSRQ",
            EXEC_HIST["RSRQ"],
            "Median −9.7 dB. At or below −20: 0.2%. "
            "55.4% in −10 to 0; 43.2% in −17 to −10.",
        ),
        (
            "SINR",
            EXEC_HIST["SINR"],
            "Median 8.1 dB. At or below 0: 12.6%. "
            "About 41% of samples are at or above 10 dB.",
        ),
    ]
    chart_w = 4.05
    chart_h = 2.55
    gap = 0.18
    left0 = 0.28
    chart_top = 2.18
    for index, (title, path, note) in enumerate(hist_notes):
        if not path.exists():
            raise SystemExit(f"Missing Active histogram: {path}")
        left = left0 + index * (chart_w + gap)
        slide.shapes.add_picture(
            _jpeg(path, width=900),
            Inches(left), Inches(chart_top), Inches(chart_w), Inches(chart_h),
        )
        _text(
            slide, Inches(left), Inches(chart_top + chart_h + 0.04),
            Inches(chart_w), Inches(0.72), note, 11, INK,
        )

    # Band mix + RSRP range tables under the charts (full width of remaining space).
    table_top = 5.55
    _text(slide, Inches(0.28), Inches(table_top), Inches(4), Inches(0.20), "Active band mix", 12, NAVY, True)
    band_rows = [
        ("Band", "Samples", "Share"),
        ("L1800", "190,026", "65.9%"),
        ("L2100", "89,020", "30.9%"),
        ("L2600", "4,898", "1.7%"),
        ("L900", "4,364", "1.5%"),
    ]
    bw = [1.4, 1.35, 1.1]
    ty = table_top + 0.22
    for r_i, row in enumerate(band_rows):
        x = 0.28
        for value, width in zip(row, bw):
            fill = NAVY if r_i == 0 else (PAPER if r_i % 2 else WHITE)
            color = WHITE if r_i == 0 else INK
            _box(slide, Inches(x), Inches(ty), Inches(width), Inches(0.22), fill, LINE)
            _text(slide, Inches(x + 0.04), Inches(ty + 0.01), Inches(width - 0.06), Inches(0.20), value, 10, color, r_i == 0)
            x += width
        ty += 0.22

    _text(slide, Inches(4.6), Inches(table_top), Inches(4.5), Inches(0.20), "Active RSRP range (dBm)", 12, NAVY, True)
    range_rows = [
        ("Range", "Samples", "Share"),
        ("−95 to 0", "116,945", "40.6%"),
        ("−110 to −95", "133,790", "46.4%"),
        ("−115 to −110", "26,949", "9.3%"),
        ("−120 to −115", "9,547", "3.3%"),
        ("−150 to −120", "1,077", "0.4%"),
    ]
    rw = [1.55, 1.25, 1.0]
    ty = table_top + 0.22
    for r_i, row in enumerate(range_rows):
        x = 4.6
        for value, width in zip(row, rw):
            fill = NAVY if r_i == 0 else (PAPER if r_i % 2 else WHITE)
            color = WHITE if r_i == 0 else INK
            _box(slide, Inches(x), Inches(ty), Inches(width), Inches(0.22), fill, LINE)
            _text(slide, Inches(x + 0.04), Inches(ty + 0.01), Inches(width - 0.06), Inches(0.20), value, 10, color, r_i == 0)
            x += width
        ty += 0.22

    _text(
        slide, Inches(9.0), Inches(table_top), Inches(4.0), Inches(1.35),
        "Executive note\n"
        "Prefer Active Mode solutions for network action. "
        "Scanner maps show where idle best-server differs from the UE serving cell — "
        "do not treat Scanner bad spots as the primary action list.",
        11, INK,
    )
    _footer(slide, 1, pages)


def _kpi(prs: Presentation, spec: dict, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, spec["title"] + "  ·  plot and summary")
    _text(slide, Inches(0.36), Inches(0.66), Inches(12.6), Inches(0.42), spec["rule"], 13, MUTED)
    left = 0.28
    for path, caption in ((spec["active"], "Active"), (spec["idle"], "Idle")):
        with Image.open(path) as source:
            aspect = source.width / source.height
        height = 4.95
        width = height * aspect
        _text(slide, Inches(left), Inches(1.12), Inches(width), Inches(0.22), caption, 13, NAVY, True, PP_ALIGN.CENTER)
        slide.shapes.add_picture(_jpeg(path, width=900), Inches(left), Inches(1.36), Inches(width), Inches(height))
        left += width + 0.16
    panel = left
    panel_w = 13.05 - panel
    _text(slide, Inches(panel), Inches(1.12), Inches(panel_w), Inches(0.26), "ACTIVE", 12, NAVY, True)
    top = 1.38
    for label, value in spec["active_stats"]:
        _stat(slide, panel, top, panel_w, label, value)
        top += 0.28
    top += 0.08
    _text(slide, Inches(panel), Inches(top), Inches(panel_w), Inches(0.26), "IDLE", 12, NAVY, True)
    top += 0.28
    for label, value in spec["idle_stats"]:
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


def _group_count(groups: list[dict], mode: str, solution: str) -> tuple[int, int]:
    for group in groups:
        if group["mode"] == mode and group["solution"] == solution:
            return int(group["spots"]), len(group["items"])
    return 0, 0


def _bad_summary(prs: Presentation, page: int, pages: int, groups: list[dict]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Bad spot summary")
    headers = ["Solution", "Active spots", "Active shown", "Idle spots", "Idle shown"]
    labels = [
        ("Azimuth", "Azimuth", GOLD),
        ("Tilt, azimuth unchanged", "Tilt", GREEN),
        ("4th sector, L900", "4th sector, L900", ORANGE),
        ("New 3-sector site", "New 3-sector site", PURPLE),
        ("On site EP Check", "On site EP Check", TEAL),
    ]
    rows = []
    for label, key, color in labels:
        idle_n, idle_shown = _group_count(groups, "Idle", key)
        active_n, active_shown = _group_count(groups, "Active", key)
        rows.append((label, str(active_n), str(active_shown), str(idle_n), str(idle_shown), color))
    idle_all = sum(_group_count(groups, "Idle", key)[0] for _label, key, _color in labels)
    active_all = sum(_group_count(groups, "Active", key)[0] for _label, key, _color in labels)
    rows.append(("All bad spots", str(active_all), "", str(idle_all), "", NAVY))
    widths = [4.3, 1.9, 1.9, 2.1, 2.1]
    left0 = 0.45
    top = 0.78
    row_h = 0.56
    x = left0
    for header, width in zip(headers, widths):
        _box(slide, Inches(x), Inches(top), Inches(width), Inches(0.42), NAVY)
        _text(slide, Inches(x), Inches(top + 0.06), Inches(width), Inches(0.3), header, 13, WHITE, True, PP_ALIGN.CENTER)
        x += width
    for row_index, (label, active, active_shown, idle, idle_shown, color) in enumerate(rows):
        y = top + 0.42 + row_index * row_h
        values = [label, active, active_shown, idle, idle_shown]
        x = left0
        fill = RGBColor(0xEE, 0xF3, 0xF7) if row_index % 2 == 0 else WHITE
        if row_index == len(rows) - 1:
            fill = RGBColor(0xE8, 0xF0, 0xF6)
        for col, (value, width) in enumerate(zip(values, widths)):
            _box(slide, Inches(x), Inches(y), Inches(width), Inches(row_h), fill, LINE)
            if col == 0:
                _box(slide, Inches(x), Inches(y), Inches(0.1), Inches(row_h), color)
            _text(
                slide, Inches(x + (0.18 if col == 0 else 0)), Inches(y + 0.1), Inches(width), Inches(0.36),
                value, 15, INK, col == 0 or row_index == len(rows) - 1,
                PP_ALIGN.CENTER if col else PP_ALIGN.LEFT,
            )
            x += width
    idle_tune = _group_count(groups, "Idle", "Azimuth")[0] + _group_count(groups, "Idle", "Tilt")[0]
    active_tune = _group_count(groups, "Active", "Azimuth")[0] + _group_count(groups, "Active", "Tilt")[0]
    idle_ep, _shown = _group_count(groups, "Idle", "On site EP Check")
    active_ep, _shown = _group_count(groups, "Active", "On site EP Check")
    _text(
        slide, Inches(0.45), Inches(4.7), Inches(12.4), Inches(2.2),
        f"Tilt and azimuth together are {active_tune} active and {idle_tune} idle. The table splits that group. "
        "An azimuth row is a spot where an existing azimuth is turned. A tilt row leaves the azimuth as it is.\n"
        f"On site EP Check is {active_ep} active and {idle_ep} idle. The nearest site already faces the red road, "
        "within 12°, but that sector is missing from the log while another sector of the same site is on air. "
        "The azimuth is not changed. The site is checked.\n"
        "A group of 5 or fewer shows every spot. A larger group shows the first 5. "
        "When a spot changes more than one cell, each action is its own numbered line.\n"
        "The example box uses the same four blocks as the Excel sheet: Serving Cells details, "
        "Bad Spot Details, Root Cause Analysis, and Suggestions. "
        "A turned sector keeps its current arm. The proposed azimuth is a blue dashed arm labeled with the new degree. "
        "On the example slides, a tenth of a degree is shown as the nearest whole degree. 0.5 rounds up.",
        14, INK,
    )
    _footer(slide, page, pages)


def _block_heights(blocks: tuple, panel_w: float, size: int) -> list[float]:
    """Height of each label-plus-body block, including wrapped lines."""
    chars = max(40, int(panel_w * 72 / (size * 0.50)))
    heights = []
    for _label, value in blocks:
        paragraphs = str(value).split("\n") or [""]
        lines = 0
        for paragraph in paragraphs:
            length = max(1, len(paragraph))
            lines += max(1, (length + chars - 1) // chars)
        body = (lines * size * 1.15 + max(0, len(paragraphs) - 1) * 3) / 72
        heights.append(0.20 + body + 0.04)
    return heights


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
    idle_picture = _idle_area_path(item["n"]) if group["mode"] == "Active" else None
    with Image.open(picture) as source:
        aspect = source.width / source.height
    idle_aspect = None
    if idle_picture is not None:
        with Image.open(idle_picture) as source:
            idle_aspect = source.width / source.height
    image = _jpeg(picture)
    change_count = len((_description(group["mode"], item["n"]).get("changes") or []))
    table_rows = 2 + min(4, change_count or 1)
    max_w = 6.55
    note = _idle_note(item["n"]) if group["mode"] == "Active" else None
    if idle_aspect:
        # Side by side, the same width and the same geographic window.
        table_h = 0.28 * table_rows
        note_h = 0.95 if note else 0.0
        spare = 6.18 - table_h - note_h - 0.12
        gap = 0.08
        width = (max_w - gap) / 2.0
        idle_w = width
        height = width / aspect
        # Same rectangle as the active map. The idle picture's caption bar
        # makes its own aspect slightly taller, and the sheet shows one size.
        idle_h = height
        map_h = max(height, idle_h)
        if map_h > spare and map_h > 0:
            fitted = spare / map_h
            width *= fitted
            height *= fitted
            idle_w *= fitted
            idle_h *= fitted
    else:
        idle_h = 0.0
        note_h = 0.0
        table_h = 0.30 * table_rows
        max_h = min(4.70, 6.85 - 0.74 - table_h)
        width = max_w
        height = width / aspect
        if height > max_h:
            height = max_h
            width = height * aspect
    top = 0.64 if idle_aspect else 0.68
    slide.shapes.add_picture(image, Inches(0.16), Inches(top), Inches(width), Inches(height))
    if idle_picture is not None and idle_aspect:
        slide.shapes.add_picture(
            _jpeg(idle_picture), Inches(0.16 + width + 0.08), Inches(top), Inches(idle_w), Inches(idle_h),
        )
        top += max(height, idle_h) + 0.04
    else:
        top += height + 0.04
    if note:
        _text(slide, Inches(0.16), Inches(top), Inches(6.5), Inches(0.18), "Idle vs active", 11, color, True)
        _text(
            slide, Inches(0.16), Inches(top + 0.16), Inches(6.5), Inches(max(0.2, note_h - 0.16)),
            f"{note.get('coverage', '')}\n{note.get('serving', '')}",
            8, INK,
        )
        top += note_h
    _change_table(slide, group["mode"], item["n"], 0.16, top, 6.7, table_h)

    panel = 7.05
    panel_w = 5.95
    _text(
        slide, Inches(panel), Inches(0.62), Inches(panel_w), Inches(0.26),
        f"Proposed Solutions : {SOLUTION_LABEL[group['solution']]}", 13, color, True,
    )
    sections = _description(group["mode"], item["n"])
    blocks = (
        ("Serving Cells details", sections.get("serving") or record.get("E", "")),
        ("Bad Spot Details", sections.get("geography") or ""),
        ("Root Cause Analysis", sections.get("rca") or ""),
        ("Suggestions", sections.get("suggestions") or _whole_degrees(record.get("H", ""))),
    )
    # Long serving lists wrap inside the panel. Pick the largest body size that
    # still keeps all four blocks above the footer, so the slide shows the same
    # text as the Excel sheet.
    available = 7.05 - 0.90
    size = 10
    heights = _block_heights(blocks, panel_w, size)
    while size > 8 and sum(heights) > available:
        size -= 1
        heights = _block_heights(blocks, panel_w, size)
    top = 0.90
    for (label, value), height in zip(blocks, heights):
        _text(slide, Inches(panel), Inches(top), Inches(panel_w), Inches(0.20), label, 11, color, True)
        _text(slide, Inches(panel), Inches(top + 0.20), Inches(panel_w), Inches(height - 0.20), value, size, INK)
        top += height
    _footer(slide, page, pages)


def _change_table(slide, mode: str, number: int, left: float, top: float, width: float, height: float) -> None:
    """Existing versus New Plan for the azimuth and tilt rows on this spot."""
    record = _description(mode, number)
    changes = list(record.get("changes") or [])[:4]
    headers = ("Site", "Cell", "Sector", "H", "Az", "E", "M", "H", "Az", "E", "M")
    body = changes or [{"site": "No azimuth or tilt change", "cell": "", "sector": "", "height": None,
                        "azimuth": None, "et": None, "mt": None, "p_height": None, "p_azimuth": None,
                        "p_et": None, "p_mt": None}]
    rows = 2 + len(body)
    table_shape = slide.shapes.add_table(rows, len(headers), Inches(left), Inches(top), Inches(width), Inches(height))
    table = table_shape.table
    for row in table.rows:
        row.height = Inches(height / rows)
    widths = (1.25, 1.15, 0.95, 0.36, 0.40, 0.32, 0.32, 0.36, 0.40, 0.32, 0.32)
    scale = width / sum(widths)
    for index, col_width in enumerate(widths):
        table.columns[index].width = Inches(col_width * scale)

    def write(cell, text, size, bold, fill, color, center=False):
        cell.text = ""
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
        frame = cell.text_frame
        frame.word_wrap = False
        frame.margin_left = Emu(10000)
        frame.margin_right = Emu(10000)
        frame.margin_top = Emu(2000)
        frame.margin_bottom = Emu(2000)
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER if center else PP_ALIGN.LEFT
        _set_run(paragraph.add_run(), "" if text is None else str(text), size, color, bold)

    group = table.rows[0]
    for index in range(len(headers)):
        label = "Existing" if index == 3 else "New Plan" if index == 7 else ""
        write(group.cells[index], label, 9, True, NAVY, WHITE, True)
    # Merge the group labels. PowerPoint needs the cells spanned, not just the text.
    group.cells[3].merge(group.cells[6])
    group.cells[7].merge(group.cells[10])
    write(group.cells[3], "Existing", 9, True, NAVY, WHITE, True)
    write(group.cells[7], "New Plan", 9, True, NAVY, WHITE, True)
    for index, label in enumerate(headers):
        write(table.rows[1].cells[index], label, 8, True, PAPER, NAVY, True)
    keys = ("site", "cell", "sector", "height", "azimuth", "et", "mt", "p_height", "p_azimuth", "p_et", "p_mt")
    for row_index, item in enumerate(body):
        if item.get("site") == "No azimuth or tilt change":
            table.rows[row_index + 2].cells[0].merge(table.rows[row_index + 2].cells[10])
            write(table.rows[row_index + 2].cells[0], "No azimuth or tilt change", 9, False, WHITE, INK, False)
            continue
        for col_index, key in enumerate(keys):
            value = item.get(key)
            if key == "site" and str(value).startswith("New site"):
                value = "New site"
            if isinstance(value, float):
                value = str(int(round(value))) if abs(value - round(value)) < 0.05 else f"{value:.1f}"
            write(table.rows[row_index + 2].cells[col_index], value or "", 8, False, WHITE, INK, col_index >= 3)


def _description(mode: str, number: int) -> dict:
    """The four blocks written beside the same bad spot in the Excel workbook."""
    if not DESCRIPTIONS.exists():
        return {}
    if not hasattr(_description, "rows"):
        rows = {}
        for item in json.loads(DESCRIPTIONS.read_text()):
            rows[(item["mode"], int(item["n"]))] = item
        _description.rows = rows
    return _description.rows.get((mode, number), {})


def _annex(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _box(slide, 0, 0, prs.slide_width, prs.slide_height, NAVY)
    _box(slide, 0, 0, Inches(0.16), prs.slide_height, GOLD)
    _text(slide, Inches(0.7), Inches(2.15), Inches(12), Inches(0.9), "Annex", 60, WHITE, True)
    _text(
        slide, Inches(0.7), Inches(3.3), Inches(11.5), Inches(1.4),
        "The report ends on the previous slide. The slides after this one show what 12° off the red road means, "
        "the rounding rule, how a solution is chosen, how to read a map, and the check that every red sample inside the circle is included.",
        20, RGBColor(0xD6, 0xE4, 0xF0),
    )
    _text(slide, Inches(0.7), Inches(6.9), Inches(10), Inches(0.28), f"Bogura drive test  ·  Summary v{SUMMARY_VERSION}", 14, RGBColor(0xAE, 0xC6, 0xD6))
    _text(slide, Inches(11.2), Inches(6.9), Inches(1.6), Inches(0.28), f"{page}  /  {pages}", 14, RGBColor(0xAE, 0xC6, 0xD6), align=PP_ALIGN.RIGHT)


def _off_angle_png() -> BytesIO:
    """Idle Bad Spot 2: the angle at each site between the antenna and the red road."""
    import matplotlib

    matplotlib.use("Agg", force=False)
    import matplotlib.pyplot as plt
    from matplotlib.patches import Wedge

    sites = [
        {"name": "BOSDR07C", "dist": 0.80, "bearing": 346, "azimuth": 350, "off": 4, "note": "faces the road", "change": False},
        {"name": "BOSDR52A", "dist": 1.11, "bearing": 17, "azimuth": 355, "off": 22, "note": "outside the wedge", "change": False},
        {"name": "BOSDR76A", "dist": 1.94, "bearing": 46, "azimuth": 40, "off": 6, "note": "faces the road", "change": False},
    ]

    def from_road(dist: float, bearing: float) -> tuple[float, float]:
        back = math.radians((bearing + 180.0) % 360.0)
        return dist * math.sin(back), dist * math.cos(back)

    def step(azimuth: float, length: float) -> tuple[float, float]:
        rad = math.radians(azimuth % 360.0)
        return length * math.sin(rad), length * math.cos(rad)

    fig = plt.figure(figsize=(7.7, 6.28), dpi=160, facecolor="#f7f9fb")
    ax = fig.add_axes([0.02, 0.08, 0.96, 0.90])
    ax.set_facecolor("#f7f9fb")
    ax.set_aspect("equal")
    ax.set_xlim(-3.05, 1.85)
    ax.set_ylim(-2.85, 1.15)
    ax.axis("off")

    for site in sites:
        sx, sy = from_road(site["dist"], site["bearing"])
        reach = site["dist"] * 0.92
        ax.add_patch(Wedge(
            (sx, sy), reach,
            90 - (site["azimuth"] + 12),
            90 - (site["azimuth"] - 12),
            facecolor="#c5ddc0", edgecolor="#6fa36a", linewidth=0.7, alpha=0.72, zorder=2,
        ))
        ax.plot([sx, 0], [sy, 0], color="#8a8175", lw=1.2, ls=(0, (4, 3)), zorder=3)
        dx, dy = step(site["azimuth"], 0.58)
        ax.annotate(
            "", xy=(sx + dx, sy + dy), xytext=(sx, sy),
            arrowprops=dict(arrowstyle="-|>", color="#1f4e79", lw=2.3, mutation_scale=14), zorder=5,
        )
        if site["change"]:
            gx, gy = step(site["new_az"], 0.82)
            ax.annotate(
                "", xy=(sx + gx, sy + gy), xytext=(sx, sy),
                arrowprops=dict(arrowstyle="-|>", color="#c48a12", lw=2.3, mutation_scale=14), zorder=5,
            )
        delta = (site["bearing"] - site["azimuth"] + 180) % 360 - 180
        color = "#c0392b" if site["off"] > 12 else "#1e7a34"
        theta_az = 90 - site["azimuth"]
        theta_road = 90 - (site["azimuth"] + delta)
        ax.add_patch(Wedge(
            (sx, sy), 0.36, min(theta_az, theta_road), max(theta_az, theta_road),
            width=0.075, facecolor=color, edgecolor=color, zorder=6,
        ))
        mid = site["azimuth"] + delta / 2.0
        lx, ly = step(mid, 0.46)
        # A small gap sits on the antenna arrow, so the number sits beside the arc.
        side = 90 if delta >= 0 else -90
        ox, oy = step(mid + side, 0.28)
        ax.text(
            sx + lx + ox, sy + ly + oy, f"{site['off']}°", ha="center", va="center", fontsize=10,
            fontweight="bold", color=color, zorder=7,
            bbox=dict(boxstyle="round,pad=0.12", fc="white", ec=color, lw=0.6),
        )
        ax.scatter([sx], [sy], s=32, color="#1d3557", zorder=8)

    labels = [
        ("BOSDR07C", 0.80, 346, (1.05, -0.35), "4°\nfaces the road\nnot in the log", "#1e7a34"),
        ("BOSDR52A", 1.11, 17, (-0.95, -2.15), "22°\noutside the wedge", "#c0392b"),
        ("BOSDR76A", 1.94, 46, (-2.55, -0.85), "6°\nfaces the road", "#1e7a34"),
    ]
    for name, dist, bearing, (tx, ty), note, color in labels:
        sx, sy = from_road(dist, bearing)
        ax.annotate(
            f"{name}\n{note}", xy=(sx, sy - 0.08), xytext=(tx, ty),
            fontsize=8, color="#1d3557", fontweight="bold", ha="center", va="center", zorder=9,
            bbox=dict(boxstyle="round,pad=0.28", fc="white", ec=color, lw=1.1),
            arrowprops=dict(arrowstyle="-", color=color, lw=0.7),
        )

    ax.plot([0, 0], [-0.14, 0.14], color="#c0392b", lw=6, solid_capstyle="round", zorder=6)
    ax.annotate(
        "red road\n(the bad samples)", xy=(0.02, 0.14), xytext=(0.42, 0.72),
        fontsize=9, color="#7b1e16", fontweight="bold", ha="left", va="center", zorder=9,
        arrowprops=dict(arrowstyle="-", color="#c0392b", lw=0.8),
    )
    ax.annotate(
        "", xy=(1.55, 0.95), xytext=(1.55, 0.55),
        arrowprops=dict(arrowstyle="-|>", color="#222", lw=1.3, mutation_scale=11),
    )
    ax.text(1.55, 1.02, "N", ha="center", va="bottom", fontsize=11, fontweight="bold", color="#222")
    ax.plot([-2.85, -1.85], [-2.62, -2.62], color="#222", lw=1.8)
    ax.plot([-2.85, -2.85], [-2.68, -2.56], color="#222", lw=1.2)
    ax.plot([-1.85, -1.85], [-2.68, -2.56], color="#222", lw=1.2)
    ax.text(-2.35, -2.74, "1 km", ha="center", va="top", fontsize=8, color="#222")
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=160, facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf


def _annex_off_angle(prs: Presentation, page: int, pages: int) -> None:
    """One annex slide: the 12° angle, drawn for Idle Bad Spot 2."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  12° off the red road")
    slide.shapes.add_picture(_off_angle_png(), Inches(0.18), Inches(0.72), Inches(7.7), Inches(6.28))
    left = 8.02
    width = 4.98
    cards = [
        (0.76, 1.52, "The angle",
         "Stand at the site. The blue arrow is where the antenna points. The dashed line is the direction from that site to the middle of the red road. The red road is the bad samples. The gap between the two lines is the angle."),
        (2.38, 1.42, "The green wedge",
         "The wedge is 12° to either side of the antenna. A dashed line inside it means the antenna already faces the red road. A dashed line outside it is the case where an azimuth would turn toward the red road."),
        (3.90, 2.95, "Idle Bad Spot 2",
         "BOSDR07C is the nearest sector, 0.8 km away. Azimuth 350°. The red road is at 346°. The gap is 4°, inside the wedge.\n\n"
         "The idle log does not decode BOSDR07C. The other sectors of BOSDR07 are on air. The solution is On site EP Check: confirm on site that the antenna is really at 350° and that the sector is radiating. The cell that serves the red road is BOSDR96L21C.\n\n"
         "BOSDR52A is 22° outside the wedge. BOSDR76A is 6° inside it. Those angles stay on this sketch. The spot is no longer a turn of BOSDR52A."),
    ]
    for top, height, title, body in cards:
        _box(slide, Inches(left), Inches(top), Inches(width), Inches(height), PAPER, LINE)
        _text(slide, Inches(left + 0.16), Inches(top + 0.08), Inches(width - 0.32), Inches(0.32), title, 15, NAVY, True)
        _text(slide, Inches(left + 0.16), Inches(top + 0.42), Inches(width - 0.32), Inches(height - 0.52), body, 12, INK)
    _footer(slide, page, pages)


def _annex_rounding(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  azimuth and tilt are whole degrees")
    rows = [
        ("What is rounded", "The proposed azimuth, E_tilt and M_tilt. The measured value in the dump is not changed."),
        ("Which way", "The nearest whole degree. 0.5 rounds up. 1.8° is shown as 2°. 0.9° is shown as 1°. 1.2° is shown as 1°."),
        ("A very small change", "If the nearest degree is the same as the degree it already has, the line says the tilt or azimuth stays."),
        ("Where it applies", "Every example in this deck, and the proposed value in the Excel workbook. The dump value is unchanged."),
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
        (GOLD, "Azimuth", "A sector inside 2.5 km can be turned when it does not already face the spot. Beside a new 4th sector, the closest existing sector is turned away until the gap is 30° to 60°. On a long blank, the nearest other sector can turn toward the far end."),
        (GREEN, "Tilt", "Downtilt of an existing sector is at most 2°. A far cell is downtilted when it serves a red sample the new sector does not cover. Idle L2600 is not downtilted. The nearest site is tuned instead. A far cell on the piece the new sector already covers is left in place."),
        (ORANGE, "4th sector", "Used when the spot sits in the null inside 1.5 km, when one or two sites inside 2 km leave it in the null, when the required downtilt would pass 2°, or when a host inside 3 km can aim at the spot. The closest existing sector moves when it would sit within 30° of the new one."),
        (PURPLE, "New site", "Used only when the nearest existing site is farther than 3 km. One L900 sector aims at the bad spot. The other two aim at the higher population around the site. The gap between any two sectors is at least 40°."),
        (TEAL, "On site EP Check", "Used when the nearest site already faces the red road, within 12°, and that sector is missing from the log while another sector of the same site is on air. The azimuth is not changed. The site is checked."),
    ]
    for index, (color, title, body) in enumerate(cards):
        col = index % 2
        row = index // 2
        left = 0.35 + col * 6.5
        top = 0.72 + row * 2.08
        _box(slide, Inches(left), Inches(top), Inches(6.25), Inches(1.96), PAPER, LINE)
        _box(slide, Inches(left), Inches(top), Inches(6.25), Inches(0.08), color)
        _text(slide, Inches(left + 0.18), Inches(top + 0.16), Inches(5.9), Inches(0.32), title, 16, NAVY, True)
        _text(slide, Inches(left + 0.18), Inches(top + 0.5), Inches(5.9), Inches(1.38), body, 13, INK)
    _footer(slide, page, pages)


def _annex_read(prs: Presentation, page: int, pages: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  how to read an example")
    items = [
        ("Red on the road", "The poor samples. The length beside the red is that red stretch, and it sits off the samples."),
        ("Orange circle", "The bad spot. A blank with no plot inside the circle is part of the same spot."),
        ("Yellow dashed line", "The cell that served the red samples. It starts at the end of that sector. It is not the proposed sector. The length is from the arm tip, so it is shorter than the site-centre distance."),
        ("Tilt on the dashed line", "When that cell is an overshooter, the line also carries the E_tilt change."),
        ("Orange arm", "A proposed 4th L900 sector, drawn from the existing site toward the spot. A turned sector keeps its current arm. The proposed azimuth is a blue dashed arm, labeled with the new degree and the size of the turn."),
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
        "Every red sample inside the orange circle is still checked. A far cell is downtilted when the new sector does not cover that piece. Idle L2600 is not downtilted. A far cell on the piece the new sector already aims at is left in place.",
        16, INK,
    )
    rows = [
        ("Idle 12", "BOSBG03A L2100 serves the separate 50 m sample, about 64° off the new sector, so it is still downtilted, E_tilt 2° to 4°. The map line reads 5,853 m because it starts at the arm tip. BOSBG26B L2600 is on the piece the 4th sector aims at, so it is not downtilted."),
        ("Idle 1", "BOSDR56A L2600 serves from 6,563 m and is not downtilted. BOGBT18C at 1.1 km is uptilted from 3° to 2°. Its azimuth already faces the road."),
        ("Active 5", "GBSGT12C L1800 serves the piece the new sector covers, so it is not downtilted. GBGBD15A turns from 25° to 41° toward the far end of the blank."),
        ("Active 7", "JYKLI02C and JYSDR26A serve red samples no new sector covers, so both are still downtilted. The solution numbers the three cells."),
    ]
    top = 1.7
    for label, body in rows:
        _box(slide, Inches(0.45), Inches(top), Inches(12.4), Inches(1.15), PAPER, LINE)
        _text(slide, Inches(0.62), Inches(top + 0.14), Inches(2.8), Inches(0.85), label, 15, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        _text(slide, Inches(3.5), Inches(top + 0.12), Inches(9.1), Inches(0.9), body, 14, INK, anchor=MSO_ANCHOR.MIDDLE)
        top += 1.25
    _footer(slide, page, pages)


def _annex_workbook(prs: Presentation, page: int, pages: int, groups: list[dict]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Annex  ·  where the full set is")
    full = []
    partial = []
    for group in groups:
        name = f"{group['mode']} {SOLUTION_LABEL[group['solution']]}"
        shown = len(group["items"])
        if shown == group["spots"]:
            full.append(f"{name} ({shown})")
        else:
            partial.append(f"{name} shows {shown} of {group['spots']}")
    deck = "The examples above. "
    if full:
        deck += "Full sets: " + ", ".join(full) + ". "
    if partial:
        deck += "Samples: " + ", ".join(partial) + "."
    rows = [
        ("Excel v1.97 / PPT v1.98", "One Bad Spot sheet: Active Mode first, Scanner second. Cover is an executive summary. PPT Executive Summary shows Active-only RSRP/RSRQ/SINR histogram charts plus band mix / RSRP ranges."),
        ("This deck", deck),
        ("Web v1.89", "Feeder draws the idle feeder swaps. Click a site for the sectors. A three-sector rotation uses red, cyan and violet. A two-sector swap stays red and cyan, with the third arm yellow. Site names stay white with a dark edge."),
        ("RSRQ and SINR", "The coverage maps are in this deck. The workbook no longer has separate RSRQ or SINR bad-spot sheets. Poor RSRQ is at or below −20 dB. Poor SINR is at or below 0 dB."),
    ]
    top = 0.95
    for label, body in rows:
        _box(slide, Inches(0.4), Inches(top), Inches(12.5), Inches(1.35), PAPER, LINE)
        _text(slide, Inches(0.58), Inches(top + 0.18), Inches(2.5), Inches(1.0), label, 16, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        _text(slide, Inches(3.2), Inches(top + 0.16), Inches(9.4), Inches(1.05), body, 15, INK, anchor=MSO_ANCHOR.MIDDLE)
        top += 1.48
    _footer(slide, page, pages)


def _feeder(prs: Presentation, page: int, pages: int) -> None:
    """One slide for the idle feeder swaps, with the three-sector rotation beside the list."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _heading(slide, "Feeder swap  ·  idle mode")
    _text(
        slide, Inches(0.36), Inches(0.68), Inches(12.6), Inches(0.48),
        "Two sectors count when each is heard in the other's area. Three sectors count when that hearing goes around the site. A sample within 500 m is left out.",
        14, MUTED,
    )
    image = PLOTS / "feeder" / "BOSDR02.png"
    if not image.exists():
        raise SystemExit(f"Missing {image}")
    slide.shapes.add_picture(str(image), Inches(0.28), Inches(1.22), Inches(7.15), Inches(5.7))
    payload = json.loads((OUTPUT / "feeder_swaps.json").read_text())
    top = 1.22
    for item in payload:
        cycle = item.get("kind") == "cycle"
        _box(slide, Inches(7.6), Inches(top), Inches(5.35), Inches(0.84), PAPER, LINE)
        if cycle:
            _box(slide, Inches(7.6), Inches(top), Inches(0.08), Inches(0.84), RGBColor(0xD5, 0x00, 0xF9))
        label = "3-sector rotation" if cycle else "2-sector swap"
        _text(slide, Inches(7.82), Inches(top + 0.06), Inches(5.0), Inches(0.28), f"{item['site']}   ·   {label}", 13, NAVY, True)
        _text(slide, Inches(7.82), Inches(top + 0.36), Inches(5.0), Inches(0.4), item["title"], 12, INK)
        top += 0.9
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
    groups = _example_groups(rows)
    missing = []
    for group in groups:
        for item in group["items"]:
            if (group["mode"], item["n"]) not in rows:
                missing.append(f"{group['mode']} {item['n']}")
            if not _map_path(group["mode"], item["n"]).exists():
                missing.append(f"map {group['mode']} {item['n']}")
    for spec in KPI_SLIDES:
        for key in ("idle", "active"):
            if not spec[key].exists():
                missing.append(spec[key].name)
    for name, path in EXEC_HIST.items():
        if not path.exists():
            missing.append(f"exec hist {name}: {path.name}")
    if missing:
        raise SystemExit("Missing example data: " + ", ".join(missing))

    example_count = sum(len(group["items"]) for group in groups)
    # cover, 3 KPI, bad-spot summary, examples, feeder, annex, 6 annex detail slides
    pages = 1 + len(KPI_SLIDES) + 1 + example_count + 1 + 1 + 6
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = f"Bogura Drive Test Summary v{SUMMARY_VERSION}"
    prs.core_properties.subject = "Route, radio KPIs, bad-spot summary, and examples"
    _report(prs, pages)
    page = 2
    _bad_summary(prs, page, pages, groups)
    page += 1
    for spec in KPI_SLIDES:
        _kpi(prs, spec, page, pages)
        page += 1
    for group in groups:
        for ordinal, item in enumerate(group["items"], start=1):
            _example(prs, group, item, rows[(group["mode"], item["n"])], ordinal, page, pages)
            page += 1
    _feeder(prs, page, pages)
    page += 1
    _annex(prs, page, pages)
    page += 1
    _annex_off_angle(prs, page, pages)
    page += 1
    _annex_rounding(prs, page, pages)
    page += 1
    _annex_rules(prs, page, pages)
    page += 1
    _annex_read(prs, page, pages)
    page += 1
    _annex_circle(prs, page, pages)
    page += 1
    _annex_workbook(prs, page, pages, groups)
    if page != pages:
        raise SystemExit(f"Page count {page} does not match {pages}")
    OUTPUT.mkdir(parents=True, exist_ok=True)
    prs.save(DEST)
    print(f"Wrote {DEST} ({DEST.stat().st_size / 1e6:.1f} MB, {len(groups)} groups, {pages} slides)")
    return DEST


if __name__ == "__main__":
    build()
