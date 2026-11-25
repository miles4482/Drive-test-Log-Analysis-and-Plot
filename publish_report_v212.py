#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v2.12.

New-site azimuths on every Active bad spot that plans a new site:

1. Two sectors aim at the blank and the poor road. Different directions get
   one sector each. One corridor gets two sectors 60° apart.
2. The third sector aims at population.
3. Neighbor L900 sectors that face the new site are retuned in azimuth and tilt.

Redraws those 12 maps and rewrites Active advice. Cover stays the v2.11 sheet.
"""
from __future__ import annotations

import io
import json
import re
import shutil
import zipfile
from pathlib import Path

import generate_excel_report as ger
from generate_excel_report import PLOTS_DIR, REPORT_VERSION, best_carrier
from merge_and_analyze import (
    MAP_CIRCLE_PIXELS,
    MAP_RSRP_COLORS,
    MAP_RSRP_LABELS,
    find_bad_spots,
    load_active_merged,
    plot_bad_spot_serial_maps,
    remember_seen_sectors,
    rsrp_map_class,
    select_active_report,
)

FIX = {3, 4, 10, 11, 15, 16, 20, 26, 59, 70, 79, 88}
SRC = Path("output/Bogura_DriveTest_Report_v2.11.xlsx")
DEST = Path(f"output/Bogura_DriveTest_Report_v{REPORT_VERSION}.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")
CIRCLES = PLOTS_DIR / "bad_spot_circles.json"

DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def _drop_spot_circles(drawing: str, number: int) -> str:
    pattern = re.compile(rf"^Circle {number}(?:[a-z]+)?$")

    def keep(match: re.Match) -> str:
        block = match.group(0)
        if 'prst="ellipse"' not in block:
            return block
        name = re.search(r'name="([^"]+)"', block)
        if name and pattern.match(name.group(1)):
            return ""
        return block

    return re.sub(r"<oneCellAnchor>.*?</oneCellAnchor>", keep, drawing, flags=re.S)


def _ensure_drawing_ns(parts: dict[str, bytes]) -> int:
    fixed = 0
    for name, body in list(parts.items()):
        if not name.startswith("xl/drawings/drawing") or not name.endswith(".xml"):
            continue
        text = body.decode("utf-8")
        match = re.match(r"(<xdr:wsDr\b)([^>]*)>", text)
        if match is None:
            match = re.match(r"(<wsDr\b)([^>]*)>", text)
        if match is None:
            continue
        attrs = match.group(2)
        extra = ""
        if "xmlns:a=" not in attrs:
            extra += f' xmlns:a="{DML}"'
        if "xmlns:r=" not in attrs:
            extra += f' xmlns:r="{REL}"'
        if "xmlns:xdr=" not in attrs and match.group(1).startswith("<xdr:"):
            extra += f' xmlns:xdr="{XDR}"'
        if not extra:
            continue
        text = text[: match.end(2)] + extra + text[match.end(2) :]
        parts[name] = text.encode("utf-8")
        fixed += 1
    return fixed


def main() -> None:
    if REPORT_VERSION != "2.12":
        raise SystemExit(f"Expected REPORT_VERSION 2.12, got {REPORT_VERSION}")
    if not SRC.exists():
        raise SystemExit(f"Missing {SRC}")
    shutil.copy2(SRC, DEST)
    active = select_active_report(load_active_merged())
    remember_seen_sectors("Active", active)
    track = best_carrier(active)
    spots = find_bad_spots(track, "RSRP").head(97)
    print(f"Redrawing {len(FIX)} new-site maps", flush=True)
    plot_bad_spot_serial_maps(
        track, spots, PLOTS_DIR, "excel_bad_spots_rsrp_spot_active",
        "RSRP", rsrp_map_class(track["RSRP"]), MAP_RSRP_LABELS, MAP_RSRP_COLORS,
        basemap_style="satellite", route_casing=True, mode="Active", only=FIX,
    )

    with zipfile.ZipFile(DEST) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    drawings = ger._sheet_drawings(parts)
    drawing_name = drawings["Bad Spot RSRP_Active"]
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
        if re.search(r'name="Idle area \d+"', block) or re.search(r'name="Scanner Image \d+"', block):
            continue
        row = int(re.search(r"<row>(\d+)</row>", block).group(1))
        pictures.append((row, block))
    pictures.sort()
    serial = pictures[1:]
    if len(serial) != 97:
        raise SystemExit(f"Expected 97 serial maps, found {len(serial)}")

    from PIL import Image

    next_shape = max((int(i) for i in re.findall(r'<cNvPr id="(\d+)"', drawing)), default=0) + 1
    extras = []
    index = json.loads(CIRCLES.read_text())
    for number in sorted(FIX):
        png = PLOTS_DIR / f"excel_bad_spots_rsrp_spot_active_{number:02d}.png"
        geom = MAP_CIRCLE_PIXELS[str(png.resolve())]
        _row, block = serial[number - 1]
        embed = re.search(r'r:embed="([^"]+)"', block).group(1)
        image = Image.open(png).convert("RGB")
        buf = io.BytesIO()
        image.save(buf, "JPEG", quality=74, optimize=True)
        parts[rid_media[embed]] = buf.getvalue()
        drawing = _drop_spot_circles(drawing, number)
        origin = re.search(
            r"<from><col>(\d+)</col><colOff>(\d+)</colOff><row>(\d+)</row><rowOff>(\d+)</rowOff></from><ext cx=\"(\d+)\" cy=\"(\d+)\"",
            block,
        )
        col, col_off, row, row_off, pic_cx, pic_cy = (int(v) for v in origin.groups())
        img_w, img_h = geom["size"]
        for circle in geom["circles"]:
            extras.append(ger._circle_anchor(circle, next_shape, col, col_off, row, row_off, pic_cx, pic_cy, img_w, img_h))
            next_shape += 1
        index[png.name] = {
            "width": geom["size"][0],
            "height": geom["size"][1],
            "circles": [
                {
                    "x": c["x"], "y": c["y"], "w": c["w"], "h": c["h"],
                    "dotted": bool(c.get("dotted")), "name": c.get("name") or "",
                }
                for c in geom["circles"]
            ],
        }
        print(f"  spot {number}: {len(geom['circles'])} circles", flush=True)
    drawing = drawing.replace("</wsDr>", "".join(extras) + "</wsDr>")
    parts[drawing_name] = drawing.encode("utf-8")
    tmp = DEST.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name])
    tmp.replace(DEST)
    CIRCLES.write_text(json.dumps(index))
    print("Rewriting Active advice", flush=True)
    ger.write_bad_spot_advice(src=DEST, sheets=("Bad Spot RSRP_Active",))

    with zipfile.ZipFile(DEST) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    _ensure_drawing_ns(parts)
    with zipfile.ZipFile(SRC) as z_src:
        parts["xl/worksheets/sheet1.xml"] = z_src.read("xl/worksheets/sheet1.xml")
    tmp = DEST.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name])
    tmp.replace(DEST)
    shutil.copy2(DEST, LATEST)
    print(f"Wrote {DEST} ({DEST.stat().st_size / 1e6:.1f} MB)", flush=True)


if __name__ == "__main__":
    main()
