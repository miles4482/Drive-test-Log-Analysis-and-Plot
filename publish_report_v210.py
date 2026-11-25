#!/usr/bin/env python3
"""Redraw the Active bad spots whose circle missed a blank end or a poor sample.

Checked all 97. These 25 failed that check on v2.09 and are replaced in v2.10.
The other 72 already had every poor sample and every blank inside a circle.
"""
from __future__ import annotations

import io
import json
import re
import shutil
import zipfile
from pathlib import Path

import generate_excel_report as ger
from generate_excel_report import PLOTS_DIR, REPORT_VERSION
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
from generate_excel_report import best_carrier

FIX = {
    6, 8, 9, 14, 16, 18, 19, 27, 28, 29, 30, 31, 40, 42, 43,
    58, 60, 65, 79, 81, 82, 84, 90, 93, 95,
}
SRC = Path("output/Bogura_DriveTest_Report_v2.09.xlsx")
DEST = Path(f"output/Bogura_DriveTest_Report_v{REPORT_VERSION}.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")
CIRCLES = PLOTS_DIR / "bad_spot_circles.json"


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


def main() -> None:
    if REPORT_VERSION != "2.10":
        raise SystemExit(f"Expected REPORT_VERSION 2.10, got {REPORT_VERSION}")
    shutil.copy2(SRC, DEST)
    active = select_active_report(load_active_merged())
    remember_seen_sectors("Active", active)
    track = best_carrier(active)
    spots = find_bad_spots(track, "RSRP").head(97)
    print(f"Redrawing {len(FIX)} Active maps", flush=True)
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
    shutil.copy2(DEST, LATEST)
    CIRCLES.write_text(json.dumps(index))
    print(f"Wrote {DEST}", flush=True)


if __name__ == "__main__":
    main()
