#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v2.05 — Excel-openable fix.

v2.03/v2.04 Cover Active histogram PNG embeds made Excel report the file as
corrupt. Restore Cover drawing1 (native legend-coloured charts + route image)
from the last openable package (v1.98), keep Cover text, drop the bad media.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as XET

SRC_BROKEN = Path("output/Bogura_DriveTest_Report_v2.04.xlsx")
SRC_GOOD = Path("output/Bogura_DriveTest_Report_v1.98.xlsx")
DEST = Path("output/Bogura_DriveTest_Report_v2.05.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")

BAD_MEDIA = (
    "xl/media/image_cover_active_rsrp.png",
    "xl/media/image_cover_active_rsrq.png",
    "xl/media/image_cover_active_sinr.png",
)


def publish() -> Path:
    if not SRC_GOOD.exists():
        raise SystemExit(f"Missing openable source {SRC_GOOD}")
    if not SRC_BROKEN.exists():
        raise SystemExit(f"Missing broken source {SRC_BROKEN}")

    with zipfile.ZipFile(SRC_BROKEN) as broken, zipfile.ZipFile(SRC_GOOD) as good:
        names = [n for n in broken.namelist() if n not in BAD_MEDIA]
        parts = {n: broken.read(n) for n in names}
        # Restore Cover drawing + rels from the last Excel-openable package.
        parts["xl/drawings/drawing1.xml"] = good.read("xl/drawings/drawing1.xml")
        parts["xl/drawings/_rels/drawing1.xml.rels"] = good.read(
            "xl/drawings/_rels/drawing1.xml.rels"
        )

    # Cover label: keep the legend note if present; otherwise leave as-is.
    cover = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    if "map legend colours, same as PPT" not in cover:
        cover2, n = re.subn(
            r"(>Active histograms \(plotted Robi rows\)<)",
            r">Active histograms (map legend colours, same as PPT)<",
            cover,
            count=1,
        )
        if n == 1:
            cover = cover2
    parts["xl/worksheets/sheet1.xml"] = cover.encode("utf-8")

    tmp = DEST.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            if name in BAD_MEDIA:
                continue
            zf.writestr(name, parts[name], compress_type=zipfile.ZIP_DEFLATED)
    tmp.replace(DEST)
    shutil.copy2(DEST, LATEST)
    print(f"Wrote {DEST} ({DEST.stat().st_size / 1e6:.1f} MB)", flush=True)
    return DEST


def verify(path: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(path) as zf, zipfile.ZipFile(SRC_GOOD) as good:
        for name in BAD_MEDIA:
            if name in zf.namelist():
                problems.append(f"bad media still present: {name}")
        if zf.read("xl/drawings/drawing1.xml") != good.read("xl/drawings/drawing1.xml"):
            problems.append("drawing1.xml does not match openable v1.98")
        if zf.read("xl/drawings/_rels/drawing1.xml.rels") != good.read(
            "xl/drawings/_rels/drawing1.xml.rels"
        ):
            problems.append("drawing1.xml.rels does not match openable v1.98")
        # Cover still has Active chart graphicFrames (rId5–7).
        drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
        for rid in ("rId5", "rId6", "rId7"):
            if f'r:id="{rid}"' not in drawing:
                problems.append(f"Active chart {rid} missing after restore")
            if f'r:embed="{rid}"' in drawing:
                problems.append(f"Active chart {rid} still a pic embed")
        # Chart parts still present.
        for chart in ("xl/charts/chart5.xml", "xl/charts/chart6.xml", "xl/charts/chart7.xml"):
            if chart not in zf.namelist():
                problems.append(f"missing {chart}")
        # Spot-check other sheets unchanged vs broken source (data intact).
        with zipfile.ZipFile(SRC_BROKEN) as broken:
            for sheet in (
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet6.xml",
                "xl/worksheets/sheet10.xml",
            ):
                if broken.read(sheet) != zf.read(sheet):
                    problems.append(f"{sheet} changed unexpectedly")
        # Parse drawing XML.
        try:
            XET.fromstring(zf.read("xl/drawings/drawing1.xml"))
        except XET.ParseError as exc:
            problems.append(f"drawing1 parse: {exc}")
        bad = zf.testzip()
        if bad:
            problems.append(f"zip CRC fail: {bad}")
    return problems


if __name__ == "__main__":
    dest = publish()
    problems = verify(dest)
    if problems:
        raise SystemExit("v2.05 verify failed: " + "; ".join(problems))
    print(
        "v2.05 verified: Cover Active native charts restored from v1.98; "
        "corrupt PNG embeds removed."
    )
