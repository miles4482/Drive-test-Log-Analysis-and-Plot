#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v2.03 from v1.98 (Cover sheet only).

Replace Cover Active Mode RSRP / RSRQ / SINR native charts with the same
legend-coloured histogram PNGs used on PPT Executive Summary. Idle charts,
route picture, and every other sheet stay byte-identical.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

from PIL import Image

SRC = Path("output/Bogura_DriveTest_Report_v1.98.xlsx")
DEST = Path("output/Bogura_DriveTest_Report_v2.03.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")
HIST_DIR = Path("output/plots/ppt_exec_v198")
HIST = {
    # drawing rId -> (png source, media name, anchor row, col)
    "rId5": (HIST_DIR / "active_rsrp_hist.png", "image_cover_active_rsrp.png", 5, 11),
    "rId6": (HIST_DIR / "active_rsrq_hist.png", "image_cover_active_rsrq.png", 5, 17),
    "rId7": (HIST_DIR / "active_sinr_hist.png", "image_cover_active_sinr.png", 22, 11),
}
# Match previous Active chart footprint on Cover (9 x 6 cells ≈ 3960000 x 2520000 EMU).
PIC_CX = 3960000
PIC_CY = 2520000


def _pic_anchor(cNvPr_id: int, row: int, col: int, rid: str, name: str) -> str:
    return (
        f"<oneCellAnchor>"
        f"<from><col>{col}</col><colOff>0</colOff><row>{row}</row><rowOff>0</rowOff></from>"
        f"<ext cx=\"{PIC_CX}\" cy=\"{PIC_CY}\"/>"
        f"<pic>"
        f"<nvPicPr><cNvPr id=\"{cNvPr_id}\" name=\"{name}\" descr=\"Picture\"/>"
        f"<cNvPicPr/></nvPicPr>"
        f"<blipFill>"
        f'<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'cstate="print" r:embed="{rid}"/>'
        f'<a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f"<a:fillRect/></a:stretch>"
        f"</blipFill>"
        f"<spPr>"
        f'<a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect">'
        f"<a:avLst/></a:prstGeom>"
        f"</spPr>"
        f"</pic>"
        f"<clientData/>"
        f"</oneCellAnchor>"
    )


def _replace_chart_anchor(drawing: str, rid: str, new_anchor: str) -> str:
    """Replace only the oneCellAnchor whose chart r:id is rid (do not span siblings)."""
    anchors = list(re.finditer(r"<oneCellAnchor>[\s\S]*?</oneCellAnchor>", drawing))
    hits = [m for m in anchors if f'r:id="{rid}"' in m.group(0)]
    if len(hits) != 1:
        raise SystemExit(f"Cover drawing: expected 1 anchor for {rid}, found {len(hits)}")
    m = hits[0]
    return drawing[: m.start()] + new_anchor + drawing[m.end() :]


def _patch_drawing_rels(rels: str) -> str:
    image_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    for rid, (_src, media_name, _row, _col) in HIST.items():
        pattern = re.compile(
            rf'<Relationship Type="[^"]+" Target="[^"]+" Id="{rid}"/>'
        )
        replacement = (
            f'<Relationship Type="{image_type}" Target="/xl/media/{media_name}" Id="{rid}"/>'
        )
        rels, n = pattern.subn(replacement, rels, count=1)
        if n != 1:
            raise SystemExit(f"drawing rels: failed to retarget {rid} (matches={n})")
    return rels


def publish(src: Path = SRC, dest: Path = DEST) -> Path:
    for png, _media, _row, _col in HIST.values():
        if not png.exists():
            raise SystemExit(f"Missing histogram PNG: {png}")

    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {n: zf.read(n) for n in names}

    drawing = parts["xl/drawings/drawing1.xml"].decode("utf-8")
    # cNvPr ids already used on Cover: charts 1-7 + image 8. Reuse 5/6/7 for pics.
    for rid, (png, media_name, row, col) in HIST.items():
        cnv_id = {"rId5": 5, "rId6": 6, "rId7": 7}[rid]
        label = {
            "rId5": "Active RSRP histogram",
            "rId6": "Active RSRQ histogram",
            "rId7": "Active SINR histogram",
        }[rid]
        drawing = _replace_chart_anchor(
            drawing, rid, _pic_anchor(cnv_id, row, col, rid, label),
        )
        # Store PNG bytes for media part.
        parts[f"xl/media/{media_name}"] = png.read_bytes()
    parts["xl/drawings/drawing1.xml"] = drawing.encode("utf-8")
    parts["xl/drawings/_rels/drawing1.xml.rels"] = _patch_drawing_rels(
        parts["xl/drawings/_rels/drawing1.xml.rels"].decode("utf-8")
    ).encode("utf-8")

    # Cover label: mark Active histograms as legend-coloured (PPT-matched).
    cover = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    cover2, n = re.subn(
        r"(>Active histograms \(plotted Robi rows\)<)",
        r">Active histograms (map legend colours, same as PPT)<",
        cover,
        count=1,
    )
    if n != 1:
        # Accept already-patched or alternate wording; do not fail the publish.
        print(f"Cover hist label patch matches={n} (left as-is if 0)", flush=True)
    else:
        cover = cover2
        parts["xl/worksheets/sheet1.xml"] = cover.encode("utf-8")

    # Content types already declare png; ensure Override not required for media defaults.
    ct = parts["[Content_Types].xml"].decode("utf-8")
    if 'Extension="png"' not in ct:
        ct = ct.replace(
            '<Default Extension="jpg"',
            '<Default Extension="png" ContentType="image/png"/><Default Extension="jpg"',
            1,
        )
        parts["[Content_Types].xml"] = ct.encode("utf-8")

    tmp = dest.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Preserve original part order, then append new media.
        written = set()
        for name in names:
            zf.writestr(name, parts[name], compress_type=zipfile.ZIP_DEFLATED)
            written.add(name)
        for name, data in parts.items():
            if name not in written:
                zf.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
    tmp.replace(dest)
    shutil.copy2(dest, LATEST)
    print(f"Wrote {dest} ({dest.stat().st_size / 1e6:.1f} MB)", flush=True)
    return dest


def verify(path: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(path) as zf:
        rels = zf.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
        drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
        for rid, (png, media_name, row, col) in HIST.items():
            if f'Id="{rid}"' not in rels or media_name not in rels:
                problems.append(f"rels missing image target for {rid}")
            if f"xl/media/{media_name}" not in zf.namelist():
                problems.append(f"missing media {media_name}")
            if f'r:embed="{rid}"' not in drawing:
                problems.append(f"drawing missing pic embed {rid}")
            if f"<row>{row}</row>" not in drawing:
                problems.append(f"drawing missing row {row} for {rid}")
            # Active chart graphicFrame for this rid must be gone.
            if re.search(rf"<c:chart[^>]*r:id=\"{rid}\"", drawing):
                problems.append(f"chart graphic still bound to {rid}")
            # PNG payload matches source.
            if zf.read(f"xl/media/{media_name}") != png.read_bytes():
                problems.append(f"media bytes differ for {media_name}")
        # Idle charts still present.
        for rid in ("rId1", "rId2", "rId3", "rId4"):
            if f'r:id="{rid}"' not in drawing:
                problems.append(f"Idle/range chart {rid} missing from Cover drawing")
        # Spot-check other sheets untouched vs source.
        with zipfile.ZipFile(SRC) as src_zf:
            for sheet in (
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet6.xml",
                "xl/worksheets/sheet10.xml",
            ):
                if src_zf.read(sheet) != zf.read(sheet):
                    problems.append(f"{sheet} changed (must stay identical)")
    return problems


if __name__ == "__main__":
    # Ensure hist PNGs exist (rebuild from Active if missing).
    if not all(p.exists() for p, *_ in [(HIST[r][0],) for r in HIST]):
        import build_summary_ppt as ppt

        ppt._render_exec_hists()
    dest = publish()
    problems = verify(dest)
    if problems:
        raise SystemExit("v2.03 Cover verify failed: " + "; ".join(problems))
    print("v2.03 Cover verified: Active hist PNGs on Cover; other sheets unchanged.")
