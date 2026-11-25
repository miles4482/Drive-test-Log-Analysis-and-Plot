#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v2.14.

Circles on every Active bad spot:

1. Sit on the poor road or on a blank beside that road — not on empty land.
2. Do not overlap. Seeds are spaced ~1 km along the road; neighbours shrink
   so ovals never stack.
3. Cap at ~1 km across so a long blank does not paint fields.

Starts from openable v2.13, refreshes all Active serial maps + advice.
Cover is put back from that base so merges stay valid.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import generate_excel_report as ger
from generate_excel_report import REPORT_VERSION, VERSIONED_XLSX

SRC = Path("output/Bogura_DriveTest_Report_v2.13.xlsx")
DEST = Path(f"output/Bogura_DriveTest_Report_v{REPORT_VERSION}.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")

DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


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


def publish() -> Path:
    if REPORT_VERSION != "2.14":
        raise SystemExit(f"Expected REPORT_VERSION 2.14, got {REPORT_VERSION}")
    if not SRC.exists():
        raise SystemExit(f"Missing {SRC}")
    shutil.copy2(SRC, DEST)
    shutil.copy2(DEST, LATEST)
    print(f"Base {SRC} -> {DEST}", flush=True)

    print("Refreshing Active serial maps", flush=True)
    path = ger.refresh_bad_spot_pies(DEST, active_only=True)
    print(f"maps refreshed -> {path}", flush=True)

    print("Rewriting Active advice", flush=True)
    advice = ger.write_bad_spot_advice(
        src=DEST, sheets=("Bad Spot RSRP_Active",),
    )
    print(f"advice -> {advice}", flush=True)

    with zipfile.ZipFile(DEST) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    n = _ensure_drawing_ns(parts)
    with zipfile.ZipFile(SRC) as z_src:
        parts["xl/worksheets/sheet1.xml"] = z_src.read("xl/worksheets/sheet1.xml")
    tmp = DEST.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name])
    tmp.replace(DEST)
    moved = ger.spread_circle_anchors(DEST)
    print(f"drawing namespaces fixed on {n} part(s); spread {moved} circles; Cover restored", flush=True)

    if VERSIONED_XLSX.resolve() != DEST.resolve():
        shutil.copy2(DEST, VERSIONED_XLSX)
    shutil.copy2(DEST, LATEST)
    print(f"Wrote {DEST} ({DEST.stat().st_size / 1e6:.1f} MB)", flush=True)
    return DEST


if __name__ == "__main__":
    publish()
