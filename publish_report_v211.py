#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v2.11.

Circle rules from the three marked Active maps, applied to all 97 spots:

1. A circle sits on poor RSRP (below -120) or on a blank with a weak end
   or beside that poor road. A gap in good RSRP is not a bad spot.
2. Holes on one straight road share one circle. A pocket on a different
   road keeps its own circle, so empty land between them is not circled.
3. Every poor sample and every counted blank sits inside a circle.

Starts from openable v2.10, refreshes Active serial maps + advice.
Cover is put back from that base so merges stay valid.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import generate_excel_report as ger
from generate_excel_report import REPORT_VERSION, VERSIONED_XLSX

SRC = Path("output/Bogura_DriveTest_Report_v2.10.xlsx")
DEST = Path(f"output/Bogura_DriveTest_Report_v{REPORT_VERSION}.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")

DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def _ensure_drawing_ns(parts: dict[str, bytes]) -> int:
    """Bind a: and r: on every drawing root so circle shapes stay well-formed."""
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
        if f'xmlns:a="{DML}"' not in attrs and "xmlns:a=" not in attrs:
            extra += f' xmlns:a="{DML}"'
        if f'xmlns:r="{REL}"' not in attrs and "xmlns:r=" not in attrs:
            extra += f' xmlns:r="{REL}"'
        if f'xmlns:xdr="{XDR}"' not in attrs and "xmlns:xdr=" not in attrs and match.group(1).startswith("<xdr:"):
            extra += f' xmlns:xdr="{XDR}"'
        if not extra:
            continue
        text = text[: match.end(2)] + extra + text[match.end(2) :]
        parts[name] = text.encode("utf-8")
        fixed += 1
    return fixed


def publish() -> Path:
    if REPORT_VERSION != "2.11":
        raise SystemExit(f"Expected REPORT_VERSION 2.11, got {REPORT_VERSION}")
    if not SRC.exists():
        raise SystemExit(f"Missing {SRC}")
    shutil.copy2(SRC, DEST)
    shutil.copy2(DEST, LATEST)
    print(f"Base {SRC} -> {DEST}", flush=True)

    print("Refreshing Active serial maps", flush=True)
    path = ger.refresh_bad_spot_pies(DEST, active_only=True)
    print(f"maps refreshed -> {path}", flush=True)

    print("Rewriting Active advice so blank metres and solutions match", flush=True)
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
        for name, body in parts.items():
            if name not in names:
                zf.writestr(name, body)
    tmp.replace(DEST)
    shutil.copy2(DEST, LATEST)
    print(f"drawing namespaces fixed on {n} part(s); Cover restored from {SRC.name}", flush=True)

    if VERSIONED_XLSX.resolve() != DEST.resolve():
        shutil.copy2(DEST, VERSIONED_XLSX)
    shutil.copy2(DEST, LATEST)
    print(f"Wrote {DEST} ({DEST.stat().st_size / 1e6:.1f} MB)", flush=True)
    return DEST


if __name__ == "__main__":
    publish()
