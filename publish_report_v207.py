#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v2.07.

Updates from Mohammad's markup on v2.06:

1. Multi-circle bad-spot maps — when one oval cannot cover every poor sample
   and blank, draw two/three/four+ circles (no hard cap). Spot 9 (Poor run
   250 m + blank 2,769 m) is the reference case.
2. Drop the duplicate Parameter change table left under each Active bad spot
   (second copy sat at header+33, outside the old strip window).
3. Numbered Serving / Bad Spot / RCA / Suggestions blocks stay as the shared
   Excel+PPT list style; single-action suggestions are numbered too.
4. PPT paints every recorded circle for a map, not only the first.

Starts from openable v2.06, refreshes serial maps + advice, rebuilds PPT/Web.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

import generate_excel_report as ger
from generate_excel_report import REPORT_VERSION, VERSIONED_XLSX, REPORT_XLSX

SRC = Path("output/Bogura_DriveTest_Report_v2.06.xlsx")
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
            match = re.match(r'(<wsDr\b)([^>]*)>', text)
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
    if REPORT_VERSION != "2.07":
        raise SystemExit(f"Expected REPORT_VERSION 2.07, got {REPORT_VERSION}")
    if not SRC.exists():
        raise SystemExit(f"Missing {SRC}")
    shutil.copy2(SRC, DEST)
    shutil.copy2(DEST, LATEST)
    print(f"Base {SRC} -> {DEST}", flush=True)

    print("Refreshing Active serial bad-spot maps (multi-circle pockets)", flush=True)
    path = ger.refresh_bad_spot_pies(DEST, active_only=True)
    print(f"maps refreshed -> {path}", flush=True)

    print("Rewriting Active advice (drop duplicate Parameter change tables)", flush=True)
    advice = ger.write_bad_spot_advice(
        src=DEST, sheets=("Bad Spot RSRP_Active",),
    )
    print(f"advice -> {advice}", flush=True)

    with zipfile.ZipFile(DEST) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}
    n = _ensure_drawing_ns(parts)
    if n:
        tmp = DEST.with_suffix(".xlsx.tmp")
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name in names:
                zf.writestr(name, parts[name])
            for name, body in parts.items():
                if name not in names:
                    zf.writestr(name, body)
        tmp.replace(DEST)
        shutil.copy2(DEST, LATEST)
        print(f"drawing namespaces fixed on {n} part(s)", flush=True)

    if VERSIONED_XLSX.resolve() != DEST.resolve():
        shutil.copy2(DEST, VERSIONED_XLSX)
    shutil.copy2(DEST, LATEST)
    print(f"Wrote {DEST} ({DEST.stat().st_size / 1e6:.1f} MB)", flush=True)
    return DEST


if __name__ == "__main__":
    publish()
