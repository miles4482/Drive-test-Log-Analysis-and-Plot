#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v1.95 from v1.94 (Excel-openable).

Fixes Cover / Active RSRP stacked histograms where a 2 dBm bar straddled a
legend boundary and painted two colours in one column (e.g. -116 dBm showed
yellow + light-red together). Each dual-colour bin is collapsed into the
single legend series that owns the bin's left edge. Chart XML, drawings, and
all other sheets stay byte-identical to v1.94.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as XET

SRC = Path("output/Bogura_DriveTest_Report_v1.94.xlsx")
DEST = Path("output/Bogura_DriveTest_Report_v1.95.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")

# (cell_ref, new_value) — left-edge series keeps the sum; the other series → 0.
# Idle RSRP: Z bin, series AA..AE. Active RSRP: BH bin, series BI..BM.
CELL_PATCHES: dict[str, int] = {
    # Idle -116: yellow AC + pink AD → all pink (left edge -116 → -120 to -115)
    "AC14": 0,
    "AD14": 16180 + 15345,
    # Idle -96: dark AA + light AB → all light green (left edge -96 → -110 to -95)
    "AA24": 0,
    "AB24": 26809 + 26551,
    # Active -116
    "BK14": 0,
    "BL14": 3495 + 3408,
    # Active -96
    "BI24": 0,
    "BJ24": 8322 + 8787,
}


def patch_chart_data(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    for ref, value in CELL_PATCHES.items():
        pattern = rf'(<c r="{re.escape(ref)}" t="n"><v>)[^<]*(</v></c>)'
        new_text, n = re.subn(pattern, rf"\g<1>{value}\2", text, count=1)
        if n != 1:
            raise SystemExit(f"Failed to patch {ref} (matches={n})")
        text = new_text
    return text.encode("utf-8")


def publish(src: Path = SRC, dest: Path = DEST) -> Path:
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {n: zf.read(n) for n in names}

    sheet = "xl/worksheets/sheet14.xml"
    if sheet not in parts:
        raise SystemExit(f"Missing {sheet}")
    parts[sheet] = patch_chart_data(parts[sheet])

    tmp = dest.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name], compress_type=zipfile.ZIP_DEFLATED)
    tmp.replace(dest)
    print(f"Wrote {dest} ({dest.stat().st_size / 1e6:.1f} MB)", flush=True)
    return dest


def verify(path: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if name.endswith((".xml", ".rels")):
                try:
                    XET.fromstring(zf.read(name))
                except XET.ParseError as exc:
                    problems.append(f"XML parse {name}: {exc}")

        text = zf.read("xl/worksheets/sheet14.xml").decode("utf-8")
        expected = {
            "AC14": "0",
            "AD14": str(16180 + 15345),
            "AA24": "0",
            "AB24": str(26809 + 26551),
            "BK14": "0",
            "BL14": str(3495 + 3408),
            "BI24": "0",
            "BJ24": str(8322 + 8787),
        }
        for ref, value in expected.items():
            m = re.search(rf'<c r="{ref}" t="n"><v>([^<]*)</v></c>', text)
            if not m or m.group(1) != value:
                problems.append(f"{ref}: got {m.group(1) if m else None}, want {value}")

        # Charts / drawings must stay byte-identical to v1.94
        with zipfile.ZipFile(SRC) as src_zf:
            if set(zf.namelist()) != set(src_zf.namelist()):
                problems.append("part names differ from v1.94")
            for name in src_zf.namelist():
                if name.startswith("xl/charts/") or name.startswith("xl/drawings/"):
                    if zf.read(name) != src_zf.read(name):
                        problems.append(f"mutated: {name}")
                if name.startswith("xl/worksheets/") and name != "xl/worksheets/sheet14.xml":
                    if zf.read(name) != src_zf.read(name):
                        problems.append(f"sheet mutated: {name}")
    return problems


if __name__ == "__main__":
    if not SRC.exists():
        raise SystemExit(f"Missing {SRC}")
    publish()
    probs = verify(DEST)
    print("problems:", probs)
    if probs:
        raise SystemExit(1)
    shutil.copy2(DEST, LATEST)
    print(f"Copied to {LATEST}")
