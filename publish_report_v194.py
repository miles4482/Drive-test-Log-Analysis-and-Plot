#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v1.94 from v1.90 (Excel-openable Neon).

v1.91–v1.93 rewrote every worksheet with ElementTree and Excel rejected the
package as corrupt. This publisher keeps the v1.90 worksheet/drawing/chart XML
bytes intact (the last known Excel-openable structure) and only:

  1. Recolours styles.xml to the Neon navy / slate / teal palette
  2. Strips clutter instruction phrases from sheet text
  3. Softens banner titles where a safe string replace applies

No ElementTree sheet rewrite. No merge rebuild. No chart XML edits.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

SRC = Path("output/Bogura_DriveTest_Report_v1.90.xlsx")
DEST = Path("output/Bogura_DriveTest_Report_v1.94.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")

# Neon palette swaps against colours present in v1.90 styles.xml
STYLE_COLOR_MAP = {
    "001B4F72": "000F2942",  # banner navy -> Neon navy
    "00F4D03F": "000D9488",  # harsh yellow -> teal
    "00D5D8DC": "00E2E8F0",  # grey bars -> slate
    "00F4F6F7": "00F8FAFC",  # paper -> cooler paper
}

# Exact phrase removals / cleanups inside <t>…</t> (worksheet text only)
CLUTTER_REPLACEMENTS = [
    ("How to read this workbook", ""),
    ("  |  Gridlines off", ""),
    ("  | Gridlines off", ""),
    ("Gridlines off", ""),
    ("  |  Freeze panes off", ""),
    ("Freeze panes off", ""),
    ("No numbering", ""),
    ("Combined view has no circles", ""),
    ("no circles", ""),
]

# Banner title soft upgrades (safe literal replaces)
BANNER_REPLACEMENTS = [
    ("  RSRP Report", "RSRP Coverage Maps"),
    ("  RSRQ Report", "RSRQ Coverage Maps"),
    ("  SINR Report", "SINR Coverage Maps"),
]


def recolor_styles(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    for old, new in STYLE_COLOR_MAP.items():
        text = text.replace(f'rgb="{old}"', f'rgb="{new}"')
        text = text.replace(f"rgb='{old}'", f"rgb='{new}'")
    return text.encode("utf-8")


def scrub_sheet(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    for old, new in CLUTTER_REPLACEMENTS + BANNER_REPLACEMENTS:
        if old in text:
            text = text.replace(old, new)
    # Drop leftover double spaces inside simple <t> runs (cosmetic)
    text = re.sub(r"(<t[^>]*>)  +", r"\1", text)
    return text.encode("utf-8")


def publish(src: Path = SRC, dest: Path = DEST) -> Path:
    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {n: zf.read(n) for n in names}

    parts["xl/styles.xml"] = recolor_styles(parts["xl/styles.xml"])

    for name in list(parts):
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            parts[name] = scrub_sheet(parts[name])

    tmp = dest.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Preserve original entry order from v1.90
        for name in names:
            zf.writestr(name, parts[name], compress_type=zipfile.ZIP_DEFLATED)
    tmp.replace(dest)
    print(f"Wrote {dest} ({dest.stat().st_size / 1e6:.1f} MB)", flush=True)
    return dest


def verify(path: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(path) as zf:
        # byte-identical worksheets to v1.90 except scrubbed text — structure check:
        # every XML part must parse
        from xml.etree import ElementTree as XET

        for name in zf.namelist():
            if name.endswith((".xml", ".rels")):
                try:
                    XET.fromstring(zf.read(name))
                except XET.ParseError as exc:
                    problems.append(f"XML parse {name}: {exc}")

        styles = zf.read("xl/styles.xml").decode("utf-8")
        if "F4D03F" in styles:
            problems.append("styles still contain #F4D03F")
        if "0F2942" not in styles and "000F2942" not in styles:
            problems.append("Neon navy not applied to styles")

        for name in zf.namelist():
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            if name.endswith("sheet14.xml"):
                continue
            text = zf.read(name).decode("utf-8", errors="ignore")
            for token in (
                "How to read this workbook",
                "Gridlines off",
                "Freeze panes off",
                "No numbering",
                "Combined view has no circles",
            ):
                if token in text:
                    problems.append(f"{name}: still has {token!r}")

        # drawings / ellipses unchanged from v1.90
        d8 = zf.read("xl/drawings/drawing8.xml").decode("utf-8")
        if d8.count('prst="ellipse"') != 99:
            problems.append(f"drawing8 ellipses {d8.count('prst=\"ellipse\"')}, expected 99")
        d5 = zf.read("xl/drawings/drawing5.xml").decode("utf-8")
        if d5.count('prst="ellipse"') != 26:
            problems.append(f"drawing5 ellipses {d5.count('prst=\"ellipse\"')}, expected 26")

        # structural identity: same part names as source
        with zipfile.ZipFile(SRC) as src_zf:
            if set(zf.namelist()) != set(src_zf.namelist()):
                problems.append("part names differ from v1.90")
            # charts must be byte-identical to v1.90 (no polish)
            for name in src_zf.namelist():
                if name.startswith("xl/charts/"):
                    if zf.read(name) != src_zf.read(name):
                        problems.append(f"chart mutated: {name}")
                if name.startswith("xl/drawings/"):
                    if zf.read(name) != src_zf.read(name):
                        problems.append(f"drawing mutated: {name}")
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
