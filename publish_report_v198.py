#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v1.98 from a fresh Active Session-ID rebuild (v1.97 layout) (Excel-openable).

Restructure:
1. One Bad Spot sheet — delete Bad Spot RSRP; keep Bad Spot RSRP_Active
2. Active Mode Bad Spot section first, Scanner Bad Spot section appended under it
3. Idle RSRP maps stay under Active for comparison; Idle solutions only in Scanner
4. Coloured Parameter change table + Solutions column; multi-circle for discrete pockets
5. Cover: Active first then Scanner, executive summary only, Active band mix + RSRP ranges
6. KPI sheets untouched. Web untouched.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

SRC = Path("output/Bogura_DriveTest_Report_v1.98.xlsx")
DEST = Path("output/Bogura_DriveTest_Report_v1.98.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")

_SCANNER_STRIDE = 40


def _sync(path: Path, dest: Path) -> Path:
    import generate_excel_report as ger

    path = Path(path)
    if path.resolve() != dest.resolve():
        shutil.copy2(path, dest)
    if dest.resolve() != ger.REPORT_XLSX.resolve():
        shutil.copy2(dest, ger.REPORT_XLSX)
    if dest.resolve() != ger.VERSIONED_XLSX.resolve():
        shutil.copy2(dest, ger.VERSIONED_XLSX)
    return dest


def _shift_row_attrs(xml: str, delta: int, min_row: int = 1) -> str:
    """Add delta to every row number reference at or above min_row."""

    def row_tag(match: re.Match) -> str:
        number = int(match.group(1))
        if number < min_row:
            return match.group(0)
        return f'<row r="{number + delta}"'

    def cell_ref(match: re.Match) -> str:
        col, number = match.group(1), int(match.group(2))
        if number < min_row:
            return match.group(0)
        return f'<c r="{col}{number + delta}"'

    def merge_ref(match: re.Match) -> str:
        a_col, a_row, b_col, b_row = match.group(1), int(match.group(2)), match.group(3), int(match.group(4))
        if a_row < min_row and b_row < min_row:
            return match.group(0)
        if a_row >= min_row:
            a_row += delta
        if b_row >= min_row:
            b_row += delta
        return f'<mergeCell ref="{a_col}{a_row}:{b_col}{b_row}"'

    xml = re.sub(r'<row r="(\d+)"', row_tag, xml)
    xml = re.sub(r'<c r="([A-Z]+)(\d+)"', cell_ref, xml)
    xml = re.sub(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', merge_ref, xml)
    return xml


def _append_scanner_section(parts: dict[str, bytes]) -> None:
    """Copy Idle Bad Spot maps/advice into Active sheet as Scanner Bad Spot section."""
    import generate_excel_report as ger

    active_path = ger._sheet_part_path(parts, "Bad Spot RSRP_Active")
    idle_path = ger._sheet_part_path(parts, "Bad Spot RSRP")
    active_xml = parts[active_path].decode("utf-8")
    idle_xml = parts[idle_path].decode("utf-8")

    active_headers = ger._title_rows(active_xml)
    expected_active = ger.LOGIC_LOCK["counts"]["Bad Spot RSRP_Active"][0]
    if len(active_headers) != expected_active:
        raise SystemExit(f"Active titles {len(active_headers)}, expected {expected_active}")
    idle_headers = ger._title_rows(idle_xml)
    if len(idle_headers) != 26:
        raise SystemExit(f"Idle titles {len(idle_headers)}, expected 26")

    active_rows = [int(n) for n in re.findall(r'<row r="(\d+)"', active_xml)]
    start_row = max(active_rows) + 3

    from openpyxl.utils import get_column_letter

    def banner(row: int, text: str) -> str:
        cells = [ger._inline_cell(f"A{row}", 5, text)]
        for col in range(2, 12):
            cells.append(f'<c r="{get_column_letter(col)}{row}" s="5" t="n" />')
        return f'<row r="{row}" ht="24" customHeight="1">{"".join(cells)}</row>'

    def insert_row_sorted(xml: str, row_xml: str, row_number: int) -> str:
        match = re.search(r"(<sheetData>)(.*?)(</sheetData>)", xml, flags=re.S)
        if match is None:
            raise SystemExit("Active sheetData missing")
        body = match.group(2)
        rows = re.findall(r"<row\b.*?</row>", body, flags=re.S)
        by_row: dict[int, str] = {}
        for block in rows:
            number = int(re.search(r'\br="(\d+)"', block).group(1))
            by_row[number] = block
        by_row[row_number] = row_xml
        new_body = "".join(by_row[n] for n in sorted(by_row))
        return xml[: match.start(2)] + new_body + xml[match.end(2) :]

    # Insert Active Mode section title above first Active bad spot if missing.
    if "Active Mode Bad Spot" not in active_xml:
        first = active_headers[1]
        banner_row = first - 2
        if banner_row < 2:
            banner_row = 2
        active_xml = insert_row_sorted(active_xml, banner(banner_row, "Active Mode Bad Spot"), banner_row)
        active_xml = re.sub(
            r'(<mergeCells count=")(\d+)(">)',
            lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
            active_xml,
            count=1,
        )
        active_xml = active_xml.replace(
            "</mergeCells>",
            f'<mergeCell ref="A{banner_row}:K{banner_row}"/></mergeCells>',
            1,
        )

    scanner_banner_row = start_row
    scanner_first = start_row + 2

    # Take Idle content from the first bad-spot title row through the end.
    idle_first = idle_headers[1]
    idle_sheet_data = re.search(r"<sheetData>(.*?)</sheetData>", idle_xml, flags=re.S)
    if idle_sheet_data is None:
        raise SystemExit("Idle sheetData missing")
    idle_row_blocks = re.findall(r"<row\b.*?</row>", idle_sheet_data.group(1), flags=re.S)
    kept = []
    for block in idle_row_blocks:
        number = int(re.search(r'\br="(\d+)"', block).group(1))
        if number < idle_first:
            continue
        kept.append(block)
    idle_slice = "".join(kept)
    delta = scanner_first - idle_first
    idle_slice = _shift_row_attrs(idle_slice, delta, min_row=idle_first)
    # Rename titles and circle labels so Active advice regex stays unique.
    idle_slice = re.sub(
        r"(>)Bad Spot (\d+)(<)",
        r"\1Scanner Bad Spot \2\3",
        idle_slice,
    )
    idle_slice = re.sub(
        r"(>)Circle (\d+)(<)",
        r"\1Circle S\2\3",
        idle_slice,
    )

    # Merges that belong to the idle bad-spot band.
    idle_merges = re.findall(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*/>', idle_xml)
    new_merges = []
    for a_col, a_row, b_col, b_row in idle_merges:
        a_n, b_n = int(a_row), int(b_row)
        if b_n < idle_first:
            continue
        new_merges.append(f'<mergeCell ref="{a_col}{a_n + delta}:{b_col}{b_n + delta}"/>')

    banner_xml = banner(scanner_banner_row, "Scanner Bad Spot")
    note_row = scanner_banner_row + 1
    note_xml = (
        f'<row r="{note_row}" ht="36" customHeight="1">'
        f'{ger._inline_cell(f"A{note_row}", 37, "Scanner (Idle) bad spots. Solutions in this section only. "
          "Idle RSRP comparison maps under Active Mode spots are for coverage difference — no Scanner solutions there.")}'
        f"</row>"
    )

    # Merge scanner rows into sheetData in sorted order.
    match = re.search(r"(<sheetData>)(.*?)(</sheetData>)", active_xml, flags=re.S)
    if match is None:
        raise SystemExit("Active sheetData missing before scanner append")
    by_row: dict[int, str] = {}
    for block in re.findall(r"<row\b.*?</row>", match.group(2), flags=re.S):
        number = int(re.search(r'\br="(\d+)"', block).group(1))
        by_row[number] = block
    by_row[scanner_banner_row] = banner_xml
    by_row[note_row] = note_xml
    for block in re.findall(r"<row\b.*?</row>", idle_slice, flags=re.S):
        number = int(re.search(r'\br="(\d+)"', block).group(1))
        by_row[number] = block
    new_body = "".join(by_row[n] for n in sorted(by_row))
    active_xml = active_xml[: match.start(2)] + new_body + active_xml[match.end(2) :]

    extra_merges = (
        [f'<mergeCell ref="A{scanner_banner_row}:K{scanner_banner_row}"/>',
         f'<mergeCell ref="A{note_row}:K{note_row}"/>']
        + new_merges
    )
    active_xml = re.sub(
        r'(<mergeCells count=")(\d+)(">)',
        lambda m: m.group(1) + str(int(m.group(2)) + len(extra_merges)) + m.group(3),
        active_xml,
        count=1,
    )
    active_xml = active_xml.replace(
        "</mergeCells>",
        "".join(extra_merges) + "</mergeCells>",
        1,
    )
    dim_last = max(by_row)
    active_xml = re.sub(
        r'<dimension ref="A1:[A-Z]+\d+"',
        f'<dimension ref="A1:T{dim_last}"',
        active_xml,
        count=1,
    )
    parts[active_path] = active_xml.encode("utf-8")

    # Merge Idle drawing anchors into Active drawing, shifted by delta.
    drawings = ger._sheet_drawings(parts)
    active_drawing = drawings["Bad Spot RSRP_Active"]
    idle_drawing = drawings["Bad Spot RSRP"]
    active_d = parts[active_drawing].decode("utf-8")
    idle_d = parts[idle_drawing].decode("utf-8")
    active_rels_name = active_drawing.replace("drawings/", "drawings/_rels/") + ".rels"
    idle_rels_name = idle_drawing.replace("drawings/", "drawings/_rels/") + ".rels"
    active_rels = parts[active_rels_name].decode("utf-8")
    idle_rels = parts[idle_rels_name].decode("utf-8")

    idle_rid_media = {}
    for rel in re.findall(r"<Relationship\b[^>]*/>", idle_rels):
        rid = re.search(r'Id="([^"]+)"', rel).group(1)
        target = re.search(r'Target="([^"]+)"', rel).group(1)
        media = target.lstrip("/") if target.startswith("/") else "xl/media/" + target.rsplit("/", 1)[-1]
        idle_rid_media[rid] = media

    next_rid = max((int(i) for i in re.findall(r'Id="rId(\d+)"', active_rels)), default=0) + 1
    next_shape = max((int(i) for i in re.findall(r'<cNvPr id="(\d+)"', active_d)), default=0) + 1
    media_numbers = [int(i) for i in re.findall(r"xl/media/image(\d+)\.", "\n".join(parts))]
    next_media = max(media_numbers, default=0) + 1

    anchors = []
    for block in re.findall(r"<oneCellAnchor>.*?</oneCellAnchor>", idle_d, flags=re.S):
        row = int(re.search(r"<row>(\d+)</row>", block).group(1))
        if row < idle_first - 5:
            # Keep overview image? Skip intro combined view above first spot.
            name = re.search(r'name="([^"]+)"', block)
            if name and name.group(1).startswith("Image") and row < idle_first - 1:
                continue
        new_row = row + delta
        block = re.sub(r"<row>\d+</row>", f"<row>{new_row}</row>", block, count=1)
        # Rename shapes.
        def rename(match: re.Match) -> str:
            name = match.group(1)
            if re.fullmatch(r"Circle \d+", name):
                name = name.replace("Circle ", "Circle S", 1)
            elif re.fullmatch(r"Image \d+", name):
                name = f"Scanner {name}"
            return f'name="{name}"'

        block = re.sub(r'name="([^"]+)"', rename, block, count=1)
        # New shape id
        block = re.sub(
            r'<cNvPr id="\d+"',
            f'<cNvPr id="{next_shape}"',
            block,
            count=1,
        )
        next_shape += 1
        # Copy media and remapp embed rid for pictures.
        embed = re.search(r'r:embed="(rId\d+)"', block)
        if embed:
            old_rid = embed.group(1)
            media = idle_rid_media[old_rid]
            blob = parts[media]
            # Duplicate media so dropping Idle sheet does not remove it.
            ext = Path(media).suffix
            new_media = f"xl/media/image{next_media}{ext}"
            next_media += 1
            parts[new_media] = blob
            new_rid = f"rId{next_rid}"
            next_rid += 1
            active_rels = active_rels.replace(
                "</Relationships>",
                f'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="/{new_media}" Id="{new_rid}" /></Relationships>',
                1,
            )
            block = block.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"', 1)
        anchors.append(block)

    active_d = active_d.replace("</wsDr>", "".join(anchors) + "</wsDr>")
    parts[active_drawing] = active_d.encode("utf-8")
    parts[active_rels_name] = active_rels.encode("utf-8")
    print(
        f"  Scanner section appended at row {scanner_first} "
        f"({len(idle_headers)} spots, {len(anchors)} drawing anchors)",
        flush=True,
    )


def _swap_cover_active_first(parts: dict[str, bytes], active_band: list[tuple[str, int, float]], active_rsrp: list[tuple[str, int, float]]) -> None:
    """Put Active Mode block first on Cover, then Scanner; fill Active band/RSRP tables."""
    import generate_excel_report as ger

    cover_path = ger._sheet_part_path(parts, "Cover")
    xml = parts[cover_path].decode("utf-8")
    drawing_name = ger._sheet_drawings(parts)["Cover"]
    drawing = parts[drawing_name].decode("utf-8")

    # Swap chart anchors: Idle charts at rows 5/22 ↔ Active at 45/57.
    def swap_from_row(text: str, a: int, b: int) -> str:
        text = text.replace(f"<row>{a}</row>", "<row>TMPROW</row>")
        text = text.replace(f"<row>{b}</row>", f"<row>{a}</row>")
        text = text.replace("<row>TMPROW</row>", f"<row>{b}</row>")
        return text

    drawing = swap_from_row(drawing, 5, 45)
    drawing = swap_from_row(drawing, 22, 57)
    parts[drawing_name] = drawing.encode("utf-8")

    # Top block was Idle; after chart swap Active charts sit on top — retitle.
    xml = xml.replace(">IDLE Mode<", ">__TOP_ACTIVE_MODE__<", 1)
    xml = xml.replace(">Active Mode<", ">Scanner (Idle Mode)<", 1)
    xml = xml.replace(">__TOP_ACTIVE_MODE__<", ">Active Mode<", 1)
    xml = xml.replace(">IDLE (scanner)<", ">__TOP_ACTIVE_UE__<", 1)
    # Lower Mode value that still says Active (UE) → Scanner.
    xml = xml.replace(">Active (UE)<", ">Scanner (Idle)<", 1)
    xml = xml.replace(">__TOP_ACTIVE_UE__<", ">Active (UE)<", 1)
    xml = xml.replace(
        ">Radio KPIs (all samples)<",
        ">Radio KPIs (Robi, plotted rows)<",
        1,
    )
    xml = xml.replace(
        ">RSRP, RSRQ and SINR plots (native Excel charts with legends)<",
        ">Active histograms (plotted Robi rows)<",
        1,
    )
    xml = xml.replace(">Band mix (DL EARFCN)<", ">Active band mix (DL EARFCN)<", 1)
    xml = xml.replace(">RSRP ranges (dBm)<", ">Active RSRP ranges (dBm)<", 1)

    xml = xml.replace(
        "Cover — Active Mode first, then Idle. The bad-spot summary lists Active, then Idle.",
        "Cover — Active Mode first, then Scanner (Idle). Executive summary for bad spots only; detail on Bad Spot RSRP_Active.",
    )
    xml = xml.replace(
        "Bad Spot RSRP / RSRQ / SINR are Idle. Bad Spot RSRP_Active / RSRQ_Active / SINR_Active apply the same poor-run rules to the Active Robi track. Each bad spot is its own large map, in order. The combined view has no circles.",
        "Bad Spot RSRP_Active holds Active Mode Bad Spots first, then Scanner Bad Spots. Idle comparison maps sit under Active spots (no Scanner solutions there). RSRQ / SINR / RSRP vs KPIs sheets are unchanged.",
    )

    # Rewrite Active band mix + RSRP range tables in the top band-mix area (rows 27+).
    def write_table_cell(xml: str, col: str, row: int, text: str, style: int = 20) -> str:
        cell = ger._inline_cell(f"{col}{row}", style, text)
        pattern = rf'<c r="{col}{row}"[^>]*(?:/>|>.*?</c>)'
        if re.search(pattern, xml, flags=re.S):
            return re.sub(pattern, cell, xml, count=1, flags=re.S)
        # Insert into the row if the row exists.
        match = re.search(rf'(<row r="{row}"[^>]*>)(.*?)(</row>)', xml, flags=re.S)
        if match is None:
            return xml
        return xml[: match.start()] + match.group(1) + match.group(2) + cell + match.group(3) + xml[match.end() :]

    xml = write_table_cell(xml, "A", 25, "Active band mix (DL EARFCN)", 5)
    xml = write_table_cell(xml, "E", 25, "Active RSRP ranges (dBm)", 5)
    for index, (label, count, share) in enumerate(active_band):
        row = 27 + index
        xml = write_table_cell(xml, "A", row, label, 20)
        xml = write_table_cell(xml, "B", row, f"{count:,}", 20)
        xml = write_table_cell(xml, "C", row, f"{share:.1%}", 20)
    for index, (label, count, share) in enumerate(active_rsrp):
        row = 27 + index
        xml = write_table_cell(xml, "E", row, label, 20)
        xml = write_table_cell(xml, "F", row, f"{count:,}", 20)
        xml = write_table_cell(xml, "G", row, f"{share:.1%}", 20)

    parts[cover_path] = xml.encode("utf-8")
    print("  Cover: Active first + Active band mix / RSRP ranges", flush=True)


def publish(src: Path = SRC, dest: Path = DEST) -> Path:
    import generate_excel_report as ger
    from merge_and_analyze import (
        best_carrier,
        best_server,
        earfcn_to_layer,
        idle_vs_active_note,
        load_active_merged,
        load_merged_csv,
        find_bad_spots,
        plot_idle_area_maps,
        remember_seen_sectors,
        select_active_report,
        PLOTS_DIR,
        LOGIC_LOCK,
    )
    import pandas as pd

    if not src.exists():
        raise SystemExit(f"Missing {src}")

    if Path(src).resolve() != Path(dest).resolve():
        shutil.copy2(src, dest)
    _sync(dest, dest)

    # 1) Redraw serial maps (multi-circle) into both sheets while Idle still exists.
    path = ger.refresh_bad_spot_pies(src=dest)
    print(f"serial maps refreshed -> {path}", flush=True)
    _sync(path, dest)

    # 2) Active advice with coloured Parameter change + Solutions column.
    advice_path = ger.write_bad_spot_advice(
        src=dest, sheets=("Bad Spot RSRP_Active",),
    )
    print(f"advice text refreshed -> {advice_path}", flush=True)
    _sync(advice_path, dest)

    # 3) Idle-area comparison maps under Active (no Scanner solutions here).
    idle = load_merged_csv()
    active = select_active_report(load_active_merged())
    remember_seen_sectors("Idle", idle)
    remember_seen_sectors("Active", active)
    idle_track = best_server(idle)
    active_track = best_carrier(active)
    spots = find_bad_spots(active_track, "RSRP").head(LOGIC_LOCK["counts"]["Bad Spot RSRP_Active"][0])
    idle_dir = PLOTS_DIR / "idle_area_v198"
    idle_dir.mkdir(parents=True, exist_ok=True)
    print(f"drawing {len(spots)} idle-area maps", flush=True)
    pngs = plot_idle_area_maps(idle_track, active_track, spots, idle_dir)
    notes = [idle_vs_active_note(active_track, idle_track, spots.iloc[i]) for i in range(len(spots))]
    ger.refresh_idle_area_maps(dest, dest, pngs, notes)
    _sync(dest, dest)

    # 4) Append Scanner section from Idle sheet, then drop Idle sheet.
    with zipfile.ZipFile(dest) as zf:
        names = zf.namelist()
        parts = {name: zf.read(name) for name in names}

    print("Appending Scanner Bad Spot section", flush=True)
    _append_scanner_section(parts)

    # Active band mix + RSRP ranges for Cover.
    layers = pd.to_numeric(active_track["DL EARFCN"], errors="coerce").map(earfcn_to_layer)
    layer_counts = layers.value_counts()
    active_band = [
        (str(name), int(n), float(n) / len(active_track))
        for name, n in layer_counts.items()
        if name
    ]
    from generate_excel_report import rsrp_range_counts

    rsrp_q = rsrp_range_counts(active_track["RSRP"])
    active_rsrp = [
        (str(label), int(n), float(n) / len(active_track))
        for label, n in rsrp_q.items()
    ]
    _swap_cover_active_first(parts, active_band, active_rsrp)

    print("Dropping Bad Spot RSRP sheet", flush=True)
    ger._drop_workbook_sheets(parts, ("Bad Spot RSRP",))

    # Content types / styles already updated by advice write; ensure coloured styles present.
    parts["xl/styles.xml"] = ger._add_advice_styles(parts["xl/styles.xml"].decode("utf-8")).encode("utf-8")

    tmp = dest.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            if name in parts:
                zf.writestr(name, parts[name])
        for name, body in parts.items():
            if name not in names:
                zf.writestr(name, body)
    tmp.replace(dest)
    _sync(dest, dest)
    print(f"Wrote {dest} ({dest.stat().st_size / 1e6:.1f} MB)", flush=True)
    return dest


if __name__ == "__main__":
    publish()
    shutil.copy2(DEST, LATEST)
    print(f"Copied to {LATEST}")
