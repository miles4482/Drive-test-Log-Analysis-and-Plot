#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v1.96 from v1.95 (Excel-openable).

Fixes from the annotated Active Bad Spot maps:

1. Idle-area maps draw the same yellow active serving lines as the active map
2. Uncovered red runs / blanks inside each circle get solutions (short-run + pocket)
3. Proposed azimuth turns are dotted; sector gaps stay at least 45°
4. Drawn circles include light-red RSRP and adjacent blanks

Maps and advice text are regenerated and grafted into the openable v1.95 package.
Charts outside Bad Spot RSRP / Bad Spot RSRP_Active stay byte-identical to v1.95.
"""
from __future__ import annotations

import shutil
from pathlib import Path

SRC = Path("output/Bogura_DriveTest_Report_v1.95.xlsx")
DEST = Path("output/Bogura_DriveTest_Report_v1.96.xlsx")
LATEST = Path("output/Bogura_DriveTest_Report.xlsx")


def publish(src: Path = SRC, dest: Path = DEST) -> Path:
    import generate_excel_report as ger
    from merge_and_analyze import (
        best_carrier,
        best_server,
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

    if not src.exists():
        raise SystemExit(f"Missing {src}")
    def _sync(path: Path) -> Path:
        path = Path(path)
        if path.resolve() != dest.resolve():
            shutil.copy2(path, dest)
        if dest.resolve() != ger.REPORT_XLSX.resolve():
            shutil.copy2(dest, ger.REPORT_XLSX)
        if dest.resolve() != ger.VERSIONED_XLSX.resolve():
            shutil.copy2(dest, ger.VERSIONED_XLSX)
        return dest

    # Start from the last openable Neon package.
    shutil.copy2(src, dest)
    _sync(dest)

    # Redraw serial bad-spot maps + movable circles into the versioned workbook.
    path = ger.refresh_bad_spot_pies(src=dest)
    print(f"serial maps refreshed -> {path}", flush=True)
    _sync(path)

    # Refresh Active advice only. Idle sheet row gaps are 38–40 (Neon package),
    # too tight for a fresh 41-row parameter block; Idle text stays from v1.95.
    advice_path = ger.write_bad_spot_advice(
        src=dest, sheets=("Bad Spot RSRP_Active",),
    )
    print(f"advice text refreshed -> {advice_path}", flush=True)
    _sync(advice_path)

    # Idle-area maps under each Active bad spot, with active+idle serving lines.
    idle = load_merged_csv()
    active = select_active_report(load_active_merged())
    remember_seen_sectors("Idle", idle)
    remember_seen_sectors("Active", active)
    idle_track = best_server(idle)
    active_track = best_carrier(active)
    spots = find_bad_spots(active_track, "RSRP").head(LOGIC_LOCK["counts"]["Bad Spot RSRP_Active"][0])
    idle_dir = PLOTS_DIR / "idle_area_v196"
    idle_dir.mkdir(parents=True, exist_ok=True)
    print(f"drawing {len(spots)} idle-area maps", flush=True)
    pngs = plot_idle_area_maps(idle_track, active_track, spots, idle_dir)
    notes = [idle_vs_active_note(active_track, idle_track, spots.iloc[i]) for i in range(len(spots))]
    ger.refresh_idle_area_maps(dest, dest, pngs, notes)
    _sync(dest)
    print(f"Wrote {dest} ({dest.stat().st_size / 1e6:.1f} MB)", flush=True)
    return dest


if __name__ == "__main__":
    publish()
    shutil.copy2(DEST, LATEST)
    print(f"Copied to {LATEST}")
