#!/usr/bin/env python3
"""Build the interactive RSRP / RSRQ coverage page.

The page follows the RSRP and RSRQ sheets: Idle and Active, all-band and
L900 / L1800 / L2100 / L2600, on the same 50 m bins. Each plotted spot keeps
the Cell ID and the Cell Name from System DUMP.xlsb.

Idle Cell ID is the Sector Id. Cell Name is the dump EUtranCellId whose
eNBId and cellId match the idle eNB-Id and Sector Id.

Active Cell ID is Cell Id mod 256, and eNBId is Cell Id divided by 256
rounded down. Cell Name is the dump EUtranCellId for that pair.
"""
from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd

from merge_and_analyze import (
    ACTIVE_MERGED,
    ACTIVE_OPERATOR,
    MAP_BIN_M,
    MERGED_CSV,
    REPORT_LAYERS,
    SITE_DB_PATH,
    _bin_steps,
    best_server,
    earfcn_to_layer,
    filter_by_layer,
    load_active_merged,
    load_merged_csv,
    strongest_fix,
    strongest_location,
)

ROOT = Path(__file__).resolve().parent
WEB_DIR = ROOT / "web"
DATA_JS = WEB_DIR / "coverage_data.js"


def load_cell_directory(path: Path = SITE_DB_PATH) -> tuple[pd.DataFrame, pd.DataFrame]:
    """EUtranCellId names keyed by eNB, cell, and EARFCN."""
    from pyxlsb import open_workbook

    rows: list[list] = []
    cols: list[str] = []
    with open_workbook(str(path)) as wb:
        sheet_name = "4G" if "4G" in wb.sheets else wb.sheets[0]
        sheet = wb.get_sheet(sheet_name)
        for i, row in enumerate(sheet.rows()):
            vals = [c.v for c in row]
            if i == 0:
                cols = [str(v).strip() if v is not None else "" for v in vals]
                continue
            rows.append(vals)
    dump = pd.DataFrame(rows, columns=cols)
    missing = [c for c in ("eNBId", "cellId", "EARFCN", "EUtranCellId") if c not in dump.columns]
    if missing:
        raise SystemExit(f"System dump is missing columns: {missing}")
    dump["eNBId"] = pd.to_numeric(dump["eNBId"], errors="coerce")
    dump["cellId"] = pd.to_numeric(dump["cellId"], errors="coerce")
    dump["EARFCN"] = pd.to_numeric(dump["EARFCN"], errors="coerce")
    dump["EUtranCellId"] = dump["EUtranCellId"].astype("string").str.strip()
    dump = dump.dropna(subset=["eNBId", "cellId", "EUtranCellId"])
    dump = dump[dump["EUtranCellId"].ne("") & dump["EUtranCellId"].ne("<NA>")]
    dump["eNBId"] = dump["eNBId"].astype(np.int64)
    dump["cellId"] = dump["cellId"].astype(np.int64)
    exact = (
        dump.dropna(subset=["EARFCN"])
        .assign(EARFCN=lambda frame: frame["EARFCN"].round().astype(np.int64))
        .drop_duplicates(["eNBId", "cellId", "EARFCN"], keep="first")
        [["eNBId", "cellId", "EARFCN", "EUtranCellId"]]
    )
    pair = (
        dump.groupby(["eNBId", "cellId"], sort=False)["EUtranCellId"]
        .agg(lambda series: series.iloc[0] if series.nunique() == 1 else "")
        .reset_index()
        .rename(columns={"EUtranCellId": "pair_name"})
    )
    pair = pair[pair["pair_name"].ne("")]
    return exact, pair


def _as_int(series: pd.Series) -> pd.Series:
    num = pd.to_numeric(series, errors="coerce")
    return num.round().astype("Int64")


def attach_identity(df: pd.DataFrame, mode: str, exact: pd.DataFrame, pair: pd.DataFrame) -> pd.DataFrame:
    """Add the Cell ID and Cell Name used by the click popup."""
    work = df.copy()
    if mode == "idle":
        enb = _as_int(work["eNB-Id"]) if "eNB-Id" in work.columns else pd.Series(pd.NA, index=work.index, dtype="Int64")
        cell = _as_int(work["Sector Id"]) if "Sector Id" in work.columns else pd.Series(pd.NA, index=work.index, dtype="Int64")
        full = pd.to_numeric(work["Cell Id"], errors="coerce")
        missing = cell.isna() & full.notna()
        if missing.any():
            enb = enb.mask(missing, (full // 256).round().astype("Int64"))
            cell = cell.mask(missing, (full % 256).round().astype("Int64"))
    else:
        full = pd.to_numeric(work["Cell Id"], errors="coerce")
        enb = (full // 256).round().astype("Int64")
        cell = (full % 256).round().astype("Int64")
    work["cell_id"] = cell
    work["_enb"] = enb
    earfcn = pd.to_numeric(work.get("DL EARFCN"), errors="coerce").round().astype("Int64")
    work["_earfcn"] = earfcn
    keyed = work.merge(exact, left_on=["_enb", "cell_id", "_earfcn"], right_on=["eNBId", "cellId", "EARFCN"], how="left")
    keyed = keyed.merge(pair, left_on=["_enb", "cell_id"], right_on=["eNBId", "cellId"], how="left", suffixes=("", "_pair"))
    if len(keyed) != len(work):
        raise RuntimeError(f"Cell name lookup changed the row count ({len(work)} -> {len(keyed)})")
    name = keyed["EUtranCellId"].astype("string")
    name = name.mask(name.isna() | name.eq("") | name.eq("<NA>"), keyed["pair_name"].astype("string"))
    cleaned = name.fillna("").astype(str)
    cleaned = cleaned.where(~cleaned.isin(["<NA>", "nan", "None"]), "")
    work["cell_name"] = cleaned.to_numpy()
    return work


def bin_track(df: pd.DataFrame) -> pd.DataFrame:
    """One spot per 50 m bin: median colour, cell from the sample nearest that median."""
    work = df.dropna(subset=["Longitude", "Latitude"]).copy().reset_index(drop=True)
    work["RSRP"] = pd.to_numeric(work["RSRP"], errors="coerce")
    work["RSRQ"] = pd.to_numeric(work["RSRQ"], errors="coerce")
    work = work[work["RSRP"].notna() | work["RSRQ"].notna()]
    if work.empty:
        return work
    lat = work["Latitude"]
    lon = work["Longitude"]
    dlat, dlon = _bin_steps(lat, MAP_BIN_M)
    work["_bin"] = list(zip(np.floor(lat / dlat).astype(np.int64), np.floor(lon / dlon).astype(np.int64)))
    med = work.groupby("_bin", sort=False)[["RSRP", "RSRQ"]].median()
    work["_med"] = work["_bin"].map(med["RSRP"])
    work["_dist"] = (work["RSRP"] - work["_med"]).abs()
    work["_dist"] = work["_dist"].fillna(1e9)
    rep_idx = work.groupby("_bin", sort=False)["_dist"].idxmin()
    rep = work.loc[rep_idx, ["_bin", "cell_id", "cell_name"]].set_index("_bin")
    pos = work.groupby("_bin", sort=False)[["Latitude", "Longitude"]].mean()
    out = pos.join(med).join(rep)
    out = out.reset_index(drop=True)
    return out


def _records(frame: pd.DataFrame, names: dict[str, int]) -> dict:
    if frame.empty:
        return {"lon": [], "lat": [], "rsrp": [], "rsrq": [], "cell": [], "name": []}
    cell = pd.to_numeric(frame["cell_id"], errors="coerce")
    name_ids = []
    for value in frame["cell_name"].astype(str):
        if value in ("", "nan", "<NA>"):
            name_ids.append(0)
            continue
        if value not in names:
            names[value] = len(names)
        name_ids.append(names[value])
    def series_or_null(col: str, digits: int) -> list:
        num = pd.to_numeric(frame[col], errors="coerce").round(digits)
        return [None if pd.isna(v) else float(v) for v in num]

    cells = [None if pd.isna(v) else int(v) for v in cell]
    return {
        "lon": series_or_null("Longitude", 5),
        "lat": series_or_null("Latitude", 5),
        "rsrp": series_or_null("RSRP", 1),
        "rsrq": series_or_null("RSRQ", 1),
        "cell": cells,
        "name": name_ids,
    }


def build() -> None:
    if not MERGED_CSV.exists() or not ACTIVE_MERGED.exists():
        raise SystemExit("Merged drive logs are missing.")
    if not SITE_DB_PATH.exists():
        raise SystemExit(f"Missing {SITE_DB_PATH.name}")
    print("Loading system dump", flush=True)
    exact, pair = load_cell_directory()
    print(f"  names by EARFCN {len(exact):,}  unique pairs {len(pair):,}", flush=True)
    print("Loading drive logs", flush=True)
    idle = load_merged_csv()
    active = load_active_merged()
    active = active.loc[active["Operator"].astype("string").str.strip().eq(ACTIVE_OPERATOR)].copy()
    idle_base = strongest_fix(idle)
    idle_all = best_server(idle_base)
    active_all = strongest_location(active)
    print(f"Idle best server {len(idle_all):,}  Active plotted {len(active_all):,}", flush=True)

    tracks: dict[str, pd.DataFrame] = {
        "idle|all": idle_all,
        "active|all": active_all,
    }
    for layer in REPORT_LAYERS:
        tracks[f"idle|{layer}"] = best_server(filter_by_layer(idle_base, layer))
        tracks[f"active|{layer}"] = filter_by_layer(active_all, layer)

    names: dict[str, int] = {"": 0}
    sets = {}
    for key, frame in tracks.items():
        mode = "idle" if key.startswith("idle") else "active"
        print(f"Binning {key} ({len(frame):,} rows)", flush=True)
        identified = attach_identity(frame, mode, exact, pair)
        binned = bin_track(identified)
        named = int((binned["cell_name"].astype(str).str.len() > 0).sum()) if not binned.empty else 0
        print(f"  {len(binned):,} spots, {named:,} with a Cell Name", flush=True)
        sets[key] = _records(binned, names)

    # names dict was filled in first-seen order starting at 1; index 0 stays blank.
    ordered = [""] * len(names)
    for value, index in names.items():
        ordered[index] = value
    payload = {"names": ordered, "sets": sets}
    WEB_DIR.mkdir(parents=True, exist_ok=True)
    text = "const COVERAGE = " + json.dumps(payload, separators=(",", ":")) + ";\n"
    DATA_JS.write_text(text)
    print(f"Wrote {DATA_JS} ({DATA_JS.stat().st_size / 1e6:.1f} MB)", flush=True)
    add_site_pies()
    package_web()


def add_site_pies(radius_km: float = 3.0) -> int:
    """Sites along the drive, one centre and every sector arm, for the web pies."""
    from merge_and_analyze import _sites_near_drive, load_site_sectors

    text = DATA_JS.read_text().strip()
    if not text.startswith("const COVERAGE = "):
        raise SystemExit(f"{DATA_JS.name} is not a coverage page")
    payload = json.loads(text[len("const COVERAGE = "):-1])
    lat: list[float] = []
    lon: list[float] = []
    for key in ("idle|all", "active|all"):
        pack = payload["sets"][key]
        lat.extend(pack["lat"])
        lon.extend(pack["lon"])
    drive = pd.DataFrame({"Latitude": lat, "Longitude": lon})
    near = _sites_near_drive(load_site_sectors(), drive, radius_km=radius_km)
    sites: list[list] = []
    for _site, part in near.groupby("SiteName", sort=False):
        part = part.drop_duplicates(subset=["Sector"]).dropna(subset=["azimuth"])
        if part.empty:
            continue
        arms = []
        for _, row in part.iterrows():
            name = str(row["Sector"]).strip()
            if not name or name.lower() == "nan":
                continue
            arms.append([name, round(float(row["azimuth"]) % 360.0, 1)])
        if not arms:
            continue
        sites.append([
            round(float(part["lat"].median()), 5),
            round(float(part["lon"].median()), 5),
            arms,
        ])
    payload["sites"] = sites
    DATA_JS.write_text("const COVERAGE = " + json.dumps(payload, separators=(",", ":")) + ";\n")
    print(f"Site pies {len(sites):,} along the drive", flush=True)
    return len(sites)


def package_web() -> Path:
    """One download named with the same version as the Excel report."""
    from generate_excel_report import OUTPUT_DIR, REPORT_VERSION

    html_text = (WEB_DIR / "index.html").read_text()
    marker = f"v{REPORT_VERSION}"
    if marker not in html_text:
        raise SystemExit(f"web/index.html is not labelled {marker}, which is the Excel report version")
    if not DATA_JS.exists():
        raise SystemExit(f"Missing {DATA_JS.name}; build the coverage data first")
    page = html_text.replace(
        '<script src="coverage_data.js"></script>',
        "<script>\n" + DATA_JS.read_text() + "</script>",
    )
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    dest = OUTPUT_DIR / f"Bogura_DriveTest_Web_v{REPORT_VERSION}.zip"
    inner = f"Bogura_DriveTest_Web_v{REPORT_VERSION}.html"
    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(inner, page)
    print(f"Wrote {dest} ({dest.stat().st_size / 1e6:.1f} MB) containing {inner}", flush=True)
    return dest


if __name__ == "__main__":
    if "--package" in sys.argv:
        add_site_pies()
        package_web()
    else:
        build()
