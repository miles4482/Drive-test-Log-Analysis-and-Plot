#!/usr/bin/env python3
"""Extract Bogura idle scanner logs, merge CSVs, and plot KPIs."""

from __future__ import annotations

import math
import re
import subprocess
import sys
import urllib.request
from collections import deque
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import patheffects as pe
from matplotlib.collections import LineCollection
from matplotlib.lines import Line2D
from matplotlib.patches import Ellipse, PathPatch, Wedge, Circle
from matplotlib.path import Path as MPath
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
EXTRACT_DIR = ROOT / "extracted"
OUTPUT_DIR = ROOT / "output"
PLOTS_DIR = OUTPUT_DIR / "plots"

IDLE_ARCHIVE = ROOT / "Borga_Idle_V2.0.7z"
IDLE_DIR = EXTRACT_DIR / "Borga_Idle_V2.0"
MERGED_CSV = OUTPUT_DIR / "Bogura_merged.csv.gz"
SUMMARY_MD = OUTPUT_DIR / "analysis_summary.md"
# Cell-level system dump. One site pie is built from its Sector column, not from the
# old node-level Physical_Site_Database_V1.0.xlsb.
SITE_DB_PATH = ROOT / "System DUMP.xlsb"
# Active Mode UE log. The 11 Nov 2026 archive replaces Borga_Active_mode.csv.
# It holds one CSV per collection, the same shape as the idle export. Robi,
# Grameenphone and Banglalink are in the files; the report draws Robi only.
ACTIVE_ARCHIVE = ROOT / "Borga_Active_mode_11Nov26.7z"
ACTIVE_DIR = EXTRACT_DIR / "Borga_Active_mode_11Nov26"
ACTIVE_MERGED = OUTPUT_DIR / "Bogura_active_merged.csv.gz"
# The CSV stores a clock ("8:40:30 AM") and no calendar date. Pin that clock
# so a later rebuild does not move the timestamp onto the day the script runs.
ACTIVE_CLOCK_ANCHOR = pd.Timestamp("1970-01-01")
# Idle maps are the Robi scanner. The active file also carries a Grameenphone UE on the
# same drive; only Robi is drawn so Active sits beside Idle for the same network.
ACTIVE_OPERATOR = "Robi"
# Primary serving cell. M-SCell1 and M-SCell2 are carrier-aggregation neighbours
# and are left out of the Active maps.
ACTIVE_CELL_TYPE = "M-PCell"

SCANNER_RENAME = {
    "NB RSRP": "RSRP",
    "NB RSRQ": "RSRQ",
    "NB RS SINR": "SINR",
}

COLUMNS = [
    "Time",
    "RSRP",
    "RSRQ",
    "SINR",
    "Longitude",
    "Latitude",
    "Cell Id",
    "DL EARFCN",
]

# LTE downlink EARFCN ranges (3GPP TS 36.101).
EARFCN_BANDS = {
    1: (0, 599, "B1 2100 MHz"),
    3: (1200, 1949, "B3 1800 MHz"),
    8: (3450, 3799, "B8 900 MHz"),
    40: (38650, 39649, "B40 2300 MHz"),
    41: (39650, 41589, "B41 2500 MHz"),
}

# The all-band best-server track is ~61k points; plot it whole so no poor sample
# is dropped from a map that bad spots are drawn on.
MAP_MAX_POINTS = 150000
# Drive-test bin: maps are coloured by the median KPI of each 50 m of road, and
# bad spots are judged on the same bins, so a circle can only sit on colour that
# the map actually shows.
MAP_BIN_M = 50.0

# One scanner sweep of all bands takes ~3 s; bound it in distance too so a fast
# stretch of road does not merge separate locations into one best-server sample.
BEST_SERVER_SWEEP_S = 4.0
BEST_SERVER_SWEEP_M = 40.0

# Operator layer names used on the RSRP / RSRQ / SINR sheets.
REPORT_LAYERS = ("L900", "L1800", "L2100", "L2600")
LAYER_EARFCN_RANGES = {
    "L900": ((3450, 3799),),
    "L1800": ((1200, 1949),),
    "L2100": ((0, 599),),
    "L2600": ((2750, 3449), (39650, 41589)),  # Band 7 FDD 2600 + Band 41 TDD 2500/2600
}

# Coverage-map legend (Excel Page 1 colour table). Best bin is listed first.
# Colours match the standard Office fills used in that table.
LEGEND_BLUE = "#0070C0"
LEGEND_DARK_GREEN = "#00B050"
LEGEND_LIGHT_GREEN = "#92D050"
LEGEND_YELLOW = "#FFFF00"
LEGEND_MAGENTA = "#FF00FF"
LEGEND_RED = "#FF0000"
LEGEND_CYAN = "#00B0F0"
LEGEND_ORANGE = "#FF6600"

# Four bins from the coverage_4g_data1 RSRP legend. Best is listed first.
# A value below -150 stays in the red bin; a value above 0 stays in the dark-green bin.
MAP_RSRP_LABELS = [
    "-95 to 0",
    "-110 to -95",
    "-115 to -110",
    "-120 to -115",
    "-150 to -120",
]
# Light red includes -115 and -120. Yellow is only better than -115.
MAP_RSRP_THRESHOLDS = [-95, -110, -115, -120]
MAP_RSRP_COLORS = ["#008000", "#00FF00", "#FFFF00", "#FF6666", "#FF0000"]

# Source table listed RSRQ worst-first and wrote "X < -5" for the blue bin.
# Blue is the best colour, so that row is X >= -5; red is X < -20.
# Three bins from the coverage_4g_data1 RSRQ legend. Best is listed first.
# A value below -30 stays in the red bin; a value above 0 stays in the green bin.
MAP_RSRQ_LABELS = [
    "-10 to 0",
    "-17 to -10",
    "-30 to -17",
]
MAP_RSRQ_THRESHOLDS = [-10, -17]
MAP_RSRQ_COLORS = ["#00FF00", "#FFFF00", "#FF0000"]

MAP_SINR_LABELS = [
    "15 <= X < Max",
    "10 <= X < 15",
    "5 <= X < 10",
    "0 <= X < 5",
    "-5 <= X < 0",
    "Min < X < -5",
]
MAP_SINR_THRESHOLDS = [15, 10, 5, 0, -5]
MAP_SINR_COLORS = [
    LEGEND_BLUE,
    LEGEND_DARK_GREEN,
    LEGEND_CYAN,
    LEGEND_YELLOW,
    LEGEND_ORANGE,
    LEGEND_RED,
]

# Cover / summary tables use the same bins as the maps.
RSRP_LABELS = MAP_RSRP_LABELS
RSRP_COLORS = MAP_RSRP_COLORS


def _map_class(series: pd.Series, thresholds: list[float], labels: list[str]) -> pd.Series:
    """Assign legend labels; thresholds are inclusive lower bounds, best first."""
    v = pd.to_numeric(series, errors="coerce")
    classified = np.select(
        [v >= t for t in thresholds],
        labels[:-1],
        default=labels[-1],
    )
    return pd.Series(classified, index=v.index).where(v.notna(), other=pd.NA)


def rsrp_map_class(series: pd.Series) -> pd.Series:
    """Classify RSRP. Light red starts at -115 and includes -115 and -120."""
    v = pd.to_numeric(series, errors="coerce")
    classified = np.select(
        [v >= -95, v >= -110, v > -115, v >= -120],
        MAP_RSRP_LABELS[:-1],
        default=MAP_RSRP_LABELS[-1],
    )
    return pd.Series(classified, index=v.index).where(v.notna(), other=pd.NA)


def rsrq_map_class(series: pd.Series) -> pd.Series:
    return _map_class(series, MAP_RSRQ_THRESHOLDS, MAP_RSRQ_LABELS)


def sinr_map_class(series: pd.Series) -> pd.Series:
    return _map_class(series, MAP_SINR_THRESHOLDS, MAP_SINR_LABELS)


MAP_CLASSIFIERS = {
    "RSRP": rsrp_map_class,
    "RSRQ": rsrq_map_class,
    "SINR": sinr_map_class,
}


def rsrp_range_counts(series: pd.Series) -> pd.Series:
    return rsrp_map_class(series).value_counts(dropna=True).reindex(MAP_RSRP_LABELS, fill_value=0)


def hide_map_axes(ax) -> None:
    """Bare coverage map: no lon/lat ticks, matching the RSRP snapshot style."""
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.tick_params(labelbottom=False, labelleft=False, length=0)
    for spine in ax.spines.values():
        spine.set_visible(False)


# RF-planner three-arm site widget: outline only, no fill inside the petals.
SITE_PIE_EDGE = "#2B2B2B"
SITE_PIE_ARM_WIDTH_FRAC = 0.16
SITE_PIE_RADIUS_KM = 0.90
_SITE_CELLS: pd.DataFrame | None = None
_SERVING_INDEX: tuple[dict, dict] | None = None
_SERVING_LIST: list[dict] | None = None
_SERVING_GRID: dict[tuple[int, int], list[dict]] | None = None


def _most_common_azimuth(series: pd.Series) -> float:
    """Mode of azimuth; ties keep the first most-frequent value."""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return float("nan")
    rounded = s.round().astype(int)
    return float(rounded.value_counts().idxmax())


def load_site_sectors(path: Path | None = None, layer: str | None = None) -> pd.DataFrame:
    """One row per site+sector: most-common azimuth when a sector has many cells."""
    global _SITE_CELLS
    if _SITE_CELLS is None:
        db_path = path or SITE_DB_PATH
        if not db_path.exists():
            _SITE_CELLS = pd.DataFrame()
            return _SITE_CELLS
        if db_path.suffix.lower() == ".xlsb":
            from pyxlsb import open_workbook

            rows: list[list] = []
            cols: list[str] = []
            with open_workbook(str(db_path)) as wb:
                sheet = wb.get_sheet(wb.sheets[0])
                for i, row in enumerate(sheet.rows()):
                    vals = [c.v for c in row]
                    if i == 0:
                        cols = [str(v) for v in vals]
                        continue
                    rows.append(vals)
            cells = pd.DataFrame(rows, columns=cols)
        else:
            cells = pd.read_excel(db_path, sheet_name=0, engine="openpyxl")
        cells = cells.rename(columns={c: c.strip() for c in cells.columns})
        if "Site" in cells.columns and "SiteName" not in cells.columns:
            cells = cells.rename(columns={"Site": "SiteName"})
        for col in ("Lat", "Lon", "Azimuth", "EARFCN"):
            if col in cells.columns:
                cells[col] = pd.to_numeric(cells[col], errors="coerce")
        need = {"SiteName", "Sector", "Lat", "Lon", "Azimuth"}
        if not need.issubset(cells.columns):
            _SITE_CELLS = pd.DataFrame()
            return _SITE_CELLS
        cells = cells.dropna(subset=["SiteName", "Sector", "Lat", "Lon"])
        if "EARFCN" in cells.columns:
            cells["layer"] = cells["EARFCN"].map(earfcn_to_layer)
        else:
            cells["layer"] = ""
        _SITE_CELLS = cells
    cells = _SITE_CELLS
    if cells.empty:
        return cells
    if layer:
        cells = cells[cells["layer"] == layer]
    grouped = (
        cells.groupby(["SiteName", "Sector"], sort=False)
        .agg(
            lat=("Lat", "median"),
            lon=("Lon", "median"),
            azimuth=("Azimuth", _most_common_azimuth),
            n_cells=("Sector", "size"),
        )
        .reset_index()
    )
    grouped["n_sectors"] = grouped.groupby("SiteName")["Sector"].transform("nunique")
    return grouped.dropna(subset=["lat", "lon", "azimuth"])


def _sites_near_drive(sectors: pd.DataFrame, drive: pd.DataFrame, radius_km: float = 2.5) -> pd.DataFrame:
    """Keep sites within about radius_km of a logged drive sample."""
    if sectors.empty or drive.empty:
        return sectors
    grid = 0.01  # ~1.1 km
    lat = pd.to_numeric(drive["Latitude"], errors="coerce")
    lon = pd.to_numeric(drive["Longitude"], errors="coerce")
    mask = lat.notna() & lon.notna()
    cells = set(
        zip(
            np.floor(lat[mask] / grid).astype(int),
            np.floor(lon[mask] / grid).astype(int),
        )
    )
    steps = max(1, int(round(radius_km / (grid * 111.32))))
    expanded: set[tuple[int, int]] = set()
    for y, x in cells:
        for dy in range(-steps, steps + 1):
            for dx in range(-steps, steps + 1):
                if dy * dy + dx * dx <= steps * steps + 1:
                    expanded.add((y + dy, x + dx))
    sy = np.floor(sectors["lat"] / grid).astype(int)
    sx = np.floor(sectors["lon"] / grid).astype(int)
    keep = [(int(y), int(x)) in expanded for y, x in zip(sy, sx)]
    return sectors.loc[keep].copy()


def _arm_path(lat: float, lon: float, azimuth: float, length_km: float, width_frac: float = SITE_PIE_ARM_WIDTH_FRAC) -> MPath:
    """Hollow rectangular arm along azimuth (0 = north). Inner edge is left open so three arms form a Y."""
    r = length_km / 111.32
    half_w = r * float(width_frac) / 2.0
    az = np.radians(float(azimuth) % 360.0)
    ux, uy = np.sin(az), np.cos(az)
    px, py = np.cos(az), -np.sin(az)
    inner_l = (lon - half_w * px, lat - half_w * py)
    outer_l = (lon + r * ux - half_w * px, lat + r * uy - half_w * py)
    outer_r = (lon + r * ux + half_w * px, lat + r * uy + half_w * py)
    inner_r = (lon + half_w * px, lat + half_w * py)
    return MPath([inner_l, outer_l, outer_r, inner_r], [MPath.MOVETO, MPath.LINETO, MPath.LINETO, MPath.LINETO])


def _pie_radius_km(ax, default_km: float = SITE_PIE_RADIUS_KM) -> float:
    xlim = ax.get_xlim()
    ylim = ax.get_ylim()
    span_km = min((xlim[1] - xlim[0]) * 111.32, (ylim[1] - ylim[0]) * 111.32)
    if span_km <= 0:
        return default_km
    return float(np.clip(span_km * 0.012, 0.35, 1.15))


def draw_site_pies(
    ax,
    drive: pd.DataFrame,
    radius_km: float | None = None,
    layer: str | None = None,
) -> list:
    """Draw outline-only three-arm site pies (grouped by sector name). No labels, no fill."""
    sectors = load_site_sectors(layer=layer)
    if sectors.empty:
        return []
    xlim = ax.get_xlim()
    ylim = ax.get_ylim()
    vis = sectors[sectors["lat"].between(ylim[0], ylim[1]) & sectors["lon"].between(xlim[0], xlim[1])]
    vis = _sites_near_drive(vis, drive, radius_km=2.5)
    if vis.empty:
        return []
    radius = _pie_radius_km(ax) if radius_km is None else float(radius_km)
    for _site_name, part in vis.groupby("SiteName", sort=False):
        part = part.drop_duplicates(subset=["Sector"])
        lat = float(part["lat"].median())
        lon = float(part["lon"].median())
        for az in part["azimuth"].tolist():
            ax.add_patch(
                PathPatch(
                    _arm_path(lat, lon, float(az), radius),
                    facecolor="none",
                    edgecolor=SITE_PIE_EDGE,
                    linewidth=0.75,
                    joinstyle="miter",
                    zorder=3,
                )
            )
    return []


def _azimuth_to_mpl(azimuth: float) -> float:
    """Compass azimuth (0 = north, clockwise) as a matplotlib angle (0 = east, CCW)."""
    return (90.0 - float(azimuth)) % 360.0


def _bearing_deg(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Compass bearing from point 1 to point 2, degrees clockwise from north."""
    dlon = np.radians(lon2 - lon1)
    lat1r, lat2r = np.radians(lat1), np.radians(lat2)
    y = np.sin(dlon) * np.cos(lat2r)
    x = np.cos(lat1r) * np.sin(lat2r) - np.sin(lat1r) * np.cos(lat2r) * np.cos(dlon)
    return float((np.degrees(np.arctan2(y, x)) + 360.0) % 360.0)


def _ang_diff(a: float, b: float) -> float:
    return abs((float(a) - float(b) + 180.0) % 360.0 - 180.0)


def _sector_wedges(part: pd.DataFrame) -> list[tuple[str, float, float, float]]:
    """One clockwise slice per sector: name, azimuth, clockwise start, clockwise end."""
    rows = part.dropna(subset=["azimuth"]).drop_duplicates(subset=["Sector"])
    if rows.empty:
        return []
    azs = sorted((float(a) % 360.0, str(name)) for name, a in zip(rows["Sector"], rows["azimuth"]))
    n = len(azs)
    if n == 1:
        az, name = azs[0]
        return [(name, az, (az - 180.0) % 360.0, (az + 180.0) % 360.0)]
    wedges = []
    for i, (az, name) in enumerate(azs):
        prev_az = azs[(i - 1) % n][0]
        next_az = azs[(i + 1) % n][0]
        incoming = (prev_az + ((az - prev_az) % 360.0) / 2.0) % 360.0
        outgoing = (az + ((next_az - az) % 360.0) / 2.0) % 360.0
        wedges.append((name, az, incoming, outgoing))
    return wedges


# Bad-spot site mark: the same unfilled Y as a planning widget. Cyan stroke only.
BAD_SPOT_PIE_EDGE = "#00E5EE"
# The serial bad-spot maps draw a yellow line from the serving cell to each red
# bin. This gate is only the older azimuth line, used when those samples are
# not passed in. Half of a 65° sector beam. A three-sector site is always
# within 60° of some arm, so a wider gate would put a line on every nearby site.
BAD_SPOT_AIM_DEG = 32.0


def _serving_cell_index() -> tuple[dict, dict]:
    """Serving cells on L900, L1800, L2100 and L2600, keyed by eNB, cell and EARFCN.

    A cell with no eNB id or cell id is left out. The exact key prefers the
    EARFCN of the sample. The pair key is the fallback when the log has a Cell Id
    but the EARFCN is missing.
    """
    global _SERVING_INDEX, _SERVING_LIST
    if _SERVING_INDEX is not None:
        return _SERVING_INDEX
    load_site_sectors()
    cells = _SITE_CELLS if _SITE_CELLS is not None else pd.DataFrame()
    exact: dict[tuple[int, int, int], dict] = {}
    pair: dict[tuple[int, int], dict] = {}
    listed: list[dict] = []
    listed_pairs: set[tuple[int, int]] = set()
    if cells is not None and not cells.empty and {"eNBId", "cellId", "Lat", "Lon"}.issubset(cells.columns):
        for _, row in cells.iterrows():
            layer = str(row.get("layer") or "")
            if layer not in REPORT_LAYERS:
                continue
            enb = pd.to_numeric(row.get("eNBId"), errors="coerce")
            cid = pd.to_numeric(row.get("cellId"), errors="coerce")
            lat = pd.to_numeric(row.get("Lat"), errors="coerce")
            lon = pd.to_numeric(row.get("Lon"), errors="coerce")
            if pd.isna(enb) or pd.isna(cid) or pd.isna(lat) or pd.isna(lon):
                continue
            az = pd.to_numeric(row.get("Azimuth"), errors="coerce")
            et = pd.to_numeric(row.get("E_tilt"), errors="coerce")
            mt = pd.to_numeric(row.get("M_tilt"), errors="coerce")
            ht = pd.to_numeric(row.get("Height"), errors="coerce")
            rec = {
                "lat": float(lat),
                "lon": float(lon),
                "sector": str(row.get("Sector") or "").strip(),
                "layer": layer,
                "site": str(row.get("SiteName") or "").strip(),
                "name": str(row.get("EUtranCellId") or "").strip(),
                "azimuth": None if pd.isna(az) else float(az) % 360.0,
                "etilt": None if pd.isna(et) else float(et),
                "mtilt": None if pd.isna(mt) else float(mt),
                "height": None if pd.isna(ht) else float(ht),
            }
            key = (int(enb), int(cid))
            previous = pair.get(key)
            if previous is None or (previous.get("azimuth") is None and rec["azimuth"] is not None):
                pair[key] = rec
            ear = pd.to_numeric(row.get("EARFCN"), errors="coerce")
            if pd.notna(ear):
                exact[(*key, int(ear))] = rec
                listed.append(rec)
                listed_pairs.add(key)
    for key, rec in pair.items():
        if key not in listed_pairs:
            listed.append(rec)
    _SERVING_INDEX = (exact, pair)
    _SERVING_LIST = listed
    return _SERVING_INDEX


def _serving_cell_list() -> list[dict]:
    """One record per L900/L1800/L2100/L2600 cell that has a Cell Id in the dump."""
    _serving_cell_index()
    return _SERVING_LIST or []


def _serving_cells_near(lat: float, lon: float, km: float = 8.0) -> list[dict]:
    """Dump cells whose site is inside a box around this point."""
    global _SERVING_GRID
    if _SERVING_GRID is None:
        grid: dict[tuple[int, int], list[dict]] = {}
        for rec in _serving_cell_list():
            key = (int(rec["lat"] / 0.08), int(rec["lon"] / 0.08))
            grid.setdefault(key, []).append(rec)
        _SERVING_GRID = grid
    reach = int(km / (0.08 * 111.32)) + 1
    cy, cx = int(lat / 0.08), int(lon / 0.08)
    found: list[dict] = []
    for dy in range(-reach, reach + 1):
        for dx in range(-reach, reach + 1):
            found.extend(_SERVING_GRID.get((cy + dy, cx + dx), ()))
    return found


def _tilt_step(
    etilt: float | None, mtilt: float | None, delta: float
) -> tuple[float | None, float | None, float | None, float | None]:
    """Apply a tilt change to E_tilt first. Leftover degrees move M_tilt.

    Positive delta is downtilt. A blank tilt is not treated as 0°. E_tilt stays
    within 0 to 12 degrees and M_tilt stays within 0 to 8 degrees.
    """
    spill = delta
    new_e = etilt
    if etilt is not None:
        new_e = float(etilt) + delta
        spill = 0.0
        if new_e < 0.0:
            spill = new_e
            new_e = 0.0
        elif new_e > 12.0:
            spill = new_e - 12.0
            new_e = 12.0
    new_m = mtilt
    if mtilt is not None and abs(spill) >= 0.05:
        new_m = min(8.0, max(0.0, float(mtilt) + spill))
    return etilt, mtilt, new_e, new_m


def _deg_text(value: float | None) -> str:
    """Degrees for a solution sentence.

    Proposed azimuth, E_tilt and M_tilt are then rounded to the nearest whole
    degree, and half a degree rounds up. The dump value stays as measured.
    """
    if value is None:
        return "blank"
    if abs(value - round(value)) < 0.05:
        return str(int(round(value)))
    return f"{value:.1f}"


def _tilt_now(cell: dict) -> str:
    e = "E_tilt blank" if cell.get("etilt") is None else f"E_tilt {_deg_text(cell['etilt'])}°"
    m = "M_tilt blank" if cell.get("mtilt") is None else f"M_tilt {_deg_text(cell['mtilt'])}°"
    return f"{e} and {m}"


def _tilt_bit(label: str, old: float | None, new: float | None) -> str:
    if old is None:
        return f"{label} is blank in the dump"
    if new is None or abs(new - old) < 0.05:
        return f"{label} stays {_deg_text(old)}°"
    return f"{label} {_deg_text(old)}° to {_deg_text(new)}°"


def _tilt_line(cell: dict, delta: float, action: str) -> str:
    """One solution row. delta > 0 downtilt, delta < 0 uptilt."""
    old_e, old_m, new_e, new_m = _tilt_step(cell.get("etilt"), cell.get("mtilt"), delta)
    name = f"{cell.get('sector') or cell.get('name') or 'cell'} {cell.get('layer') or ''}".strip()
    e_same = old_e is None or new_e is None or abs(new_e - old_e) < 0.05
    m_same = old_m is None or new_m is None or abs(new_m - old_m) < 0.05
    if e_same and m_same:
        return (
            f"{name}: {_tilt_bit('E_tilt', old_e, new_e)}. {_tilt_bit('M_tilt', old_m, new_m)}. "
            f"There is no further {action} left inside 0–12° E_tilt and 0–8° M_tilt."
        )
    return (
        f"{name}: {_tilt_bit('E_tilt', old_e, new_e)}. {_tilt_bit('M_tilt', old_m, new_m)}. "
        f"{action.capitalize()}."
    )


# L900 planning reach. Higher bands shorten in proportion to frequency, which is
# the free-space rule (distance for the same loss scales as 1/frequency).
L900_REACH_KM = 8.0
_LAYER_FREQ_MHZ = {"L900": 900.0, "L1800": 1800.0, "L2100": 2100.0, "L2600": 2600.0}
_POP_GRID: tuple | None = None


def _layer_reach_km(layer: str) -> float:
    freq = _LAYER_FREQ_MHZ.get(layer or "", 1800.0)
    return L900_REACH_KM * 900.0 / freq


def _population_grid() -> tuple[np.ndarray, float, float, float, float]:
    """WorldPop 2020 people per km² for the Bogura drive, about 1 km cells."""
    global _POP_GRID
    if _POP_GRID is None:
        path = ROOT / "data" / "bogura_worldpop_2020_1km.npz"
        loaded = np.load(path)
        grid = loaded["pop"].astype(np.float32)
        grid = np.where(np.isfinite(grid) & (grid < 1.0e6), grid, 0.0)
        _POP_GRID = (
            grid,
            float(loaded["xmin"]),
            float(loaded["ymin"]),
            float(loaded["xmax"]),
            float(loaded["ymax"]),
        )
    return _POP_GRID


def population_per_km2(lat: float, lon: float) -> float:
    grid, xmin, ymin, xmax, ymax = _population_grid()
    if not (xmin <= lon <= xmax and ymin <= lat <= ymax):
        return 0.0
    height, width = grid.shape
    col = int(np.clip((lon - xmin) / (xmax - xmin) * (width - 1), 0, width - 1))
    row = int(np.clip((ymax - lat) / (ymax - ymin) * (height - 1), 0, height - 1))
    return float(grid[row, col])


def _people(value: float) -> str:
    return f"{value:,.0f}/km²"


def _destination(lat: float, lon: float, bearing: float, dist_km: float) -> tuple[float, float]:
    br = math.radians(bearing % 360.0)
    lat2 = lat + (dist_km / 111.32) * math.cos(br)
    lon2 = lon + (dist_km / max(_km_per_deg_lon(lat), 0.2)) * math.sin(br)
    return lat2, lon2


def _signed_deg(azimuth: float, bearing: float) -> float:
    """Bearing relative to azimuth, degrees, negative to the left."""
    return (float(bearing) - float(azimuth) + 180.0) % 360.0 - 180.0


def _ground_tilt_deg(height_m: float | None, dist_km: float) -> float:
    height = 30.0 if height_m is None else max(10.0, float(height_m))
    return math.degrees(math.atan(height / max(dist_km, 0.05) / 1000.0))


def _wedge_population(lat: float, lon: float, azimuth: float) -> float:
    """People per km² along a sector bearing, averaged at 1.5, 3 and 4.5 km."""
    values = []
    for dist in (1.5, 3.0, 4.5):
        lat2, lon2 = _destination(lat, lon, azimuth, dist)
        values.append(population_per_km2(lat2, lon2))
    return float(sum(values) / len(values))


def _sectors_near(lat: float, lon: float, km: float = 8.5) -> list[dict]:
    """One record per sector around a point. L900 is kept when the sector has it."""
    rank = {"L900": 0, "L1800": 1, "L2100": 2, "L2600": 3}
    best: dict[tuple, dict] = {}
    for rec in _serving_cells_near(lat, lon, km):
        if rec.get("azimuth") is None:
            continue
        away = float(_equirect_km(rec["lat"], rec["lon"], lat, lon))
        if away > km:
            continue
        key = (rec.get("sector"), round(float(rec["lat"]), 4), round(float(rec["lon"]), 4))
        current = best.get(key)
        if current is None or rank.get(rec.get("layer"), 9) < rank.get(current.get("layer"), 9):
            best[key] = rec
    return list(best.values())


def _blank_segments(samples: pd.DataFrame, spot: pd.Series) -> list[dict]:
    """Driven stretches inside the circle where the map has no plot.

    A normal sample is a few metres from the next one. A jump of 80 m to 1.6 km
    at driving speed, on the same log, is a blank: the road was driven and
    nothing was plotted. Those metres are part of this bad spot.
    """
    if samples is None or samples.empty or "outline_width_km" not in getattr(spot, "index", []):
        return []
    if pd.isna(spot.get("outline_width_km")):
        return []
    need = ["Latitude", "Longitude", "Time"]
    if any(col not in samples.columns for col in need):
        return []
    work = samples.dropna(subset=need).sort_values("Time")
    if len(work) < 2:
        return []
    lat = work["Latitude"].to_numpy(dtype=float)
    lon = work["Longitude"].to_numpy(dtype=float)
    time_s = pd.to_datetime(work["Time"]).to_numpy(dtype="datetime64[ns]").astype("int64") / 1e9
    source = work["Source"].to_numpy() if "Source" in work.columns else np.zeros(len(work))
    lat0 = float(spot["outline_lat"])
    lon0 = float(spot["outline_lon"])
    half_w = float(spot["outline_width_km"]) / _km_per_deg_lon(lat0) / 2.0
    half_h = float(spot["outline_height_km"]) / 111.32 / 2.0
    if half_w <= 0 or half_h <= 0:
        return []
    same = source[:-1] == source[1:]
    dt = time_s[1:] - time_s[:-1]
    dist_m = _equirect_km(lat[:-1], lon[:-1], lat[1:], lon[1:]) * 1000.0
    speed = dist_m / np.maximum(dt, 1e-6) * 3.6
    mid_lat = (lat[:-1] + lat[1:]) / 2.0
    mid_lon = (lon[:-1] + lon[1:]) / 2.0
    inside = ((mid_lon - lon0) / half_w) ** 2 + ((mid_lat - lat0) / half_h) ** 2 <= 1.0
    keep = (
        same
        & (dt >= 8.0)
        & (dt <= 150.0)
        & (dist_m >= 80.0)
        & (dist_m <= 1600.0)
        & (speed >= 15.0)
        & (speed <= 120.0)
        & inside
    )
    idx = np.flatnonzero(keep)
    order = idx[np.argsort(-dist_m[idx])]
    return [
        {
            "lat": float(mid_lat[i]),
            "lon": float(mid_lon[i]),
            "lat1": float(lat[i]),
            "lon1": float(lon[i]),
            "lat2": float(lat[i + 1]),
            "lon2": float(lon[i + 1]),
            "length_m": float(dist_m[i]),
        }
        for i in order
    ]


def _not_required(reason: str) -> str:
    return "Not required. " + reason


# A sector this close to the bearing is already facing the red road.
FACE_ON_ROAD_DEG = 12.0
# New-site sectors stay at least this far apart. One sector aims at the bad spot.
NEW_SITE_MIN_GAP_DEG = 40.0
_SEEN_SECTORS: dict[str, set[str]] = {}


def remember_seen_sectors(mode: str, samples: pd.DataFrame) -> None:
    """Sectors decoded in this mode. Idle is every scanner cell. Active is the serving cell."""
    if samples is None or samples.empty or "Cell Id" not in samples.columns:
        _SEEN_SECTORS[mode] = set()
        return
    work = samples.dropna(subset=["Cell Id"])
    if "DL EARFCN" in work.columns:
        pairs = work[["Cell Id", "DL EARFCN"]].drop_duplicates()
        ears = pairs["DL EARFCN"]
    else:
        pairs = work[["Cell Id"]].drop_duplicates()
        ears = [None] * len(pairs)
    seen: set[str] = set()
    for eci, ear in zip(pairs["Cell Id"], ears):
        cell = _resolve_serving_cell(eci, ear)
        sector = "" if cell is None else str(cell.get("sector") or "")
        if sector:
            seen.add(sector)
    _SEEN_SECTORS[mode] = seen


# Azimuth and tilt are used when a sector is within about 2 to 2.5 km.
# A 4th sector is required in the null inside 1.5 km. It also replaces a new site
# when a host inside 3 km can aim at the bad spot. Downtilt is limited to 2°.
TUNE_KM = 2.5
FOURTH_MUST_KM = 1.5
FOURTH_RISK_KM = 2.0
FOURTH_SAVE_KM = 3.0
DOWNTILT_MAX_DEG = 2.0
# A cell serving the red road from this far, and much farther than the nearest
# site, is overshooting. It is downtilted before the nearest sector is moved.
OVERSHOOT_KM = 4.0


def _sites_around(sectors: list[dict], lat: float, lon: float) -> list[dict]:
    """One record per site, with the distance and whether the spot sits in a null."""
    groups: dict[tuple, list[dict]] = {}
    for rec in sectors:
        if rec.get("azimuth") is None:
            continue
        key = (
            rec.get("site") or rec.get("sector") or "",
            round(float(rec["lat"]), 4),
            round(float(rec["lon"]), 4),
        )
        groups.setdefault(key, []).append(rec)
    sites = []
    for (name, _lat, _lon), recs in groups.items():
        dist = float(_equirect_km(recs[0]["lat"], recs[0]["lon"], lat, lon))
        bearing = _bearing_deg(recs[0]["lat"], recs[0]["lon"], lat, lon)
        azimuths = [float(rec["azimuth"]) for rec in recs]
        nearest_off = min(abs(_signed_deg(az, bearing)) for az in azimuths)
        sites.append({
            "name": name or str(recs[0].get("sector") or "site"),
            "lat": float(recs[0]["lat"]),
            "lon": float(recs[0]["lon"]),
            "dist": dist,
            "bearing": bearing,
            "nearest_off": nearest_off,
            "in_null": len(azimuths) >= 2 and 30.0 < nearest_off <= 70.0,
            "recs": recs,
            "height": next((rec.get("height") for rec in recs if rec.get("height")), None),
        })
    sites.sort(key=lambda item: item["dist"])
    return sites


def _tune_candidate(rec: dict, dist: float, bearing: float, spot_pop: float) -> dict | None:
    """Azimuth and tilt for one sector within 2.5 km, or nothing if it cannot cover."""
    reach = _layer_reach_km(rec.get("layer") or "")
    if dist > min(TUNE_KM, reach + 0.3):
        return None
    signed = _signed_deg(float(rec["azimuth"]), bearing)
    off = abs(signed)
    if off > 50.0:
        return None
    desired = _ground_tilt_deg(rec.get("height"), dist)
    current = 0.0 if rec.get("etilt") is None else float(rec["etilt"])
    raw = desired - current
    # A downtilt of more than 2° shrinks the existing footprint, so it is not proposed.
    overshoot = raw > DOWNTILT_MAX_DEG
    if raw > DOWNTILT_MAX_DEG:
        delta = DOWNTILT_MAX_DEG
    elif raw < -6.0:
        delta = -6.0
    else:
        delta = raw
    tilt_change = abs(delta) >= 0.8
    turn = 0.0
    old_pop = new_pop = None
    if off > 12.0:
        turn = math.copysign(min(max(off - 8.0, 0.0), 25.0), signed)
    if abs(turn) >= 8.0:
        new_az = (float(rec["azimuth"]) + turn) % 360.0
        old_pop = _wedge_population(rec["lat"], rec["lon"], rec["azimuth"])
        new_pop = _wedge_population(rec["lat"], rec["lon"], new_az)
        if new_pop + 0.5 * spot_pop < old_pop * 0.90:
            turn = 0.0
            old_pop = new_pop = None
    if off > 32.0 and abs(turn) < 8.0:
        return None
    if not tilt_change and abs(turn) < 8.0:
        return None
    return {
        "rec": rec,
        "dist": dist,
        "off": off,
        "delta": delta if tilt_change else 0.0,
        "turn": turn,
        "old_pop": old_pop,
        "new_pop": new_pop,
        "overshoot": overshoot,
    }


def _proposal_draw(kind: str, lat: float, lon: float, azimuths: list[float], target_lat: float, target_lon: float, label: str) -> dict:
    """Where to draw a proposed site or 4th sector on the bad-spot map."""
    return {
        "kind": kind,
        "lat": lat,
        "lon": lon,
        "azimuths": azimuths,
        "target_lat": target_lat,
        "target_lon": target_lon,
        "label": label,
    }


def _fourth_text(site: dict, reason: str = "null") -> str:
    tilt = int(round(min(10.0, max(0.0, _ground_tilt_deg(site.get("height"), site["dist"])))))
    base = (
        f"New L900 sector on {site['name']}, azimuth {site['bearing']:.0f}°, "
        f"E_tilt {tilt}°, M_tilt 0°. The site is {site['dist']:.1f} km away and the spot "
        f"is {site['nearest_off']:.0f}° off the nearest existing sector. "
    )
    if reason == "downtilt":
        why = (
            "A downtilt of more than 2° would shrink the existing coverage, so this sector "
            "is aimed at the bad spot from a site within 1 km. Existing sectors are not moved."
        )
    elif reason == "save":
        why = (
            "A new site is not added. This sector is aimed at the bad spot, and at the blank "
            "inside the circle, from the existing site. Existing sectors are not moved."
        )
    else:
        why = "This sector sits in that null. Existing sectors are not moved."
    return base + why


def _ep_num(value: float | None) -> float | None:
    if value is None:
        return None
    number = float(value)
    if abs(number - round(number)) < 0.05:
        return float(int(round(number)))
    return round(number, 1)


def _fourth_action(site: dict) -> dict:
    """One new L900 sector on an existing site, for the engineering-parameter sheet."""
    tilt = int(round(min(10.0, max(0.0, _ground_tilt_deg(site.get("height"), site["dist"])))))
    height = site.get("height")
    return {
        "site": site["name"],
        "sector": "New L900",
        "cell": "New L900",
        "height": None,
        "azimuth": None,
        "et": None,
        "mt": None,
        "p_height": _ep_num(height if height else 30.0),
        "p_azimuth": _ep_num(float(site["bearing"])),
        "p_et": float(tilt),
        "p_mt": 0.0,
        "action": "4th sector, L900",
    }


def _tune_action(item: dict) -> dict:
    """Existing cell, with the proposed azimuth and tilt beside the current values."""
    rec = item["rec"]
    azimuth = float(rec["azimuth"])
    new_az = azimuth
    if abs(item["turn"]) >= 8.0:
        new_az = (azimuth + item["turn"]) % 360.0
    _old_e, _old_m, new_e, new_m = _tilt_step(rec.get("etilt"), rec.get("mtilt"), item["delta"])
    if abs(item["delta"]) < 0.8:
        new_e, new_m = rec.get("etilt"), rec.get("mtilt")
    cell = str(rec.get("name") or "").strip() or f"{rec.get('sector')} {rec.get('layer')}".strip()
    return {
        "site": item["site"],
        "sector": str(rec.get("sector") or "").strip(),
        "cell": cell,
        "height": _ep_num(rec.get("height")),
        "azimuth": _ep_num(azimuth),
        "et": _ep_num(rec.get("etilt")),
        "mt": _ep_num(rec.get("mtilt")),
        "p_height": _ep_num(rec.get("height")),
        "p_azimuth": _ep_num(new_az),
        "p_et": _ep_num(new_e),
        "p_mt": _ep_num(new_m),
        "action": "Tilt and azimuth",
    }


def _azimuth_gap(left: float, right: float) -> float:
    """Smallest angle between two azimuths, in degrees."""
    return abs(_signed_deg(left, right))


def _new_site_azimuths(site_lat: float, site_lon: float, face: float) -> list[tuple[float, float]]:
    """One sector on the bad spot, two toward the higher population, gaps at least 40°.

    Returns (azimuth, people per km²) with the bad-spot sector first. Equal spacing
    is used only when the population grid has no people to choose between.
    """
    face = face % 360.0
    face_pop = _wedge_population(site_lat, site_lon, face)
    candidates = []
    for deg in range(0, 360, 5):
        if _azimuth_gap(face, deg) < NEW_SITE_MIN_GAP_DEG:
            continue
        candidates.append((float(deg), _wedge_population(site_lat, site_lon, deg)))
    best: tuple[float, float, float] | None = None
    for index, (az1, pop1) in enumerate(candidates):
        for az2, pop2 in candidates[index + 1 :]:
            if _azimuth_gap(az1, az2) < NEW_SITE_MIN_GAP_DEG:
                continue
            score = pop1 + pop2
            if best is None or score > best[0]:
                best = (score, az1, az2)
    if best is None or best[0] <= 0:
        offsets = (0.0, 120.0, 240.0)
        return [( (face + offset) % 360.0, _wedge_population(site_lat, site_lon, (face + offset) % 360.0)) for offset in offsets]
    second, third = best[1], best[2]
    return [
        (face, face_pop),
        (second, _wedge_population(site_lat, site_lon, second)),
        (third, _wedge_population(site_lat, site_lon, third)),
    ]


def _new_site_text(lat: float, lon: float, red_lat: float, red_lon: float, sectors: list[dict]) -> tuple[str, str]:
    site_lat = lat + 0.18 / 111.32
    site_lon = lon
    face = _bearing_deg(site_lat, site_lon, red_lat, red_lon)
    planned = _new_site_azimuths(site_lat, site_lon, face)
    azimuths = [az for az, _pop in planned]
    near = min(
        (float(_equirect_km(rec["lat"], rec["lon"], site_lat, site_lon)) for rec in sectors),
        default=5.0,
    )
    face_tilt = int(round(min(10.0, max(2.0, _ground_tilt_deg(30.0, 0.25)))))
    outer_tilt = int(round(min(8.0, max(2.0, _ground_tilt_deg(30.0, max(near / 2.0, 0.6))))))
    ordered = sorted(azimuths)
    gaps = [((ordered[(i + 1) % 3] - az) % 360.0) for i, az in enumerate(ordered)]
    gap_text = ", ".join(f"{gap:.0f}°" for gap in gaps)
    sentences = []
    for index, (az, pop) in enumerate(planned):
        tilt = face_tilt if index == 0 else outer_tilt
        if index == 0:
            aim = "aims at the bad spot"
        else:
            aim = f"aims toward the higher population, {_people(pop)}"
        sentences.append(
            f"New L900 sector at {az:.0f}°, E_tilt {tilt}°, M_tilt 0°, {aim}."
        )
    text = (
        " ".join(sentences)
        + f" The site is at {site_lat:.5f}, {site_lon:.5f}. "
        + f"The nearest existing site is {near:.1f} km away. "
        + f"The gaps between the sectors are {gap_text}."
    )
    where = f"New site {site_lat:.5f}, {site_lon:.5f}"
    actions = []
    for index, az in enumerate(azimuths):
        actions.append({
            "site": where,
            "sector": f"S{index + 1}",
            "cell": "New L900",
            "height": None,
            "azimuth": None,
            "et": None,
            "mt": None,
            "p_height": 30.0,
            "p_azimuth": _ep_num(az),
            "p_et": float(face_tilt if index == 0 else outer_tilt),
            "p_mt": 0.0,
            "action": "New 3-sector site",
        })
    return "New site", text, _proposal_draw(
        "new", site_lat, site_lon, azimuths, red_lat, red_lon, "New site",
    ), actions


def _choose_coverage_fix(
    sites: list[dict],
    sectors: list[dict],
    spot_pop: float,
    target_lat: float,
    target_lon: float,
    red_lat: float,
    red_lon: float,
) -> dict:
    """Pick one final solution from the distance rules.

    A null inside 1.5 km is always a 4th L900 sector. Downtilt is limited to 2°.
    When the spot needs more downtilt, a 4th sector within 1 km is used instead.
    A 4th sector also replaces a new site when a host inside 3 km can aim at the
    bad spot. A new 3-sector site is kept only when no such host exists.
    """
    nearby = [site for site in sites if site["dist"] <= TUNE_KM]
    nulls = [site for site in sites if site["in_null"]]
    close_null = next((site for site in nulls if site["dist"] <= FOURTH_MUST_KM), None)
    usable_null = next((site for site in nulls if site["dist"] <= FOURTH_RISK_KM), None)
    nearest = sites[0] if sites else None

    tunes = []
    for site in nearby:
        bearing = site["bearing"]
        for rec in site["recs"]:
            hit = _tune_candidate(rec, site["dist"], bearing, spot_pop)
            if hit is not None:
                hit["site"] = site["name"]
                tunes.append(hit)
    tunes.sort(key=lambda item: (
        0 if item["rec"].get("layer") == "L900" else 1,
        item["dist"],
        item["off"],
    ))

    def pack(kind: str, site: str, final: str, steps: list[dict], draw: dict | None = None, actions: list[dict] | None = None) -> dict:
        return {"kind": kind, "site": site, "final": final, "steps": steps, "draw": draw, "actions": actions or []}

    def steps_for(tilt: str, azimuth: str, fourth: str, new_site: str) -> list[dict]:
        return [
            {"step": "1", "action": "E-tilt / M-tilt", "proposal": tilt},
            {"step": "2", "action": "Azimuth", "proposal": azimuth},
            {"step": "3", "action": "4th sector, L900", "proposal": fourth},
            {"step": "4", "action": "New 3-sector site", "proposal": new_site},
        ]

    def tune_lines(chosen: list[dict]) -> tuple[str, str, str, str]:
        tilt_bits = []
        az_bits = []
        for item in chosen:
            rec = item["rec"]
            name = f"{rec.get('sector')} {rec.get('layer')}".strip()
            if abs(item["delta"]) >= 0.8:
                action = "uptilt so the beam reaches the spot" if item["delta"] < 0 else "downtilt so the beam lands on the spot"
                line = f"From {item['dist']:.1f} km, {item['off']:.0f}° off azimuth. {_tilt_line(rec, item['delta'], action)}"
                if item.get("overshoot"):
                    line += " Downtilt is limited to 2° so the existing coverage is not reduced."
                tilt_bits.append(line)
            else:
                tilt_bits.append(
                    f"{name} already has the tilt for {item['dist']:.1f} km (E_tilt {_deg_text(rec.get('etilt'))}°)."
                )
            if abs(item["turn"]) >= 8.0:
                new_az = (float(rec["azimuth"]) + item["turn"]) % 360.0
                direction = "clockwise" if item["turn"] > 0 else "counter-clockwise"
                pop_line = ""
                if item["old_pop"] is not None:
                    pop_line = (
                        f" Population on the current bearing is {_people(item['old_pop'])}, "
                        f"on the new bearing {_people(item['new_pop'])}, and on the spot {_people(spot_pop)}."
                    )
                az_bits.append(
                    f"{name}: azimuth {float(rec['azimuth']):.0f}° to {new_az:.0f}° "
                    f"({direction} {abs(item['turn']):.0f}°).{pop_line}"
                )
            else:
                az_bits.append(f"{name}: azimuth stays {float(rec['azimuth']):.0f}°. It already faces the spot.")
        site_names = []
        for item in chosen:
            if item["site"] not in site_names:
                site_names.append(item["site"])
        final_bits = []
        for item in chosen:
            rec = item["rec"]
            name = f"{rec.get('sector')} {rec.get('layer')}".strip()
            parts = [f"{name} at {item['dist']:.1f} km"]
            if abs(item["delta"]) >= 0.8:
                _old_e, _old_m, new_e, _new_m = _tilt_step(rec.get("etilt"), rec.get("mtilt"), item["delta"])
                parts.append(f"E_tilt {_deg_text(rec.get('etilt'))}° to {_deg_text(new_e)}°")
            if abs(item["turn"]) >= 8.0:
                new_az = (float(rec["azimuth"]) + item["turn"]) % 360.0
                parts.append(f"azimuth {float(rec['azimuth']):.0f}° to {new_az:.0f}°")
            else:
                parts.append(f"azimuth stays {float(rec['azimuth']):.0f}°")
            final_bits.append(", ".join(parts))
        return (
            " ".join(tilt_bits),
            " ".join(az_bits),
            ", ".join(site_names),
            "; ".join(final_bits) + ".",
        )

    def fourth_draw(site: dict) -> dict:
        return _proposal_draw(
            "fourth", site["lat"], site["lon"], [float(site["bearing"])],
            target_lat, target_lon, "4th L900",
        )

    def fourth_pack(site: dict, reason: str, tilt_why: str, azimuth_why: str, new_why: str) -> dict:
        text = _fourth_text(site, reason)
        return pack("fourth", site["name"], text, steps_for(
            _not_required(tilt_why),
            _not_required(azimuth_why),
            "Final. " + text,
            _not_required(new_why),
        ), fourth_draw(site), [_fourth_action(site)])

    def near_fourth_host() -> dict | None:
        """A site close enough that a 4th sector can replace a heavy downtilt."""
        band = [site for site in sites if 0.5 <= site["dist"] <= 1.0]
        if not band:
            band = [site for site in sites if site["dist"] <= 1.0]
        return band[0] if band else None

    if close_null is not None:
        text = _fourth_text(close_null)
        return pack("fourth", close_null["name"], text, steps_for(
            _not_required(
                f"The spot is in the null of {close_null['name']}, {close_null['dist']:.1f} km away, "
                "inside 1.5 km. A 4th sector is required here, so the existing tilts stay."
            ),
            _not_required(
                "Turning an existing azimuth would move that beam off the customers it already serves. "
                f"Population on the spot is {_people(spot_pop)}."
            ),
            "Final. " + text,
            _not_required("The 4th sector is inside 1.5 km, so a new site is not required."),
        ), fourth_draw(close_null), [_fourth_action(close_null)])

    few_sites = len(nearby) <= 2
    if nearby and few_sites and usable_null is not None and usable_null["dist"] <= FOURTH_RISK_KM:
        # One or two sites, and the spot is in their null inside 2 km.
        if not tunes or tunes[0]["off"] > 25.0:
            text = _fourth_text(usable_null)
            return pack("fourth", usable_null["name"], text, steps_for(
                _not_required(
                    f"Only {len(nearby)} site(s) are inside 2.5 km. "
                    "A tilt change on those few sectors would have to cover the null as well as their own roads."
                ),
                _not_required(
                    f"Only {len(nearby)} site(s) are nearby, so a 4th sector is used instead of turning their only beams. "
                    f"Population on the spot is {_people(spot_pop)}."
                ),
                "Final. " + text,
                _not_required(f"The 4th sector is {usable_null['dist']:.1f} km away, inside 2 km."),
            ), fourth_draw(usable_null), [_fourth_action(usable_null)])

    if tunes:
        chosen = []
        seen = set()
        for item in tunes:
            if item["site"] in seen:
                continue
            chosen.append(item)
            seen.add(item["site"])
            if len(chosen) == 2:
                break
        # Several sites do not force a tune. Use it when a sector already faces the spot.
        # Otherwise a 4th sector inside 2 km is the better fill for a null.
        if any(item.get("overshoot") for item in chosen):
            host = near_fourth_host()
            if host is not None:
                return fourth_pack(
                    host, "downtilt",
                    "The spot needs more than 2° of downtilt. That would shrink the existing coverage, "
                    f"so the tilts stay and a 4th sector within 1 km is used. Population on the spot is {_people(spot_pop)}.",
                    "An azimuth turn is not used. The 4th sector carries the new beam.",
                    f"The 4th sector is on {host['name']}, {host['dist']:.1f} km away, so a new site is not required.",
                )
        faces = chosen[0]["off"] <= 32.0
        if not faces and usable_null is not None and not few_sites:
            text = _fourth_text(usable_null)
            return pack("fourth", usable_null["name"], text, steps_for(
                _not_required(
                    f"{len(nearby)} sites are inside 2.5 km, but none faces the spot. "
                    "Tilting them would not put the beam on it."
                ),
                _not_required(
                    "An azimuth turn was not the better change. "
                    f"A 4th sector at {usable_null['name']} fills the null without moving the busier beams. "
                    f"Population on the spot is {_people(spot_pop)}."
                ),
                "Final. " + text,
                _not_required(f"The 4th sector is {usable_null['dist']:.1f} km away, inside 2 km."),
            ), fourth_draw(usable_null), [_fourth_action(usable_null)])
        tilt, azimuth, site, final = tune_lines(chosen)
        return pack("tune", site, final, steps_for(
            "Final. " + tilt,
            "Final. " + azimuth,
            _not_required(
                "The spot is inside 2.5 km of a sector that can be tuned, and it is not in a null inside 1.5 km. "
                "A 4th sector is not added on top of that beam."
            ),
            _not_required("Azimuth and tilt cover this spot, so a new site is not required."),
        ), None, [_tune_action(item) for item in chosen])

    saver = next((site for site in sites if site["dist"] <= FOURTH_SAVE_KM), None)
    if saver is not None:
        return fourth_pack(
            saver, "save",
            f"The nearest site is {saver['name']} at {saver['dist']:.1f} km. "
            "Azimuth and tilt of the existing sectors are not used for this spot. "
            "A 4th sector aimed at the bad spot is used instead of a new site.",
            f"Population on the spot is {_people(spot_pop)}. The new sector takes the bearing, so existing azimuths stay.",
            f"{saver['name']} is inside 3 km, so a new 3-sector site is not added.",
        )

    new_site, new_text, new_draw, new_actions = _new_site_text(target_lat, target_lon, red_lat, red_lon, sectors)
    if nearest is None:
        far = "No sector was found inside the L900 reach."
    elif nearest["dist"] > TUNE_KM:
        far = (
            f"The nearest sector is {nearest['name']} at {nearest['dist']:.1f} km. "
            "Azimuth and tilt are used only inside about 2.5 km."
        )
    else:
        far = (
            f"{len(nearby)} site(s) are inside 2.5 km, and a tilt or azimuth change "
            "would miss the spot or leave the busier side."
        )
    if usable_null is None:
        fourth = (
            "No existing site inside 3 km can take a sector aimed at this spot. "
            + (f"The nearest site is {nearest['name']} at {nearest['dist']:.1f} km." if nearest else "No site was found inside the L900 reach.")
        )
    else:
        fourth = f"A 4th sector was not the chosen step. {usable_null['name']} is {usable_null['dist']:.1f} km away."
    return pack("new", new_site, new_text, steps_for(
        _not_required(far),
        _not_required(far + f" Population on the spot is {_people(spot_pop)}."),
        _not_required(fourth),
        "Final. " + new_text,
    ), new_draw, new_actions)


def _served_slots(bins: list[dict]) -> list[dict]:
    """Serving cells on the red bins, with how far each one reaches."""
    served: dict[tuple, dict] = {}
    for item in bins:
        for end in item.get("ends") or ():
            cell = _resolve_serving_cell(end.get("eci"), end.get("earfcn"))
            if cell is None or cell.get("azimuth") is None:
                continue
            key = (cell.get("sector"), cell.get("layer"))
            slot = served.setdefault(key, {"by": {}, "n": 0, "lat": 0.0, "lon": 0.0})
            slot["n"] += 1
            slot["lat"] += float(end["lat"])
            slot["lon"] += float(end["lon"])
            name = cell.get("name") or ""
            bucket = slot["by"].setdefault(name, {"cell": cell, "n": 0})
            bucket["n"] += 1
    rows = []
    for slot in served.values():
        slat = slot["lat"] / slot["n"]
        slon = slot["lon"] / slot["n"]
        cell = max(slot["by"].values(), key=lambda item: item["n"])["cell"]
        dist = float(_equirect_km(cell["lat"], cell["lon"], slat, slon))
        off = abs(_signed_deg(cell["azimuth"], _bearing_deg(cell["lat"], cell["lon"], slat, slon)))
        rows.append({
            "cell": cell,
            "n": slot["n"],
            "dist": dist,
            "off": off,
            "lat": slat,
            "lon": slon,
        })
    rows.sort(key=lambda item: (-item["n"], -item["dist"]))
    return rows


def _downtilt_moves(cell: dict) -> bool:
    _old_e, _old_m, new_e, new_m = _tilt_step(cell.get("etilt"), cell.get("mtilt"), DOWNTILT_MAX_DEG)
    old_e, old_m = cell.get("etilt"), cell.get("mtilt")
    e_moves = old_e is not None and new_e is not None and abs(new_e - old_e) >= 0.4
    m_moves = old_m is not None and new_m is not None and abs(new_m - old_m) >= 0.4
    return e_moves or m_moves


def _overshoot_sentence(row: dict) -> str:
    cell = row["cell"]
    name = f"{cell.get('sector')} {cell.get('layer')}".strip()
    old_e, old_m, new_e, new_m = _tilt_step(cell.get("etilt"), cell.get("mtilt"), DOWNTILT_MAX_DEG)
    dist_m = row["dist"] * 1000.0
    return (
        f"{name} serves this road from {dist_m:,.0f} m. That is overshoot, so the beam is "
        f"pulled back: {_tilt_bit('E_tilt', old_e, new_e)}. {_tilt_bit('M_tilt', old_m, new_m)}. "
        f"Azimuth stays {float(cell['azimuth']):.0f}°."
    )


def _overshoot_item(row: dict) -> dict:
    cell = row["cell"]
    return {
        "rec": cell,
        "dist": row["dist"],
        "off": row["off"],
        "delta": DOWNTILT_MAX_DEG,
        "turn": 0.0,
        "old_pop": None,
        "new_pop": None,
        "overshoot": True,
        "site": cell.get("site") or cell.get("sector") or "",
    }


def _local_companion(sites: list[dict], spot_pop: float) -> list[dict]:
    """A nearest-sector change kept only when that beam itself misses the road.

    Downtilt of the nearest sector is not used once a far cell is the overshooter.
    An uptilt, or a large azimuth turn, is kept with the overshooter downtilt.
    """
    if not sites or sites[0]["dist"] > TUNE_KM:
        return []
    site = sites[0]
    for rec in site["recs"]:
        hit = _tune_candidate(rec, site["dist"], site["bearing"], spot_pop)
        if hit is None:
            continue
        if hit["delta"] > 0.4:
            hit["delta"] = 0.0
        if abs(hit["turn"]) < 15.0:
            hit["turn"] = 0.0
        if abs(hit["delta"]) < 0.8 and abs(hit["turn"]) < 15.0:
            continue
        hit["site"] = site["name"]
        hit["overshoot"] = False
        return [hit]
    return []


def _overshoot_rows(slots: list[dict], sites: list[dict]) -> list[dict]:
    """Far servers on every red piece inside the circle, not only the longest run.

    A short red sample is included with the longer one. The 50 m piece served
    by BOSBG03A is the case this covers. Later report builds must keep passing
    every red bin in the circle into this check.
    """
    nearest_km = sites[0]["dist"] if sites else 99.0
    if nearest_km > FOURTH_SAVE_KM:
        return []
    limit = max(OVERSHOOT_KM, nearest_km * 2.5)
    return [row for row in slots if row["dist"] >= limit and _downtilt_moves(row["cell"])]


def _overshoot_mark(row: dict) -> str:
    """The tilt change written on that cell's dashed line."""
    cell = row["cell"]
    old_e, _old_m, new_e, _new_m = _tilt_step(cell.get("etilt"), cell.get("mtilt"), DOWNTILT_MAX_DEG)
    return _tilt_bit("E_tilt", old_e, new_e)


def _on_planned_piece(row: dict, choice: dict) -> bool:
    """True when this far cell's samples sit on the piece the 4th sector or new site aims at."""
    if choice.get("kind") not in ("fourth", "new"):
        return False
    draw = choice.get("draw") or {}
    azimuths = draw.get("azimuths") or []
    if row.get("lat") is None or not azimuths or draw.get("lat") is None:
        return False
    bearing = _bearing_deg(float(draw["lat"]), float(draw["lon"]), float(row["lat"]), float(row["lon"]))
    off = abs(_signed_deg(float(azimuths[0]), bearing))
    apart = float(_equirect_km(
        float(draw["target_lat"]), float(draw["target_lon"]), float(row["lat"]), float(row["lon"]),
    ))
    # Same piece: the aim is within 40° or the sample is within 500 m of the aim point.
    # A separate red sample, such as the 50 m piece, stays outside this and is still downtilted.
    return off <= 40.0 or apart <= 0.50


def _idle_l2600(row: dict, mode: str) -> bool:
    """Idle keeps L2600 as the priority layer, so a far L2600 cell is not downtilted."""
    return mode == "Idle" and str(row["cell"].get("layer") or "") == "L2600"


def _spared_sentence(row: dict, mode: str) -> str:
    cell = row["cell"]
    name = f"{cell.get('sector')} {cell.get('layer')}".strip()
    dist_m = row["dist"] * 1000.0
    if _idle_l2600(row, mode):
        return (
            f"{name} serves this road from {dist_m:,.0f} m and is not downtilted. "
            f"L2600 is the idle priority layer, so a downtilt from that distance would shrink coverage elsewhere."
        )
    return (
        f"{name} serves this road from {dist_m:,.0f} m and is not downtilted. "
        f"The new sector covers that piece, so the far cell is left as it is."
    )


def _prefer_overshoot(
    choice: dict,
    slots: list[dict],
    sites: list[dict],
    spot_pop: float,
    mode: str = "",
) -> dict:
    """Downtilt a far serving cell, except where a new sector already covers that piece.

    Every red bin is still checked. A far cell is left in place when the 4th sector
    or the new site is aimed at the same piece, and when the cell is L2600 in idle.
    A far cell on a different red piece is still downtilted.
    """
    far = _overshoot_rows(slots, sites)
    if not far:
        choice = dict(choice)
        choice["overshoot_marks"] = {}
        choice["spared"] = []
        return choice
    kept = [row for row in far if not _on_planned_piece(row, choice) and not _idle_l2600(row, mode)]
    kept_ids = {id(row) for row in kept}
    spared = [row for row in far if id(row) not in kept_ids]
    spared_lines = [_spared_sentence(row, mode) for row in spared]
    marks = {
        (row["cell"].get("sector"), row["cell"].get("layer")): _overshoot_mark(row)
        for row in kept
    }
    if choice.get("kind") in ("fourth", "new"):
        choice = dict(choice)
        sentences = [_overshoot_sentence(row) for row in kept] + spared_lines
        if sentences:
            choice["final"] = f"{choice['final']} {' '.join(sentences)}"
        if kept:
            choice["actions"] = list(choice.get("actions") or []) + [
                _tune_action(_overshoot_item(row)) for row in kept
            ]
            if choice.get("steps"):
                choice["steps"][0]["proposal"] = " ".join(_overshoot_sentence(row) for row in kept)
        choice["overshoot_marks"] = marks
        choice["spared"] = spared
        return choice
    if not kept:
        # The nearest-site plan stays. The far L2600 cell is reported and not moved.
        choice = dict(choice)
        if spared_lines:
            choice["final"] = f"{choice['final']} {' '.join(spared_lines)}"
        choice["overshoot_marks"] = {}
        choice["spared"] = spared
        return choice
    sentences = [_overshoot_sentence(row) for row in kept]
    extra_actions = [_tune_action(_overshoot_item(row)) for row in kept]
    companions = _local_companion(sites, spot_pop)
    bits = list(sentences)
    actions = list(extra_actions)
    if companions:
        _tilt, _azimuth, _site, companion_final = _companion_final(companions)
        bits.append(companion_final)
        actions.extend(_tune_action(item) for item in companions)
    else:
        near = sites[0]["name"] if sites else "The nearest sector"
        nearest_km = sites[0]["dist"] if sites else 99.0
        bits.append(f"{near} at {nearest_km:.1f} km is not moved. The bad samples come from the overshooting cell.")
    bits.extend(spared_lines)
    final = " ".join(bits)
    names = []
    for item in [_overshoot_item(row) for row in kept] + companions:
        if item["site"] not in names:
            names.append(item["site"])
    steps = [
        {"step": "1", "action": "E-tilt / M-tilt", "proposal": "Final. " + final},
        {
            "step": "2",
            "action": "Azimuth",
            "proposal": _not_required(
                "The overshooting cell keeps its azimuth. Downtilt is what pulls the beam back."
            ),
        },
        {
            "step": "3",
            "action": "4th sector, L900",
            "proposal": _not_required(
                "A 4th sector is not added on the nearest site. The far cell is downtilted first."
            ),
        },
        {
            "step": "4",
            "action": "New 3-sector site",
            "proposal": _not_required("The overshoot is restricted, so a new site is not added for this reason."),
        },
    ]
    return {
        "kind": "tune",
        "site": ", ".join(names),
        "final": final,
        "steps": steps,
        "draw": choice.get("draw"),
        "actions": actions,
        "azimuth_overrides": dict(choice.get("azimuth_overrides") or {}),
        "overshoot_marks": marks,
        "spared": spared,
    }


def _host_site(choice: dict, sites: list[dict]) -> dict | None:
    draw = choice.get("draw") or {}
    if draw.get("lat") is None:
        return None
    for site in sites:
        if abs(site["lat"] - float(draw["lat"])) < 0.0008 and abs(site["lon"] - float(draw["lon"])) < 0.0008:
            return site
    return None


def _note_azimuth_step(choice: dict, sentence: str) -> None:
    for step in choice.get("steps") or []:
        if step.get("action") == "Azimuth":
            text = step.get("proposal") or ""
            if "not moved" in text or text.startswith("Not required") or "not the chosen" in text:
                step["proposal"] = sentence
            elif sentence not in text:
                step["proposal"] = f"{text} {sentence}"
            return


def _separate_fourth_sector(choice: dict, sites: list[dict], spot_pop: float) -> dict:
    """Turn the existing sector that sits closest to a new 4th sector, until the gap is 30–60°.

    Only that one sector moves. The new sector stays aimed at the bad spot. The
    turn goes away from the new azimuth, and it stops short of the next sector.
    """
    if choice.get("kind") != "fourth":
        return choice
    host = _host_site(choice, sites)
    draw = choice.get("draw") or {}
    azimuths = draw.get("azimuths") or []
    if host is None or not azimuths:
        return choice
    new_az = float(azimuths[0])
    recs = [rec for rec in host["recs"] if rec.get("azimuth") is not None and rec.get("sector")]
    if not recs:
        return choice
    closest = min(recs, key=lambda rec: abs(_signed_deg(new_az, float(rec["azimuth"]))))
    current = abs(_signed_deg(new_az, float(closest["azimuth"])))
    if current >= 30.0:
        return choice
    signed = _signed_deg(new_az, float(closest["azimuth"]))
    direction = 1.0 if signed >= 0 else -1.0
    if abs(signed) < 0.5:
        left = (float(closest["azimuth"]) - 45.0) % 360.0
        right = (float(closest["azimuth"]) + 45.0) % 360.0
        pop_left = _wedge_population(closest["lat"], closest["lon"], left)
        pop_right = _wedge_population(closest["lat"], closest["lon"], right)
        direction = -1.0 if pop_left >= pop_right else 1.0
    room = 80.0
    for rec in recs:
        if rec is closest:
            continue
        delta = _signed_deg(float(closest["azimuth"]), float(rec["azimuth"]))
        if direction > 0 and delta > 0:
            room = min(room, delta - 30.0)
        elif direction < 0 and delta < 0:
            room = min(room, abs(delta) - 30.0)
    room = max(0.0, room)
    needed = 45.0 - current
    minimum = 30.0 - current
    turn = min(needed, room)
    if turn + 0.5 < minimum:
        if room < 8.0:
            return choice
        turn = room
    old_az = float(closest["azimuth"])
    new_bearing = (old_az + direction * turn) % 360.0
    old_pop = _wedge_population(closest["lat"], closest["lon"], old_az)
    new_pop = _wedge_population(closest["lat"], closest["lon"], new_bearing)
    if new_pop + 0.5 * spot_pop < old_pop * 0.60 and turn > minimum + 0.5 and room >= minimum:
        turn = minimum
        new_bearing = (old_az + direction * turn) % 360.0
    if turn < 8.0:
        return choice
    item = {
        "rec": closest,
        "dist": host["dist"],
        "off": current,
        "delta": 0.0,
        "turn": direction * turn,
        "old_pop": old_pop,
        "new_pop": new_pop,
        "overshoot": False,
        "site": host["name"],
    }
    action = _tune_action(item)
    action["action"] = "Azimuth, open a gap"
    gap = current + turn
    sentence = (
        f"{closest.get('sector')} turns from {old_az:.0f}° to {new_bearing:.0f}° "
        f"so the gap to the new sector is {gap:.0f}°. It was {current:.0f}° off that new azimuth."
    )
    choice = dict(choice)
    text = choice["final"]
    if "Existing sectors are not moved." in text:
        text = text.replace("Existing sectors are not moved.", sentence)
    else:
        text = f"{text} {sentence}"
    choice["final"] = text
    choice["actions"] = list(choice.get("actions") or []) + [action]
    overrides = dict(choice.get("azimuth_overrides") or {})
    overrides[str(closest.get("sector")).strip()] = new_bearing
    choice["azimuth_overrides"] = overrides
    _note_azimuth_step(choice, sentence)
    return choice


def _turn_neighbor_to_long_end(
    choice: dict,
    sites: list[dict],
    blanks: list[dict],
    spot_pop: float,
    red_lat: float,
    red_lon: float,
    blank_m: float = 0.0,
) -> dict:
    """Turn the nearest other sector toward the far end of a long bad spot.

    Used when the blank is at least 1.5 km and that end sits about 2 km from the
    site that hosts the new sector. The new sector stays. This adds an azimuth
    on the neighbour, such as GBGBD15A toward the long blank past GBGBD24.
    """
    if choice.get("kind") not in ("fourth", "new") or blank_m < 1500:
        return choice
    host = _host_site(choice, sites)
    if host is None:
        return choice
    points = [(red_lat, red_lon)]
    for gap in blanks:
        points.append((float(gap["lat1"]), float(gap["lon1"])))
        points.append((float(gap["lat2"]), float(gap["lon2"])))
    far = max(points, key=lambda point: float(_equirect_km(host["lat"], host["lon"], point[0], point[1])))
    far_km = float(_equirect_km(host["lat"], host["lon"], far[0], far[1]))
    if far_km < 2.0:
        return choice
    others = []
    for site in sites:
        if abs(site["lat"] - host["lat"]) < 0.0008 and abs(site["lon"] - host["lon"]) < 0.0008:
            continue
        dist = float(_equirect_km(site["lat"], site["lon"], far[0], far[1]))
        if dist <= 3.2:
            others.append((dist, site))
    if not others:
        return choice
    dist, site = min(others, key=lambda item: item[0])
    bearing = _bearing_deg(site["lat"], site["lon"], far[0], far[1])
    best = None
    best_off = 999.0
    for rec in site["recs"]:
        if rec.get("azimuth") is None or not rec.get("sector"):
            continue
        off = abs(_signed_deg(float(rec["azimuth"]), bearing))
        if off < best_off:
            best_off = off
            best = rec
    if best is None or best_off < 12.0:
        return choice
    name = str(best.get("sector")).strip()
    if name in (choice.get("azimuth_overrides") or {}):
        return choice
    signed = _signed_deg(float(best["azimuth"]), bearing)
    # Point the sector at the far end. Cap the turn so a large swing is not proposed.
    turn = math.copysign(min(best_off, 25.0), signed)
    if abs(turn) < 10.0:
        return choice
    new_az = (float(best["azimuth"]) + turn) % 360.0
    old_pop = _wedge_population(best["lat"], best["lon"], best["azimuth"])
    new_pop = _wedge_population(best["lat"], best["lon"], new_az)
    if new_pop + 0.5 * spot_pop < old_pop * 0.90:
        return choice
    item = {
        "rec": best,
        "dist": dist,
        "off": best_off,
        "delta": 0.0,
        "turn": turn,
        "old_pop": old_pop,
        "new_pop": new_pop,
        "overshoot": False,
        "site": site["name"],
    }
    action = _tune_action(item)
    action["action"] = "Azimuth toward the far end"
    sentence = (
        f"{name} at {dist:.1f} km from the far end turns from {float(best['azimuth']):.0f}° to {new_az:.0f}° "
        f"toward that part of the bad spot."
    )
    choice = dict(choice)
    text = choice["final"]
    if "Existing sectors are not moved." in text:
        text = text.replace("Existing sectors are not moved.", sentence)
    else:
        text = f"{text} {sentence}"
    choice["final"] = text
    choice["actions"] = list(choice.get("actions") or []) + [action]
    overrides = dict(choice.get("azimuth_overrides") or {})
    overrides[name] = new_az
    choice["azimuth_overrides"] = overrides
    _note_azimuth_step(choice, sentence)
    return choice


def _whole_degrees(text: str) -> str:
    """Nearest whole degree. Half a degree rounds up. A move that lands on the same degree stays."""
    if not text:
        return text

    def _repl(match: re.Match) -> str:
        return f"{math.floor(float(match.group(1)) + 0.5):.0f}°"

    text = re.sub(r"(\d+\.\d+)\s*°", _repl, text)
    text = re.sub(
        r"(E_tilt|M_tilt|azimuth) (\d+)° to \2°",
        lambda match: f"{match.group(1)} stays {match.group(2)}°",
        text,
    )
    return text


_ACTION_START = re.compile(r"^(?:New L900|New 3-sector|[A-Z]{2,}[A-Z0-9]*\d[A-Z0-9]*)\b")


def _number_actions(text: str) -> str:
    """Number the cell actions when a spot changes more than one cell.

    Tilt and azimuth sentences that belong to the same cell stay in one point.
    A new point starts at the next cell, new sector, or new site. Each point is
    its own line.
    """
    parts = [part.strip() for part in re.split(r"(?<=\.)\s+", text.strip()) if part.strip()]
    points: list[str] = []
    tail: list[str] = []
    for part in parts:
        sentence = part if part.endswith(".") else part + "."
        if sentence.startswith("This covers") or sentence.startswith("A foot inside"):
            tail.append(sentence)
            continue
        if points and not _ACTION_START.match(sentence):
            points[-1] = points[-1].rstrip(".") + ". " + sentence
        else:
            points.append(sentence)
    if len(points) < 2:
        return text.strip()
    lines = [f"{number}. {sentence}" for number, sentence in enumerate(points, start=1)]
    lines.extend(tail)
    return "\n".join(lines)


def _round_proposed(row: dict) -> dict:
    """Proposed azimuth and tilt become whole degrees. A rounded no-change keeps the dump value.

    The whole degree is the one printed in the suggestion, so the parameter table
    matches that sentence.
    """
    row = dict(row)
    for src, dst in (("azimuth", "p_azimuth"), ("et", "p_et"), ("mt", "p_mt")):
        if row.get(dst) is None:
            continue
        # Same whole degree the suggestion sentence prints, so the table does not
        # disagree with the text on an exact half (82.5° is written 82°).
        proposed = float(f"{float(row[dst]):.0f}")
        if row.get(src) is None:
            row[dst] = proposed
            continue
        current = float(row[src])
        if float(math.floor(current + 0.5)) == proposed:
            row[dst] = current
        else:
            row[dst] = proposed
    return row


def _companion_final(chosen: list[dict]) -> tuple[str, str, str, str]:
    """Short final line for a nearest sector that still needs an uptilt or a large turn."""
    bits = []
    for item in chosen:
        rec = item["rec"]
        name = f"{rec.get('sector')} {rec.get('layer')}".strip()
        parts = [f"{name} at {item['dist']:.1f} km"]
        if abs(item["delta"]) >= 0.8:
            _old_e, _old_m, new_e, _new_m = _tilt_step(rec.get("etilt"), rec.get("mtilt"), item["delta"])
            parts.append(f"E_tilt {_deg_text(rec.get('etilt'))}° to {_deg_text(new_e)}°")
        if abs(item["turn"]) >= 15.0:
            new_az = (float(rec["azimuth"]) + item["turn"]) % 360.0
            parts.append(f"azimuth {float(rec['azimuth']):.0f}° to {new_az:.0f}°")
        bits.append(", ".join(parts))
    return "", "", "", "; ".join(bits) + "."


def _facing_for_ep_check(sites: list[dict], seen: set[str]) -> dict | None:
    """Nearest site already faces the red road, but that sector is absent from this log.

    The other sectors of the same site have to be present, so a site the drive
    never reached is left alone. The check is the physical azimuth and whether
    the sector is radiating.
    """
    if not sites or sites[0]["dist"] > TUNE_KM:
        return None
    site = sites[0]
    facing = None
    for rec in site["recs"]:
        if rec.get("azimuth") is None:
            continue
        off = abs(_signed_deg(float(rec["azimuth"]), site["bearing"]))
        reach = _layer_reach_km(rec.get("layer") or "")
        if off > FACE_ON_ROAD_DEG or site["dist"] > reach:
            continue
        if facing is None or off < facing["off"]:
            facing = {"site": site, "rec": rec, "off": off, "dist": site["dist"]}
    if facing is None:
        return None
    sector = str(facing["rec"].get("sector") or "")
    if not sector or sector in seen:
        return None
    siblings = []
    for rec in site["recs"]:
        name = str(rec.get("sector") or "")
        if name and name != sector and name in seen and name not in siblings:
            siblings.append(name)
    if not siblings:
        return None
    facing["siblings"] = siblings
    return facing


def _ep_check_choice(hit: dict, spot_pop: float, mode: str) -> dict:
    """Replace a radio retune when the facing sector is missing from the log."""
    rec = hit["rec"]
    site = hit["site"]
    sector = str(rec.get("sector") or "")
    cell = str(rec.get("name") or sector)
    log = "active" if "active" in mode.lower() else "idle"
    siblings = ", ".join(hit["siblings"])
    text = (
        f"On site EP Check. {sector} is {hit['dist']:.1f} km from the red road. "
        f"The dump azimuth is {float(rec['azimuth']):.0f}°, {hit['off']:.0f}° off the road, "
        f"so this sector should be the dominant server. The {log} log does not decode {sector}, "
        f"while {siblings} on the same site are on air. Check on site that the antenna azimuth "
        f"is {float(rec['azimuth']):.0f}° and that the sector is radiating."
    )
    steps = [
        {"step": "1", "action": "E-tilt / M-tilt", "proposal": _not_required(
            "Tilt is not changed until the site check confirms the sector is on air and the azimuth is the dump value."
        )},
        {"step": "2", "action": "Azimuth", "proposal": "Final. " + text},
        {"step": "3", "action": "4th sector, L900", "proposal": _not_required(
            "A 4th sector is not added while the facing sector may be off air or on a different azimuth."
        )},
        {"step": "4", "action": "New 3-sector site", "proposal": _not_required(
            f"Population on the spot is {_people(spot_pop)}. A new site is not planned until the existing sector is checked."
        )},
    ]
    return {
        "kind": "ep",
        "site": site["name"],
        "final": text,
        "steps": steps,
        "draw": None,
        "actions": [{
            "site": site["name"],
            "sector": sector,
            "cell": cell,
            "height": _ep_num(rec.get("height")),
            "azimuth": _ep_num(rec.get("azimuth")),
            "et": _ep_num(rec.get("etilt")),
            "mt": _ep_num(rec.get("mtilt")),
            "p_height": _ep_num(rec.get("height")),
            "p_azimuth": None,
            "p_et": None,
            "p_mt": None,
            "action": "On site EP Check",
        }],
        "overshoot_marks": {},
        "azimuth_overrides": {},
        "spared": [],
    }


def _proposed_beams(choice: dict, sites: list[dict]) -> list[tuple[float, float, float]]:
    """Origin and azimuth of every beam this solution already draws."""
    beams: list[tuple[float, float, float]] = []
    draw = choice.get("draw") or {}
    if draw.get("lat") is not None and draw.get("azimuths"):
        for azimuth in draw["azimuths"]:
            beams.append((float(draw["lat"]), float(draw["lon"]), float(azimuth)))
    by_name = {site["name"]: site for site in sites}
    for action in choice.get("actions") or []:
        azimuth = action.get("p_azimuth")
        site = by_name.get(str(action.get("site") or ""))
        if site is None or azimuth is None:
            continue
        beams.append((float(site["lat"]), float(site["lon"]), float(azimuth)))
    for extra in draw.get("more") or []:
        beams.append((float(extra["lat"]), float(extra["lon"]), float(extra["azimuth"])))
    return beams


def _offbeam_pockets(samples: pd.DataFrame, column: str, spot: pd.Series, beams: list[tuple[float, float, float]]) -> list[dict]:
    """Red samples and a blank inside the circle that sit well off every planned beam.

    Active Bad Spot 11 is the pattern: the 4th sector covers one red bin and the
    blank beside it, while another red bin and a long blank sit about 45° away
    and have no solution.
    """
    if not beams:
        return []
    bins = _red_bins_inside(samples, column, spot)
    runs = _red_runs(bins) if bins else []
    points = []
    for run in runs:
        points.append({
            "lat": sum(item["lat"] for item in run) / len(run),
            "lon": sum(item["lon"] for item in run) / len(run),
            "red": _run_length_m(run),
            "blank": 0.0,
        })
    for gap in _blank_segments(samples, spot):
        points.append({
            "lat": (float(gap["lat1"]) + float(gap["lat2"])) / 2.0,
            "lon": (float(gap["lon1"]) + float(gap["lon2"])) / 2.0,
            "red": 0.0,
            "blank": float(gap["length_m"]),
        })
    uncovered = []
    for point in points:
        offset = min(
            abs(_signed_deg(azimuth, _bearing_deg(lat, lon, point["lat"], point["lon"])))
            for lat, lon, azimuth in beams
        )
        if offset > 35.0:
            uncovered.append({**point, "off": offset})
    groups: list[dict] = []
    for point in uncovered:
        placed = False
        for group in groups:
            if float(_equirect_km(group["lat"], group["lon"], point["lat"], point["lon"])) <= 1.0:
                count = group["n"] + 1
                group["lat"] = (group["lat"] * group["n"] + point["lat"]) / count
                group["lon"] = (group["lon"] * group["n"] + point["lon"]) / count
                group["n"] = count
                group["red"] += point["red"]
                group["blank"] += point["blank"]
                group["off"] = max(group["off"], point["off"])
                placed = True
                break
        if not placed:
            groups.append({**point, "n": 1})
    kept = [group for group in groups if group["red"] >= 50.0 and group["blank"] >= 150.0]
    kept.sort(key=lambda group: group["red"] + group["blank"], reverse=True)
    return kept[:2]


def _cover_offbeam_pocket(
    choice: dict,
    samples: pd.DataFrame,
    column: str,
    spot: pd.Series,
    sites: list[dict],
    spot_pop: float,
) -> dict:
    """Add a beam for a red-and-blank pocket the first solution does not cover."""
    if choice.get("kind") == "ep":
        return choice
    beams = _proposed_beams(choice, sites)
    pockets = _offbeam_pockets(samples, column, spot, beams)
    if not pockets:
        return choice
    choice = dict(choice)
    choice["actions"] = list(choice.get("actions") or [])
    choice["azimuth_overrides"] = dict(choice.get("azimuth_overrides") or {})
    if choice.get("draw"):
        choice["draw"] = dict(choice["draw"])
        choice["draw"]["azimuths"] = list(choice["draw"].get("azimuths") or [])
        choice["draw"]["arm_targets"] = dict(choice["draw"].get("arm_targets") or {})
        choice["draw"]["more"] = list(choice["draw"].get("more") or [])
    sentences = []
    for pocket in pockets:
        beams = _proposed_beams(choice, sites)
        still = min(
            abs(_signed_deg(azimuth, _bearing_deg(lat, lon, pocket["lat"], pocket["lon"])))
            for lat, lon, azimuth in beams
        ) if beams else 999.0
        if still <= 35.0:
            continue
        fix = _pocket_beam(choice, pocket, spot_pop)
        if fix is None:
            continue
        sentences.append(fix)
    if not sentences:
        return choice
    choice["final"] = (choice.get("final") or "").rstrip() + " " + " ".join(sentences)
    return choice


def _pocket_beam(choice: dict, pocket: dict, spot_pop: float) -> str | None:
    """One extra sector, or one turn, aimed at the uncovered pocket. None when no host is inside 3 km."""
    del spot_pop
    near = _sites_around(_sectors_near(pocket["lat"], pocket["lon"], 8.5), pocket["lat"], pocket["lon"])
    near = [site for site in near if site["dist"] <= FOURTH_SAVE_KM]
    if not near:
        return None
    taken: dict[str, list[float]] = {}
    draw = choice.get("draw") or {}
    if choice.get("kind") in ("fourth", "new") and draw.get("azimuths"):
        taken.setdefault(str(choice.get("site") or ""), []).extend(float(az) for az in draw["azimuths"])
    for action in choice.get("actions") or []:
        if action.get("p_azimuth") is None:
            continue
        taken.setdefault(str(action.get("site") or ""), []).append(float(action["p_azimuth"]))

    facing = None
    for site in near:
        for rec in site["recs"]:
            if rec.get("azimuth") is None or not rec.get("sector"):
                continue
            off = abs(_signed_deg(float(rec["azimuth"]), site["bearing"]))
            if off > 12.0:
                continue
            score = (site["dist"], 0 if rec.get("layer") == "L900" else 1, off)
            if facing is None or score < facing[0]:
                facing = (score, site, rec)
    if facing is not None:
        _score, site, rec = facing
        return (
            f"{rec.get('sector')} {rec.get('layer')} already faces this side of the circle, "
            f"azimuth {float(rec['azimuth']):.0f}°, {site['dist']:.1f} km away. "
            f"It covers the red samples and the {pocket['blank']:,.0f} m blank there."
        )

    turn = None
    for site in near:
        if site["dist"] > TUNE_KM:
            continue
        for rec in site["recs"]:
            if rec.get("azimuth") is None or rec.get("layer") not in ("L900", "L1800"):
                continue
            sector = str(rec.get("sector") or "").strip()
            if not sector or sector in choice.get("azimuth_overrides", {}):
                continue
            off = abs(_signed_deg(float(rec["azimuth"]), site["bearing"]))
            if not 12.0 < off <= 32.0:
                continue
            score = (0 if rec.get("layer") == "L900" else 1, site["dist"], off)
            if turn is None or score < turn[0]:
                turn = (score, site, rec, off)
    if turn is not None:
        _score, site, rec, off = turn
        signed = _signed_deg(float(rec["azimuth"]), site["bearing"])
        delta = math.copysign(min(off, 25.0), signed)
        new_az = (float(rec["azimuth"]) + delta) % 360.0
        item = {
            "rec": rec, "dist": site["dist"], "off": off, "delta": 0.0, "turn": delta,
            "old_pop": None, "new_pop": None, "overshoot": False, "site": site["name"],
        }
        action = _tune_action(item)
        action["action"] = "Azimuth toward the other side"
        choice["actions"].append(action)
        choice["azimuth_overrides"][str(rec.get("sector")).strip()] = new_az
        return (
            f"{rec.get('sector')} turns from {float(rec['azimuth']):.0f}° to {new_az:.0f}° "
            f"to cover the red samples and the {pocket['blank']:,.0f} m blank "
            f"on the other side of the circle."
        )

    for site in near:
        bearing = float(site["bearing"])
        existing = [float(rec["azimuth"]) for rec in site["recs"] if rec.get("azimuth") is not None]
        blocked = existing + taken.get(site["name"], [])
        if any(_azimuth_gap(bearing, azimuth) < 40.0 for azimuth in blocked):
            continue
        tilt = int(round(min(10.0, max(0.0, _ground_tilt_deg(site.get("height"), site["dist"])))))
        action = _fourth_action(site)
        action["p_azimuth"] = _ep_num(bearing)
        action["p_et"] = float(tilt)
        choice["actions"].append(action)
        if (
            choice.get("draw")
            and choice.get("kind") == "fourth"
            and abs(float(choice["draw"]["lat"]) - site["lat"]) < 1e-4
            and abs(float(choice["draw"]["lon"]) - site["lon"]) < 1e-4
        ):
            choice["draw"]["azimuths"].append(bearing)
            choice["draw"]["arm_targets"][round(bearing, 1)] = (pocket["lat"], pocket["lon"])
        elif choice.get("draw") is None:
            choice["draw"] = _proposal_draw(
                "fourth", site["lat"], site["lon"], [bearing], pocket["lat"], pocket["lon"], "4th L900",
            )
            choice["draw"]["arm_targets"] = {round(bearing, 1): (pocket["lat"], pocket["lon"])}
            choice["draw"]["more"] = []
        else:
            choice["draw"].setdefault("more", []).append({
                "lat": site["lat"], "lon": site["lon"], "azimuth": bearing,
                "target_lat": pocket["lat"], "target_lon": pocket["lon"],
            })
        return (
            f"New L900 sector on {site['name']}, azimuth {bearing:.0f}°, E_tilt {tilt}°, M_tilt 0°, "
            f"covers the red samples and the {pocket['blank']:,.0f} m blank "
            f"{pocket['off']:.0f}° off the first beam."
        )
    return None


def bad_spot_advice(samples: pd.DataFrame, column: str, spot: pd.Series, mode: str = "") -> dict:
    """Finding and one final solution for a bad spot, including a blank inside the circle.

    Inside about 2.5 km the final step can retune azimuth and tilt. Downtilt is
    limited to 2°. A larger downtilt becomes a 4th L900 sector on a site within
    1 km. A 4th sector also replaces a new site when a host inside 3 km can aim
    at the bad spot. A new 3-sector site is kept only past that. The side table
    shows all four steps. The final one is marked.
    """
    blanks = _blank_segments(samples, spot)
    blank_m = sum(item["length_m"] for item in blanks)
    bins = _red_bins_inside(samples, column, spot)
    runs = _red_runs(bins) if bins else []
    main: list[dict] = []
    if runs:
        # The sector is aimed at the poor road, not at a blank that can sit on the other side of a site.
        main = max(runs, key=_run_length_m)
        red_lat = sum(item["lat"] for item in main) / len(main)
        red_lon = sum(item["lon"] for item in main) / len(main)
    elif bins:
        red_lat = sum(item["lat"] for item in bins) / len(bins)
        red_lon = sum(item["lon"] for item in bins) / len(bins)
    else:
        red_lat = float(spot["lat"])
        red_lon = float(spot["lon"])
    target_lat, target_lon = red_lat, red_lon
    target_name = "the red road"
    spot_pop = population_per_km2(target_lat, target_lon)
    sectors = _sectors_near(target_lat, target_lon, 8.5)
    finding_lines = []
    if blank_m >= 80:
        finding_lines.append(
            f"Blank, no plot, {blank_m:,.0f} m inside this circle is counted as part of this bad spot."
        )
    else:
        finding_lines.append("No driven blank inside this circle. The poor road is the red plot.")
    def serving_bits(chosen: list[dict]) -> list[str]:
        served: dict[tuple, dict] = {}
        for item in chosen:
            for end in item.get("ends") or ():
                cell = _resolve_serving_cell(end.get("eci"), end.get("earfcn"))
                if cell is None or cell.get("azimuth") is None:
                    continue
                key = (cell.get("sector"), cell.get("layer"))
                slot = served.setdefault(key, {"cell": cell, "n": 0, "lat": 0.0, "lon": 0.0})
                slot["n"] += 1
                slot["lat"] += float(end["lat"])
                slot["lon"] += float(end["lon"])
        bits = []
        for slot in sorted(served.values(), key=lambda item: -item["n"])[:4]:
            cell = slot["cell"]
            slat = slot["lat"] / slot["n"]
            slon = slot["lon"] / slot["n"]
            dist = float(_equirect_km(cell["lat"], cell["lon"], slat, slon))
            off = abs(_signed_deg(cell["azimuth"], _bearing_deg(cell["lat"], cell["lon"], slat, slon)))
            who = f"{cell.get('sector')} {cell.get('layer')}".strip()
            noun = "bin" if slot["n"] == 1 else "bins"
            bits.append(f"{who} ({slot['n']} {noun}, {dist:.1f} km, {off:.0f}° off)")
        return bits

    road_bits = serving_bits(main or bins)
    if road_bits:
        finding_lines.append("Serving cells on this red road: " + "; ".join(road_bits) + ".")
    other = [item for item in bins if main and item not in main]
    other_bits = serving_bits(other)
    if other_bits:
        finding_lines.append("Other red samples in this circle: " + "; ".join(other_bits) + ".")
    finding_lines.append(
        f"WorldPop 2020 on {target_name} is {_people(spot_pop)}. "
        f"L900 reach is {L900_REACH_KM:.0f} km, L1800 {_layer_reach_km('L1800'):.1f} km, "
        f"L2100 {_layer_reach_km('L2100'):.1f} km, L2600 {_layer_reach_km('L2600'):.1f} km."
    )

    sites = _sites_around(sectors, target_lat, target_lon)
    if sites:
        finding_lines.append(
            f"Nearest sector is {sites[0]['name']} at {sites[0]['dist']:.1f} km. "
            f"{sum(site['dist'] <= TUNE_KM for site in sites)} site(s) are inside 2.5 km."
        )
    choice = _choose_coverage_fix(
        sites, sectors, spot_pop, target_lat, target_lon, red_lat, red_lon,
    )
    choice = _separate_fourth_sector(choice, sites, spot_pop)
    choice = _turn_neighbor_to_long_end(
        choice, sites, blanks, spot_pop, red_lat, red_lon, blank_m,
    )
    # All red bins in the circle. The longest run alone would drop a far cell
    # on a shorter piece, such as BOSBG03A on the 50 m sample.
    served = _served_slots(bins) if bins else []
    far = _overshoot_rows(served, sites)
    plan_mode = "Active" if "active" in str(mode).lower() else ("Idle" if "idle" in str(mode).lower() else "")
    choice = _prefer_overshoot(choice, served, sites, spot_pop, plan_mode)
    seen = _SEEN_SECTORS.get(plan_mode)
    if seen is not None:
        silent = _facing_for_ep_check(sites, seen)
        if silent is not None:
            choice = _ep_check_choice(silent, spot_pop, plan_mode)
    if choice.get("kind") != "ep":
        choice = _cover_offbeam_pocket(choice, samples, column, spot, sites, spot_pop)
    serving_names: list[str] = []
    for row in served:
        name = str(row["cell"].get("name") or "").strip()
        if name and name not in serving_names:
            serving_names.append(name)
    spared = choice.get("spared") or []
    if far:
        names = "; ".join(
            f"{row['cell'].get('sector')} {row['cell'].get('layer')} at {row['dist'] * 1000:,.0f} m"
            for row in far
        )
        line = f"Far server inside this circle: {names}."
        if spared:
            left = ", ".join(
                f"{row['cell'].get('sector')} {row['cell'].get('layer')}".strip() for row in spared
            )
            line += f" Left in place: {left}."
        finding_lines.append(line)
    final = choice["final"]
    if blank_m >= 80:
        final = f"{final} This covers the red road and the {blank_m:,.0f} m blank."
    final = _whole_degrees(_number_actions(final))
    sections = _spot_sections(
        served, sites, spot, blank_m, spot_pop, target_lat, target_lon, choice, final,
    )
    steps = []
    for step in choice.get("steps") or []:
        step = dict(step)
        proposal = _whole_degrees(step.get("proposal") or "")
        if proposal.startswith("Final. "):
            proposal = "Final. " + final
        step["proposal"] = proposal
        steps.append(step)
    return {
        "finding": " ".join(finding_lines),
        "blank_m": blank_m,
        "steps": steps,
        "kind": choice["kind"],
        "site": choice["site"],
        "serving": ", ".join(serving_names),
        "population": _people(spot_pop),
        "final": final,
        "draw": choice.get("draw"),
        "actions": [_round_proposed(row) for row in (choice.get("actions") or [])],
        "overshoot_marks": choice.get("overshoot_marks") or {},
        "azimuth_overrides": choice.get("azimuth_overrides") or {},
        "sections": sections,
        "outline": {
            "lat": float(spot.get("outline_lat", spot["lat"])),
            "lon": float(spot.get("outline_lon", spot["lon"])),
            "width_km": float(spot.get("outline_width_km") or 0),
            "height_km": float(spot.get("outline_height_km") or 0),
        },
    }


def _server_comment(layer: str, dist_km: float, off_deg: float) -> str:
    """Why this serving cell is weak on the red samples."""
    reach = _layer_reach_km(layer or "")
    if off_deg >= 90:
        return "The sample is behind the sector, so only a back lobe reaches it."
    if off_deg >= 35:
        return "The sample is outside the main beam, so this is a side-lobe server."
    if dist_km > reach:
        return f"It is past the {reach:.1f} km {layer} reach, so this is overshoot."
    if off_deg <= 12:
        return "It points near the road, but it is not a stable dominant server."
    return "It is off the beam centre, so the serving signal is weak."


def _spot_sections(
    served: list[dict],
    sites: list[dict],
    spot: pd.Series,
    blank_m: float,
    spot_pop: float,
    lat: float,
    lon: float,
    choice: dict,
    final: str,
) -> dict[str, str]:
    """The four description blocks used on the Excel sheet, the PPT, and the web.

    The engineering decision stays the one already chosen. The text only explains it:
    one line per serving cell, the geography, the root cause, and the suggestion.
    """
    serving_lines = []
    for index, row in enumerate(served[:8], start=1):
        cell = row["cell"]
        name = str(cell.get("name") or cell.get("sector") or "cell")
        layer = str(cell.get("layer") or "")
        height = cell.get("height")
        height_bit = f", height {float(height):.0f} m" if height else ""
        look = _ground_tilt_deg(height, row["dist"])
        serving_lines.append(
            f"{index}. {name} {layer}, {row['dist']:.2f} km, azimuth {float(cell['azimuth']):.0f}°, "
            f"{row['off']:.0f}° off{height_bit}, look-down {look:.2f}°. "
            f"{_server_comment(layer, row['dist'], row['off'])}"
        )
    extra = len(served) - 8
    if extra > 0:
        serving_lines.append(f"{len(serving_lines) + 1}. {extra} more serving cells are on the red samples.")
    if not serving_lines:
        serving_lines.append("1. No serving cell was decoded on the red samples.")

    poor_m = float(spot.get("length_m") or 0)
    geo = [
        f"1. Poor run {poor_m:.0f} m."
        + (f" Blank inside the circle {blank_m:,.0f} m." if blank_m >= 80 else " No driven blank inside the circle."),
        f"2. Centre {lat:.5f}, {lon:.5f}. WorldPop 2020 {_people(spot_pop)}.",
    ]
    nearest_bits = []
    for site in sites[:4]:
        height = site.get("height")
        height_bit = f", antenna {float(height):.0f} m" if height else ""
        look = _ground_tilt_deg(height, site["dist"])
        nearest_bits.append(f"{site['name']} {site['dist']:.2f} km{height_bit}, look-down {look:.2f}°")
    if nearest_bits:
        geo.append("3. Nearest sites: " + "; ".join(nearest_bits) + ".")
    geo.append(
        "4. No terrain model is loaded. Elevation here is antenna height and the flat-ground look-down angle."
    )

    kind = str(choice.get("kind") or "")
    nearest = sites[0] if sites else None
    near_bit = (
        f"Nearest site {nearest['name']} is {nearest['dist']:.2f} km and {nearest['nearest_off']:.0f}° off the road."
        if nearest else "No site was found near this spot."
    )
    if kind == "new":
        cause = (
            f"1. This is a coverage hole. The poor run is {poor_m:.0f} m"
            + (f" and the blank is {blank_m:,.0f} m." if blank_m >= 80 else ".")
            + f" {near_bit} No host inside 3 km can take a 4th sector."
        )
        cause += "\n2. The logged servers are side-lobe or overshoot signals. No cell holds the road."
        cause += "\n3. Tilt cannot turn a beam or close a hole this far from the nearest site. A new 3-sector L900 site is the fix."
    elif kind == "fourth":
        cause = (
            f"1. The poor run is {poor_m:.0f} m"
            + (f" and the blank is {blank_m:,.0f} m." if blank_m >= 80 else ".")
            + f" {near_bit}"
        )
        cause += "\n2. The road sits off the existing beams, so the samples are side-lobe or overshoot energy."
        cause += "\n3. Tilt cannot rotate the beam onto the road. A 4th L900 sector on the host inside 3 km covers the hole."
    elif kind == "ep":
        cause = (
            f"1. {near_bit} That sector already faces the red road, within 12°, but it is missing from the log while a sibling sector is on air."
        )
        cause += "\n2. A retune would aim a sector that may already be correct. The azimuth has to be checked on site."
        cause += "\n3. The suggestion is On site EP Check. The proposed azimuth is left blank."
    elif kind == "azimuth":
        cause = (
            f"1. The poor run is {poor_m:.0f} m. {near_bit}"
        )
        cause += "\n2. The existing sector is close enough to retune, and the miss is horizontal."
        cause += "\n3. The azimuth turn below brings the beam onto the red road. Downtilt stays within 2°."
    else:
        cause = (
            f"1. The poor run is {poor_m:.0f} m. {near_bit}"
        )
        cause += "\n2. A far cell is overshooting the red samples. The azimuth already faces the road."
        cause += "\n3. Downtilt is limited to 2°. The beam is pulled back. The azimuth stays as it is."
    return {
        "serving": "\n".join(serving_lines),
        "geography": "\n".join(geo),
        "rca": cause,
        "suggestions": final,
    }


def _resolve_serving_cell(eci: object, earfcn: object) -> dict | None:
    """Site position of the cell that served this sample. Requires a Cell Id."""
    exact, pair = _serving_cell_index()
    number = pd.to_numeric(pd.Series([eci]), errors="coerce").iloc[0]
    if pd.isna(number):
        return None
    enb, cid = divmod(int(number), 256)
    ear = pd.to_numeric(pd.Series([earfcn]), errors="coerce").iloc[0]
    if pd.notna(ear):
        hit = exact.get((enb, cid, int(ear)))
        if hit is not None:
            return hit
    return pair.get((enb, cid))


def _samples_inside_outline(samples: pd.DataFrame, spot: pd.Series) -> pd.DataFrame:
    """Rows that sit inside this bad spot's orange circle."""
    if samples is None or samples.empty:
        return samples
    if "outline_width_km" not in spot.index or pd.isna(spot.get("outline_width_km")):
        return samples
    lon0 = float(spot["outline_lon"])
    lat0 = float(spot["outline_lat"])
    half_w = float(spot["outline_width_km"]) / _km_per_deg_lon(lat0) / 2.0
    half_h = float(spot["outline_height_km"]) / 111.32 / 2.0
    if half_w <= 0 or half_h <= 0:
        return samples
    lon = pd.to_numeric(samples["Longitude"], errors="coerce")
    lat = pd.to_numeric(samples["Latitude"], errors="coerce")
    inside = ((lon - lon0) / half_w) ** 2 + ((lat - lat0) / half_h) ** 2 <= 1.0
    return samples.loc[inside.fillna(False)]


def _red_bins_inside(samples: pd.DataFrame, column: str, spot: pd.Series) -> list[dict]:
    """50 m bins inside the orange circle whose median is poor, with the serving Cell Id.

    The colour on the map is that median, so the line ends on the red piece the
    reader sees. A bin with no Cell Id is kept for the length label and gets no line.
    """
    work = _samples_inside_outline(samples, spot)
    if work is None or work.empty or column not in work.columns or "Cell Id" not in work.columns:
        return []
    work = work.dropna(subset=["Longitude", "Latitude", column]).copy()
    if work.empty:
        return []
    rule = BAD_SPOT_RULES[column]
    threshold = float(rule["threshold"])
    strict = bool(rule.get("strict_less"))
    lat = pd.to_numeric(work["Latitude"], errors="coerce")
    lon = pd.to_numeric(work["Longitude"], errors="coerce")
    dlat, dlon = _bin_steps(lat, MAP_BIN_M)
    work["_by"] = np.floor(lat / dlat).astype("int64")
    work["_bx"] = np.floor(lon / dlon).astype("int64")
    bins: list[dict] = []
    for _, part in work.groupby(["_by", "_bx"], sort=False):
        values = pd.to_numeric(part[column], errors="coerce")
        if values.notna().sum() == 0:
            continue
        median = float(values.median())
        poor_bin = median < threshold if strict else median <= threshold
        if not poor_bin:
            continue
        poor_rows = part.loc[_poor_mask(part[column], threshold, strict_less=strict)]
        use = poor_rows if not poor_rows.empty else part
        cells = pd.to_numeric(use["Cell Id"], errors="coerce")
        identified = use.loc[cells.notna()]
        ends: list[dict] = []
        if not identified.empty:
            identified = identified.copy()
            identified["_eci"] = cells.loc[identified.index].astype("int64")
            for eci_val, grp in identified.groupby("_eci", sort=False):
                earfcn = None
                if "DL EARFCN" in grp.columns:
                    ears = pd.to_numeric(grp["DL EARFCN"], errors="coerce").dropna()
                    if not ears.empty:
                        earfcn = int(ears.mode().iloc[0])
                ends.append({
                    "lat": float(pd.to_numeric(grp["Latitude"], errors="coerce").mean()),
                    "lon": float(pd.to_numeric(grp["Longitude"], errors="coerce").mean()),
                    "eci": int(eci_val),
                    "earfcn": earfcn,
                })
        when = pd.to_datetime(part["Time"], errors="coerce").min() if "Time" in part.columns else pd.NaT
        bins.append({
            "lat": float(pd.to_numeric(part["Latitude"], errors="coerce").mean()),
            "lon": float(pd.to_numeric(part["Longitude"], errors="coerce").mean()),
            "eci": ends[0]["eci"] if ends else None,
            "earfcn": ends[0]["earfcn"] if ends else None,
            "ends": ends,
            "time": when,
        })
    bins.sort(key=lambda item: (pd.Timestamp.min if pd.isna(item["time"]) else item["time"], item["lat"], item["lon"]))
    return bins


def _red_runs(bins: list[dict]) -> list[list[dict]]:
    """Group red bins that touch the same stretch of road, whatever the time order."""
    n = len(bins)
    if n == 0:
        return []
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    # A short stretch of good coverage still belongs to the same red spot.
    limit_m = 400.0
    for i in range(n):
        for j in range(i + 1, n):
            gap_m = float(_equirect_km(bins[i]["lat"], bins[i]["lon"], bins[j]["lat"], bins[j]["lon"])) * 1000.0
            if gap_m <= limit_m:
                parent[find(j)] = find(i)
    groups: dict[int, list[dict]] = {}
    for i, item in enumerate(bins):
        groups.setdefault(find(i), []).append(item)
    return list(groups.values())


def _run_length_m(run: list[dict]) -> float:
    """Road length along the red bins, walking nearest neighbours so a bend is not a shortcut."""
    if len(run) <= 1:
        return float(MAP_BIN_M)
    pending = list(run)
    ordered = [pending.pop(0)]
    while pending:
        last = ordered[-1]
        nxt = min(
            pending,
            key=lambda item: float(_equirect_km(last["lat"], last["lon"], item["lat"], item["lon"])),
        )
        pending.remove(nxt)
        ordered.append(nxt)
    total = 0.0
    for left, right in zip(ordered, ordered[1:]):
        total += float(_equirect_km(left["lat"], left["lon"], right["lat"], right["lon"])) * 1000.0
    # Path from the first red bin to the last, plus one bin so both ends count.
    # A short good gap that joined the bins is included, so the label is the
    # full stretch rather than 50 m on every piece.
    return total + float(MAP_BIN_M)


_LABEL_STROKE = [pe.withStroke(linewidth=2.6, foreground="#14202B")]


def _ink(ax, lon: float, lat: float, text: str, size: float, color: str = "#F7FEFF") -> None:
    """Map text with a dark edge and no filled box."""
    ax.text(
        lon, lat, text,
        ha="center", va="center", fontsize=size, color=color, fontweight="bold",
        zorder=35, clip_on=True, path_effects=_LABEL_STROKE,
    )


def _inside_axes(ax, lon: float, lat: float, margin: float = 0.03) -> bool:
    x0, x1 = ax.get_xlim()
    y0, y1 = ax.get_ylim()
    mx = margin * (x1 - x0)
    my = margin * (y1 - y0)
    return (x0 + mx) <= lon <= (x1 - mx) and (y0 + my) <= lat <= (y1 - my)


def _px_shift(ax) -> tuple[float, float]:
    """Data degrees per screen pixel, from the axes box and its limits."""
    pos = ax.get_position()
    fig = ax.figure
    # savefig writes MAP_DPI, so a 16 px gap is 16 px in the PNG, not at the screen dpi.
    width = max(fig.get_figwidth() * MAP_DPI * pos.width, 1.0)
    height = max(fig.get_figheight() * MAP_DPI * pos.height, 1.0)
    x0, x1 = ax.get_xlim()
    y0, y1 = ax.get_ylim()
    return (x1 - x0) / width, (y1 - y0) / height


def _beside(
    ax, lon: float, lat: float, dx: float, dy: float, pixels: float, side: float,
) -> tuple[float, float]:
    """A point `pixels` to one side of the direction (dx, dy)."""
    dlon, dlat = _px_shift(ax)
    vx = dx / dlon if dlon else 0.0
    vy = dy / dlat if dlat else 0.0
    norm = math.hypot(vx, vy) or 1.0
    px, py = -vy / norm, vx / norm
    return lon + side * pixels * px * dlon, lat + side * pixels * py * dlat


def _px_from(ax, lon: float, lat: float, other_lon: float, other_lat: float) -> float:
    dlon, dlat = _px_shift(ax)
    return math.hypot(
        (lon - other_lon) / (dlon or 1e-12),
        (lat - other_lat) / (dlat or 1e-12),
    )


def _label_on_segment(
    ax, near_lon: float, near_lat: float, far_lon: float, far_lat: float,
    fracs: tuple[float, ...] = (0.30, 0.42, 0.20, 0.55),
    pixels: float = 16.0,
    side: float = 1.0,
    clear_px: float = 0.0,
    avoid: list[tuple[float, float]] | None = None,
) -> tuple[float, float]:
    """A point beside the line, off `near`, with no filled box.

    The offset is in screen pixels, measured after the axes has its final size.
    When `avoid` points are given, the side farther from them is used.
    """
    dx = far_lon - near_lon
    dy = far_lat - near_lat

    def clearance(point: tuple[float, float]) -> float:
        if not avoid:
            return 0.0
        return min(_px_from(ax, point[0], point[1], lon, lat) for lon, lat in avoid)

    if math.hypot(dx, dy) < 1e-12:
        _dlon, dlat = _px_shift(ax)
        up = (near_lon, near_lat + pixels * abs(dlat))
        down = (near_lon, near_lat - pixels * abs(dlat))
        if avoid:
            return max((up, down), key=clearance)
        return up if _inside_axes(ax, up[0], up[1]) else down

    best = None
    best_score = -1.0
    fallback = None
    for frac in fracs:
        lon = near_lon + frac * dx
        lat = near_lat + frac * dy
        choices = []
        for sign in (side, -side):
            point = _beside(ax, lon, lat, dx, dy, pixels, sign)
            if clear_px and _px_from(ax, point[0], point[1], near_lon, near_lat) < clear_px:
                continue
            choices.append(point)
        if not choices:
            continue
        choices.sort(key=clearance, reverse=True)
        if fallback is None:
            fallback = choices[0]
        for point in choices:
            if not _inside_axes(ax, point[0], point[1]):
                continue
            if not avoid:
                return point
            score = clearance(point)
            if score > best_score:
                best, best_score = point, score
            break
    if best is not None:
        return best
    if fallback is not None:
        return fallback
    return near_lon + fracs[0] * dx, near_lat + fracs[0] * dy


def _view_arm_deg(ax, spot_lat: float) -> float:
    """Same arm length the cyan sector mark uses, in degrees of latitude."""
    xlim, ylim = ax.get_xlim(), ax.get_ylim()
    span_km = min((xlim[1] - xlim[0]) * _km_per_deg_lon(spot_lat), (ylim[1] - ylim[0]) * 111.32)
    radius_km = float(np.clip(span_km * 0.07, 0.30, 0.85))
    return radius_km / 111.32


def _draw_serving_spider(
    ax, samples: pd.DataFrame, column: str, spot: pd.Series,
    avoid: list[tuple[float, float]] | None = None,
    overshoot: dict | None = None,
) -> None:
    """Dashed yellow line from the serving sector arm to each red bin.

    The line starts at the end of that cell's sector, not at the middle of the
    site, so the sector that served the sample is the one the line leaves.
    Every Cell Id on the red samples is drawn when it is an L900, L1800, L2100
    or L2600 cell in the system dump. The length label is the full red stretch.
    """
    bins = _red_bins_inside(samples, column, spot)
    arm = _view_arm_deg(ax, float(spot["lat"]))
    segments = []
    labeled: dict[tuple, dict] = {}
    for item in bins:
        for end in item.get("ends") or ():
            cell = _resolve_serving_cell(end["eci"], end["earfcn"])
            if cell is None:
                continue
            dist_km = float(_equirect_km(cell["lat"], cell["lon"], end["lat"], end["lon"]))
            if dist_km < 0.03:
                continue
            lat0 = cell["lat"]
            lon0 = cell["lon"]
            az = cell.get("azimuth")
            if az is not None:
                rad = np.radians(float(az) % 360.0)
                lon0 = lon0 + arm * np.sin(rad)
                lat0 = lat0 + arm * np.cos(rad)
            segments.append([(lon0, lat0), (end["lon"], end["lat"])])
            key = (cell.get("sector"), cell.get("layer"), round(float(cell["lat"]), 4), round(float(cell["lon"]), 4))
            # The number on the line is the length of that line, from the sector arm to the red bin.
            dist_m = float(_equirect_km(lat0, lon0, end["lat"], end["lon"])) * 1000.0
            current = labeled.get(key)
            if current is None or dist_m > current["dist_m"]:
                labeled[key] = {
                    "lon0": lon0, "lat0": lat0,
                    "lon1": float(end["lon"]), "lat1": float(end["lat"]),
                    "dist_m": dist_m,
                    "sector": str(cell.get("sector") or "").strip(),
                    "layer": cell.get("layer"),
                    "site_lat": float(cell["lat"]),
                    "site_lon": float(cell["lon"]),
                }
    if segments:
        ax.add_collection(LineCollection(
            segments,
            colors="#F5C400",
            linewidths=1.15,
            linestyles=(0, (6, 3.5)),
            capstyle="butt",
            zorder=23,
            alpha=0.95,
        ))
    for item in labeled.values():
        if item["dist_m"] < 80:
            continue
        # near = the red bin, so the words leave the bad samples and sit beside the line.
        mx, my = _label_on_segment(
            ax, item["lon1"], item["lat1"], item["lon0"], item["lat0"],
            fracs=(0.32, 0.46, 0.22, 0.58), pixels=28, avoid=avoid,
        )
        on_map = _inside_axes(ax, item["site_lon"], item["site_lat"], margin=0.0)
        text = f"{item['dist_m']:,.0f} m"
        sector = item.get("sector") or ""
        if sector and sector.lower() != "nan" and not on_map:
            text = f"{sector}\n{text}"
        mark = (overshoot or {}).get((item.get("sector"), item.get("layer")))
        if mark:
            text = f"{text}\n{mark}"
        _ink(ax, mx, my, text, 8, "#F7FEFF")
    spot_lon = float(spot["lon"])
    spot_lat = float(spot["lat"])
    for run in _red_runs(bins):
        first, last = run[0], run[-1]
        # Use the end farther from the spot, so the words are not on the site or the red samples.
        d0 = math.hypot(float(first["lon"]) - spot_lon, float(first["lat"]) - spot_lat)
        d1 = math.hypot(float(last["lon"]) - spot_lon, float(last["lat"]) - spot_lat)
        end, other = (first, last) if d0 >= d1 else (last, first)
        mx, my = _label_on_segment(
            ax, float(end["lon"]), float(end["lat"]),
            float(other["lon"]), float(other["lat"]),
            fracs=(0.0, 0.2, -0.25, -0.5, 0.4), pixels=40, avoid=avoid,
        )
        _ink(ax, mx, my, f"{_run_length_m(run):.0f} m", 8, "#FFE4E4")


def _proposal_arm_km(ax, lat: float, draw: dict, azimuth: float, facing: bool) -> float:
    """Arm length for a proposed sector. The arm toward the bad spot reaches past it."""
    del ax
    targets = draw.get("arm_targets") or {}
    hit = targets.get(round(float(azimuth), 1))
    if hit is not None:
        reach = float(_equirect_km(float(draw["lat"]), float(draw["lon"]), float(hit[0]), float(hit[1])))
        return max(0.9, reach * 1.15)
    side_km = 1.35 if draw.get("kind") == "new" else 0.9
    if not facing:
        return side_km
    reach = float(_equirect_km(draw["lat"], draw["lon"], draw["target_lat"], draw["target_lon"]))
    # A new site sits on the spot, so the facing arm is drawn long enough to cross it.
    # A 4th sector starts at the existing site, so its arm lands on the spot and a little past it.
    return max(side_km, reach * 1.25)


def _extent_with_proposal(extent: tuple[float, float, float, float], draw: dict) -> tuple[float, float, float, float]:
    """Widen a bad-spot view just enough to keep the proposed site and its arms."""
    lat = float(draw["lat"])
    reach = float(_equirect_km(draw["lat"], draw["lon"], draw["target_lat"], draw["target_lon"]))
    # New site: three arms of 1.35 km. 4th sector: one arm back toward the spot, so only a small margin past the site.
    margin_km = 1.7 if draw.get("kind") == "new" else 0.55
    if draw.get("kind") != "new":
        margin_km = max(margin_km, 0.35)
    pad_lon = margin_km / _km_per_deg_lon(lat)
    pad_lat = margin_km / 111.32
    # The facing arm may extend past the target. Keep that tip inside the frame.
    overshoot_km = max(0.0, max(1.35, reach * 1.25) - reach) if draw.get("kind") == "new" else max(0.0, reach * 0.25)
    tip_pad_lon = overshoot_km / _km_per_deg_lon(lat)
    tip_pad_lat = overshoot_km / 111.32
    lons = [float(draw["lon"]), float(draw["target_lon"])]
    lats = [lat, float(draw["target_lat"])]
    for point in (draw.get("arm_targets") or {}).values():
        lats.append(float(point[0]))
        lons.append(float(point[1]))
    for extra in draw.get("more") or []:
        lats.extend((float(extra["lat"]), float(extra["target_lat"])))
        lons.extend((float(extra["lon"]), float(extra["target_lon"])))
    return (
        min(extent[0], min(lons) - pad_lon, min(lons) - tip_pad_lon),
        max(extent[1], max(lons) + pad_lon, max(lons) + tip_pad_lon),
        min(extent[2], min(lats) - pad_lat, min(lats) - tip_pad_lat),
        max(extent[3], max(lats) + pad_lat, max(lats) + tip_pad_lat),
    )


def _draw_proposal(ax, draw: dict | None) -> None:
    """Draw a proposed new site, or one new 4th sector, long enough to reach the spot."""
    if not draw:
        return
    lat = float(draw["lat"])
    lon = float(draw["lon"])
    color = "#6C3483" if draw.get("kind") == "new" else "#D35400"
    azimuths = [float(az) for az in draw.get("azimuths") or []]
    for index, az in enumerate(azimuths):
        length = _proposal_arm_km(ax, lat, draw, az, index == 0)
        rad = np.radians(az % 360.0)
        x1 = lon + (length / _km_per_deg_lon(lat)) * np.sin(rad)
        y1 = lat + (length / 111.32) * np.cos(rad)
        ax.plot([lon, x1], [lat, y1], color="white", linewidth=6.0, solid_capstyle="round", zorder=31, clip_on=True)
        ax.plot([lon, x1], [lat, y1], color=color, linewidth=3.2, solid_capstyle="round", zorder=32, clip_on=True)
        # Azimuth sits toward the tip, on one side of the arm, clear of the site.
        mx, my = _label_on_segment(
            ax, lon, lat, x1, y1, (0.86, 0.74, 0.94), pixels=32, side=1.0, clear_px=48,
        )
        _ink(ax, mx, my, f"{az:.0f}°", 8, "#F7FEFF")
    ax.scatter([lon], [lat], s=42, c=color, marker="o", zorder=36, edgecolors="white", linewidths=0.9)
    if azimuths:
        face = float(azimuths[0])
        length = _proposal_arm_km(ax, lat, draw, face, True)
        rad = np.radians(face % 360.0)
        tip_lon = lon + (length / _km_per_deg_lon(lat)) * np.sin(rad)
        tip_lat = lat + (length / 111.32) * np.cos(rad)
        # The solution name is on the other side of the facing arm, not on the site.
        lx, ly = _label_on_segment(
            ax, lon, lat, tip_lon, tip_lat, (0.62, 0.50, 0.74), pixels=40, side=-1.0, clear_px=56,
        )
        _ink(ax, lx, ly, str(draw.get("label") or "Proposed"), 8, "#F7FEFF")
    for extra in draw.get("more") or []:
        elat, elon = float(extra["lat"]), float(extra["lon"])
        reach = float(_equirect_km(elat, elon, float(extra["target_lat"]), float(extra["target_lon"])))
        length = max(0.9, reach * 1.15)
        az = float(extra["azimuth"])
        rad = np.radians(az % 360.0)
        x1 = elon + (length / _km_per_deg_lon(elat)) * np.sin(rad)
        y1 = elat + (length / 111.32) * np.cos(rad)
        ax.plot([elon, x1], [elat, y1], color="white", linewidth=6.0, solid_capstyle="round", zorder=31, clip_on=True)
        ax.plot([elon, x1], [elat, y1], color="#D35400", linewidth=3.2, solid_capstyle="round", zorder=32, clip_on=True)
        mx, my = _label_on_segment(ax, elon, elat, x1, y1, (0.86, 0.74, 0.94), pixels=32, side=1.0, clear_px=48)
        _ink(ax, mx, my, f"{az:.0f}°", 8, "#F7FEFF")
        ax.scatter([elon], [elat], s=42, c="#D35400", marker="o", zorder=36, edgecolors="white", linewidths=0.9)


def draw_bad_spot_sector_pies(
    ax,
    spot: pd.Series,
    samples: pd.DataFrame | None = None,
    column: str | None = None,
    proposal: dict | None = None,
    overshoot: dict | None = None,
    turns: dict | None = None,
) -> None:
    """Transparent sector marks for every site that touches this bad-spot view.

    One cyan line per sector, with no fill. If any sector of a site falls inside
    the view, every sector of that site is drawn. When the samples under the map
    are passed in, a yellow line runs from the serving cell to each red bin
    inside the orange circle. Without those samples, the older azimuth line to
    the middle of the spot is kept.
    """
    sectors = load_site_sectors()
    if sectors.empty:
        return
    turns = turns or {}
    xlim, ylim = ax.get_xlim(), ax.get_ylim()
    spot_lat = float(spot["lat"])
    spot_lon = float(spot["lon"])
    inside = sectors["lat"].between(ylim[0], ylim[1]) & sectors["lon"].between(xlim[0], xlim[1])
    name_points: list[tuple[float, float]] = []
    if proposal:
        name_points.append((float(proposal["lon"]), float(proposal["lat"])))
    if not inside.any():
        if samples is not None and column:
            _draw_serving_spider(ax, samples, column, spot, avoid=name_points, overshoot=overshoot)
        return
    present = sectors.loc[inside, "SiteName"].unique()
    chosen = sectors[sectors["SiteName"].isin(present)]
    span_km = min((xlim[1] - xlim[0]) * _km_per_deg_lon(spot_lat), (ylim[1] - ylim[0]) * 111.32)
    radius_km = float(np.clip(span_km * 0.07, 0.30, 0.85))
    r_deg = radius_km / 111.32
    for site_name, part in chosen.groupby("SiteName", sort=False):
        part = part.drop_duplicates(subset=["Sector"]).dropna(subset=["azimuth"])
        if part.empty:
            continue
        here = part[inside.reindex(part.index).fillna(False).to_numpy()]
        if here.empty:
            continue
        lat = float(here["lat"].median())
        lon = float(here["lon"].median())
        bearing = _bearing_deg(lat, lon, spot_lat, spot_lon)
        facing: list[float] = []
        for _, sec in part.iterrows():
            name = str(sec["Sector"]).strip()
            # The existing arm stays on the dump azimuth and in the normal colour.
            # A proposed turn is a separate dashed arm, with the new azimuth on it.
            current = float(sec["azimuth"]) % 360.0
            proposed = None
            if name in turns:
                candidate = float(turns[name]) % 360.0
                if abs(_signed_deg(current, candidate)) >= 0.5:
                    proposed = candidate
            az = current
            rad = np.radians(az)
            if _ang_diff(az, bearing) <= BAD_SPOT_AIM_DEG:
                facing.append(az)
            ax.plot(
                [lon, lon + r_deg * np.sin(rad)],
                [lat, lat + r_deg * np.cos(rad)],
                color="#8FD8E0",
                linewidth=3.0,
                solid_capstyle="round",
                zorder=24,
                clip_on=True,
            )
            ax.plot(
                [lon, lon + r_deg * np.sin(rad)],
                [lat, lat + r_deg * np.cos(rad)],
                color="#E7FBFD",
                linewidth=1.5,
                solid_capstyle="round",
                zorder=25,
                clip_on=True,
            )
            name_xy = None
            if name and name.lower() != "nan":
                tip_lon = lon + r_deg * 1.28 * np.sin(rad)
                tip_lat = lat + r_deg * 1.28 * np.cos(rad)
                # One side of the arm. The solution name uses the other side, so the two do not stack.
                name_xy = _beside(ax, tip_lon, tip_lat, float(np.sin(rad)), float(np.cos(rad)), 48, 1.0)
            if proposed is not None:
                prad = np.radians(proposed)
                # Longer than the cyan arm, so the planned turn is not lost on the yellow serving line.
                reach = r_deg * 2.35
                x1 = lon + reach * np.sin(prad)
                y1 = lat + reach * np.cos(prad)
                ax.plot(
                    [lon, x1], [lat, y1],
                    color="white", linewidth=4.2, linestyle=(0, (8, 4.5)),
                    solid_capstyle="butt", zorder=28, clip_on=True,
                )
                ax.plot(
                    [lon, x1], [lat, y1],
                    color="#3D7EFF", linewidth=2.3, linestyle=(0, (8, 4.5)),
                    solid_capstyle="butt", zorder=29, clip_on=True,
                )
                delta = _signed_deg(current, proposed)
                # Out near the tip, in the wedge, and clear of the sector name on the current arm.
                p_dx, p_dy = math.sin(prad), math.cos(prad)
                c_dx, c_dy = math.sin(rad), math.cos(rad)
                toward_current = 1.0 if (-p_dy * c_dx + p_dx * c_dy) > 0 else -1.0
                avoid = [(spot_lon, spot_lat), *name_points]
                if name_xy is not None:
                    avoid.append(name_xy)
                deg_lon, deg_lat = _label_on_segment(
                    ax, lon, lat, x1, y1,
                    fracs=(0.88, 0.76, 0.96),
                    pixels=32,
                    side=toward_current,
                    clear_px=44,
                    avoid=avoid,
                )
                _ink(ax, deg_lon, deg_lat, f"{proposed:.0f}° ({delta:+.0f}°)", 8, "#D6E6FF")
                name_points.append((deg_lon, deg_lat))
                name_points.append((x1, y1))
            if name_xy is None:
                continue
            _ink(ax, name_xy[0], name_xy[1], name, 8)
            name_points.append(name_xy)
        if samples is not None and column:
            continue
        if not facing:
            continue
        dist_km = float(_equirect_km(lat, lon, spot_lat, spot_lon))
        if dist_km < radius_km * 1.8:
            continue
        text = f"{dist_km:.1f} kM" if dist_km < 10 else f"{dist_km:.0f} kM"
        for az in facing:
            rad = np.radians(az)
            x0 = lon + r_deg * np.sin(rad)
            y0 = lat + r_deg * np.cos(rad)
            ax.annotate(
                "",
                xy=(spot_lon, spot_lat),
                xytext=(x0, y0),
                arrowprops={"arrowstyle": "-|>", "color": "#F5C400", "lw": 1.4, "mutation_scale": 7},
                zorder=26,
                clip_on=True,
            )
            mx = (x0 + spot_lon) / 2.0
            my = (y0 + spot_lat) / 2.0
            lx, ly = _label_on_segment(ax, spot_lon, spot_lat, x0, y0, avoid=name_points)
            _ink(ax, lx, ly, text, 8, "#F7FEFF")
    if samples is not None and column:
        _draw_serving_spider(ax, samples, column, spot, avoid=name_points, overshoot=overshoot)


def _coverage_extent_from_geo(geo: pd.DataFrame) -> tuple[float, float, float, float]:
    lon_min, lon_max = float(geo["Longitude"].min()), float(geo["Longitude"].max())
    lat_min, lat_max = float(geo["Latitude"].min()), float(geo["Latitude"].max())
    pad_lon = 0.02 * (lon_max - lon_min)
    pad_lat = 0.02 * (lat_max - lat_min)
    return lon_min - pad_lon, lon_max + pad_lon, lat_min - pad_lat, lat_max + pad_lat


def _bin_steps(lat: pd.Series, bin_m: float) -> tuple[float, float]:
    """Bin grid in degrees. The reference latitude is rounded so that maps and bad
    spots land on the same grid even when built from slightly different subsets."""
    ref_lat = round(float(lat.median()), 1)
    return bin_m / 1000.0 / 111.32, bin_m / 1000.0 / _km_per_deg_lon(ref_lat)


def _bin_route_values(
    geo: pd.DataFrame,
    values: pd.Series,
    classifier,
    bin_m: float,
) -> tuple[pd.DataFrame, pd.Series]:
    """Reduce the route to one point per bin, coloured by that bin's median KPI.

    A city-scale map is about 200 m per pixel, so a raw scatter just shows
    whichever class happens to be drawn last over a whole segment of road.
    Binning makes the colour of a road mean the coverage measured on it.
    """
    lat = pd.to_numeric(geo["Latitude"], errors="coerce")
    lon = pd.to_numeric(geo["Longitude"], errors="coerce")
    val = pd.to_numeric(values, errors="coerce")
    ok = lat.notna() & lon.notna() & val.notna()
    if not ok.any():
        return geo.iloc[0:0], pd.Series(dtype=object)
    lat, lon, val = lat[ok], lon[ok], val[ok]
    dlat, dlon = _bin_steps(lat, bin_m)
    key = pd.MultiIndex.from_arrays(
        [np.floor(lat / dlat).astype(np.int64), np.floor(lon / dlon).astype(np.int64)]
    )
    binned = (
        pd.DataFrame(
            {"Latitude": lat.to_numpy(), "Longitude": lon.to_numpy(), "value": val.to_numpy()},
            index=key,
        )
        .groupby(level=[0, 1])
        .agg(Latitude=("Latitude", "mean"), Longitude=("Longitude", "mean"), value=("value", "median"))
        .reset_index(drop=True)
    )
    return binned, classifier(binned["value"])


# A drive-test export draws the route as a solid trace, not a cloud of
# sub-pixel dots. At city scale a 50 m bin is smaller than one pixel, so a
# scatter turns the two greens into one blurry line.
MAP_DPI = 200
MAP_ROUTE_LW = 3.2
_TILE_CACHE = Path("/tmp/bogura_sat_tiles")
_LIGHT_BASE_CACHE = Path("/tmp/bogura_light_base")
_BOUNDARY_CACHE = Path("/tmp/bogura_boundary_tiles")
_LIGHT_BASE_URL = (
    "https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}"
)
_BOUNDARY_URL = (
    "https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}"
)
_BASEMAP_CACHE: dict[tuple, object] = {}
# Circle geometries for the report, in pixels of the saved PNG (origin top-left).
# Each circle becomes its own Excel shape so it can be dragged on its own.
MAP_CIRCLE_PIXELS: dict[str, dict] = {}


def _merc_lat(y: float, n: float) -> float:
    merc = math.pi * (1.0 - 2.0 * y / n)
    return math.degrees(math.atan(math.sinh(merc)))


def _tile_bounds(x: int, y: int, z: int) -> tuple[float, float, float, float]:
    """West, east, south, north of one Web-Mercator tile."""
    n = 2**z
    west = x / n * 360.0 - 180.0
    east = (x + 1) / n * 360.0 - 180.0
    north = _merc_lat(y, n)
    south = _merc_lat(y + 1, n)
    return west, east, south, north


def _choose_zoom(span_deg: float) -> int:
    """Zoom where the satellite view stays within about twenty tiles on a side."""
    best = 10
    for z in range(9, 17):
        tiles = span_deg / (360.0 / 2**z)
        if tiles <= 20:
            best = z
        else:
            break
    return best


def _fetch_cached_tile(cache: Path, url: str, z: int, x: int, y: int):
    """Download one map tile into its own cache so styles never share a filename."""
    from PIL import Image

    for ext in (".jpg", ".png"):
        path = cache / f"{z}_{x}_{y}{ext}"
        if path.exists() and path.stat().st_size > 80:
            return Image.open(path)
    req = urllib.request.Request(url, headers={"User-Agent": "BoguraDriveTestReport/1.47 (coverage map export)"})
    cache.mkdir(parents=True, exist_ok=True)
    import time
    last_error: Exception | None = None
    data = b""
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = resp.read()
            break
        except Exception as exc:
            last_error = exc
            time.sleep(1.5 * (attempt + 1))
    else:
        raise last_error if last_error else RuntimeError(f"tile {z}/{x}/{y}")
    ext = ".jpg" if data[:3] == b"\xff\xd8\xff" else ".png"
    path = cache / f"{z}_{x}_{y}{ext}"
    path.write_bytes(data)
    return Image.open(path)


def _fetch_tile(z: int, x: int, y: int):
    # Esri World Imagery is z/y/x. x is the column, y is the row.
    url = (
        "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/"
        f"{z}/{y}/{x}"
    )
    return _fetch_cached_tile(_TILE_CACHE, url, z, x, y).convert("RGB")


def _add_basemap(ax) -> None:
    """Satellite image under the route."""
    try:
        _add_basemap_image(ax)
    except Exception as exc:
        print(f"Basemap skipped: {exc}")


def _add_basemap_image(ax) -> None:
    from PIL import Image

    west, east = ax.get_xlim()
    south, north = ax.get_ylim()
    if east <= west or north <= south:
        return
    z = _choose_zoom(max(east - west, north - south))
    n = 2**z
    x0 = max(0, int((west + 180.0) / 360.0 * n))
    x1 = min(n - 1, int((east + 180.0) / 360.0 * n))
    def tile_y(lat: float) -> int:
        lat_r = math.radians(max(-85.0, min(85.0, lat)))
        return int((1.0 - math.log(math.tan(lat_r) + 1.0 / math.cos(lat_r)) / math.pi) / 2.0 * n)
    y0 = max(0, min(tile_y(north), n - 1))
    y1 = max(0, min(tile_y(south), n - 1))
    if y1 < y0:
        y0, y1 = y1, y0
    key = (z, x0, x1, y0, y1)
    image = _BASEMAP_CACHE.get(key)
    if image is None:
        from concurrent.futures import ThreadPoolExecutor

        width = (x1 - x0 + 1) * 256
        height = (y1 - y0 + 1) * 256
        canvas = Image.new("RGB", (width, height), (40, 48, 36))
        coords = [(x, y) for y in range(y0, y1 + 1) for x in range(x0, x1 + 1)]

        def _load(xy: tuple[int, int]):
            x, y = xy
            try:
                return xy, _fetch_tile(z, x, y)
            except Exception:
                return xy, None

        with ThreadPoolExecutor(max_workers=8) as pool:
            fetched = list(pool.map(_load, coords))
        for (x, y), tile in fetched:
            if tile is None:
                continue
            canvas.paste(tile, ((x - x0) * 256, (y - y0) * 256))
        image = np.asarray(canvas)
        _BASEMAP_CACHE[key] = image
    west_m, _, _, north_m = _tile_bounds(x0, y0, z)
    _, east_m, south_m, _ = _tile_bounds(x1, y1, z)
    ax.imshow(
        image, extent=(west_m, east_m, south_m, north_m),
        origin="upper", interpolation="bilinear", zorder=0, aspect="auto",
    )
    ax.set_xlim(west, east)
    ax.set_ylim(south, north)
    ax.set_aspect("equal", adjustable="box")
    ax.text(
        0.01, 0.01, "Satellite © Esri",
        transform=ax.transAxes, fontsize=6, color="#FFFFFF", ha="left", va="bottom", zorder=6,
    )


def _compose_tiles(z: int, x0: int, x1: int, y0: int, y1: int, *, kind: str):
    """Stitch one basemap. Boundary maps are the light canvas plus admin lines."""
    from PIL import Image
    from concurrent.futures import ThreadPoolExecutor

    width = (x1 - x0 + 1) * 256
    height = (y1 - y0 + 1) * 256
    if kind == "boundary":
        canvas = Image.new("RGBA", (width, height), (245, 245, 245, 255))
        services = (
            (_LIGHT_BASE_CACHE, _LIGHT_BASE_URL, "RGB"),
            (_BOUNDARY_CACHE, _BOUNDARY_URL, "RGBA"),
        )
    else:
        canvas = Image.new("RGB", (width, height), (40, 48, 36))
        services = ((_TILE_CACHE, (
            "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"
        ), "RGB"),)
    coords = [(x, y) for y in range(y0, y1 + 1) for x in range(x0, x1 + 1)]

    def _load(item):
        cache, template, _mode = item[0]
        x, y = item[1]
        try:
            tile = _fetch_cached_tile(cache, template.format(z=z, y=y, x=x), z, x, y)
            return (x, y), tile
        except Exception:
            return (x, y), None

    # Fetch each service on its own so a missing boundary tile still leaves the canvas.
    for service in services:
        with ThreadPoolExecutor(max_workers=8) as pool:
            fetched = list(pool.map(_load, [(service, xy) for xy in coords]))
        for (x, y), tile in fetched:
            if tile is None:
                continue
            if tile.mode == "RGBA":
                canvas.paste(tile, ((x - x0) * 256, (y - y0) * 256), tile)
            else:
                canvas.paste(tile.convert("RGB"), ((x - x0) * 256, (y - y0) * 256))
    return np.asarray(canvas.convert("RGB"))


def _add_boundary_basemap(ax) -> None:
    """White admin-boundary map. The coloured KPI trace is the only line drawn on it."""
    try:
        west, east = ax.get_xlim()
        south, north = ax.get_ylim()
        if east <= west or north <= south:
            return
        z = _choose_zoom(max(east - west, north - south))
        n = 2**z
        x0 = max(0, int((west + 180.0) / 360.0 * n))
        x1 = min(n - 1, int((east + 180.0) / 360.0 * n))

        def tile_y(lat: float) -> int:
            lat_r = math.radians(max(-85.0, min(85.0, lat)))
            return int((1.0 - math.log(math.tan(lat_r) + 1.0 / math.cos(lat_r)) / math.pi) / 2.0 * n)

        y0 = max(0, min(tile_y(north), n - 1))
        y1 = max(0, min(tile_y(south), n - 1))
        if y1 < y0:
            y0, y1 = y1, y0
        key = ("boundary", z, x0, x1, y0, y1)
        image = _BASEMAP_CACHE.get(key)
        if image is None:
            image = _compose_tiles(z, x0, x1, y0, y1, kind="boundary")
            _BASEMAP_CACHE[key] = image
        west_m, _, _, north_m = _tile_bounds(x0, y0, z)
        _, east_m, south_m, _ = _tile_bounds(x1, y1, z)
        ax.imshow(
            image, extent=(west_m, east_m, south_m, north_m),
            origin="upper", interpolation="bilinear", zorder=0, aspect="auto",
        )
        ax.set_xlim(west, east)
        ax.set_ylim(south, north)
        ax.set_aspect("equal", adjustable="box")
        ax.text(
            0.99, 0.01, "Boundaries © Esri",
            transform=ax.transAxes, fontsize=6, color="#666666", ha="right", va="bottom", zorder=6,
        )
    except Exception as exc:
        print(f"Boundary basemap skipped: {exc}")


def _route_pieces(geo: pd.DataFrame, column: str, bin_m: float) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Time-ordered 50 m vertices, coloured by that bin's median, split on gaps."""
    work = geo.dropna(subset=["Longitude", "Latitude", column]).copy()
    if "Time" in work.columns:
        sort_cols = ["Source", "Time"] if "Source" in work.columns else ["Time"]
        work = work.sort_values(sort_cols)
    lat = pd.to_numeric(work["Latitude"], errors="coerce")
    lon = pd.to_numeric(work["Longitude"], errors="coerce")
    val = pd.to_numeric(work[column], errors="coerce")
    ok = lat.notna() & lon.notna() & val.notna()
    work = work.loc[ok]
    if work.empty:
        return np.empty((0, 2)), np.empty(0, dtype=object), np.empty(0, dtype=bool)
    lat = lat[ok]
    lon = lon[ok]
    val = val[ok]
    dlat, dlon = _bin_steps(lat, bin_m)
    key = pd.MultiIndex.from_arrays(
        [np.floor(lat / dlat).astype(np.int64), np.floor(lon / dlon).astype(np.int64)]
    )
    median = val.groupby(key).transform("median")
    classes = MAP_CLASSIFIERS[column](median)
    # One vertex per visit to a bin, in drive order, so the trace follows the road.
    changed = np.empty(len(work), dtype=bool)
    changed[0] = True
    keys = key.to_numpy()
    changed[1:] = keys[1:] != keys[:-1]
    step = np.zeros(len(work))
    step[1:] = _equirect_km(lat.to_numpy()[:-1], lon.to_numpy()[:-1], lat.to_numpy()[1:], lon.to_numpy()[1:]) * 1000.0
    source = work["Source"].to_numpy() if "Source" in work.columns else np.zeros(len(work))
    new_pass = np.zeros(len(work), dtype=bool)
    new_pass[1:] = (source[1:] != source[:-1]) | (step[1:] > 120.0)
    keep = changed | new_pass
    pts = np.column_stack([lon.to_numpy()[keep], lat.to_numpy()[keep]])
    return pts, classes.to_numpy()[keep], new_pass[keep]


def _draw_coverage_points(
    ax,
    geo: pd.DataFrame,
    class_series: pd.Series,
    labels: list[str],
    colors: list[str],
    max_points: int = MAP_MAX_POINTS,
    *,
    poor_on_top: bool = True,
    bin_m: float | None = None,
    column: str | None = None,
    route_casing: bool = True,
) -> None:
    """Solid drive trace, worst colour on top, same 50 m bins as the bad spots.

    A dark casing under the colour keeps yellow and the two greens readable on
    the street map, the way a log-export route sits on the road.
    """
    if geo.empty:
        return
    classifier = MAP_CLASSIFIERS.get(column or "")
    if not (bin_m and classifier is not None and column in geo.columns):
        sample = downsample(geo, max_points=max_points)
        classes = class_series.reindex(sample.index)
        for label, color in zip(labels, colors):
            part = sample[classes == label]
            if part.empty:
                continue
            ax.scatter(part["Longitude"], part["Latitude"], s=12, c=color, linewidths=0, zorder=3)
        return
    pts, classes, new_pass = _route_pieces(geo, column, bin_m)
    if len(pts) == 0:
        return
    color_of = dict(zip(labels, colors))
    order = list(labels) if poor_on_top else list(reversed(labels))
    # Consecutive segments of one class become one polyline, so a vector export
    # has clean joins and a few thousand paths rather than one per 50 m.
    segments: dict[str, list] = {lab: [] for lab in labels}
    casing: list = []
    run_start = 0
    breaks = list(np.flatnonzero(new_pass)) + [len(pts)]
    alone: dict[str, list] = {lab: [] for lab in labels}
    for end in breaks:
        if end - run_start == 1:
            lab = classes[run_start]
            if lab in alone:
                alone[lab].append(pts[run_start])
        elif end - run_start > 1:
            casing.append(pts[run_start:end])
            i = run_start
            while i < end - 1:
                lab = classes[i]
                j = i
                while j + 1 < end - 1 and classes[j + 1] == lab:
                    j += 1
                if lab in segments:
                    segments[lab].append(pts[i:j + 2])
                i = j + 1
        run_start = end
    # A white casing keeps the route readable on satellite imagery.
    # Bad-spot boundary maps leave it off: the coloured KPI trace is the chart.
    if casing and route_casing:
        ax.add_collection(LineCollection(
            casing, colors="#FFFFFF", linewidths=MAP_ROUTE_LW + 2.4,
            capstyle="round", joinstyle="round", zorder=2,
        ))
    for z, lab in enumerate(order):
        segs = segments.get(lab) or []
        if not segs:
            continue
        ax.add_collection(LineCollection(
            segs, colors=color_of[lab], linewidths=MAP_ROUTE_LW,
            capstyle="round", joinstyle="round", zorder=3 + z,
        ))
        if alone[lab]:
            xy = np.vstack(alone[lab])
            ax.scatter(xy[:, 0], xy[:, 1], s=18, c=color_of[lab], linewidths=0, zorder=3 + z)


def _legend_sample_counts(df: pd.DataFrame, column: str, classifier, labels: list[str]) -> list[int]:
    """Sample counts for the rows actually drawn, one number per legend bin."""
    classes = classifier(df[column])
    counts = classes.value_counts(dropna=True)
    return [int(counts.get(lab, 0)) for lab in labels]


def _count_legend_handles(labels: list[str], colors: list[str], counts: list[int]) -> list:
    """Square swatches with the bin count in parentheses, matching the coverage legend."""
    from matplotlib.patches import Patch

    return [
        Patch(
            facecolor=color,
            edgecolor="#333333",
            linewidth=0.6,
            label=f"{label} ({count})",
        )
        for label, color, count in zip(labels, colors, counts)
    ]


def _place_count_legend(ax, handles: list, title: str):
    """Legend box inside the map, on the left, as in the coverage export."""
    legend = ax.legend(
        handles=handles,
        title=title,
        loc="upper left",
        bbox_to_anchor=(0.015, 0.985),
        frameon=True,
        fontsize=8,
        title_fontsize=9,
        borderpad=0.7,
        labelspacing=0.4,
        handlelength=1.15,
        handleheight=1.05,
        borderaxespad=0.2,
    )
    frame = legend.get_frame()
    frame.set_edgecolor("#000000")
    frame.set_linewidth(0.9)
    frame.set_facecolor("#FFFFFF")
    frame.set_alpha(1.0)
    legend.get_title().set_fontweight("bold")
    # The poorest colour is drawn above the other bins, so the box has to sit higher
    # than that trace. The legend then covers the road the way the reference export does.
    legend.set_zorder(20)
    return legend


def _coverage_legend_handles(labels: list[str], colors: list[str]) -> list:
    return [
        Line2D(
            [0],
            [0],
            marker="o",
            color="none",
            markerfacecolor=c,
            markeredgecolor="#333333",
            markeredgewidth=0.4,
            markersize=8,
            label=lab,
        )
        for lab, c in zip(labels, colors)
    ]


def plot_discrete_coverage_map(
    df: pd.DataFrame,
    path: Path,
    column: str,
    class_series: pd.Series,
    labels: list[str],
    colors: list[str],
    title: str,
    legend_title: str,
    max_points: int = MAP_MAX_POINTS,
    extent: tuple[float, float, float, float] | None = None,
    layer: str | None = None,
    empty_message: str = "No IDLE samples on this band",
) -> Path:
    """Drive route colored by discrete KPI ranges, with an RSRP-style side legend."""
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    geo = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    fig, ax = plt.subplots(figsize=(8.0, 9.2))
    _draw_coverage_points(
        ax, geo, class_series, labels, colors, max_points=max_points, bin_m=MAP_BIN_M, column=column
    )
    if extent is not None:
        ax.set_xlim(extent[0], extent[1])
        ax.set_ylim(extent[2], extent[3])
    elif not geo.empty:
        lon_min, lon_max, lat_min, lat_max = _coverage_extent_from_geo(geo)
        ax.set_xlim(lon_min, lon_max)
        ax.set_ylim(lat_min, lat_max)
    if geo.empty:
        ax.text(
            0.5,
            0.5,
            empty_message,
            transform=ax.transAxes,
            ha="center",
            va="center",
            color="#7F8C8D",
            fontsize=12,
        )
    ax.legend(
        handles=_coverage_legend_handles(labels, colors),
        title=legend_title,
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=True,
        fontsize=8,
        title_fontsize=10,
    )
    ax.set_aspect("equal", adjustable="box")
    _add_basemap(ax)
    add_map_scale_bar(ax)
    ax.set_title(title)
    hide_map_axes(ax)
    fig.tight_layout()
    fig.savefig(path, dpi=MAP_DPI, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return path


def plot_road_map(
    df: pd.DataFrame,
    path: Path,
    extent: tuple[float, float, float, float] | None = None,
) -> Path:
    """Satellite trace of the drive. One road colour, no legend and no mode title."""
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    geo = df.dropna(subset=["Longitude", "Latitude", "RSRP"]).copy()
    fig, ax = plt.subplots(figsize=(7.4, 9.2))
    if geo.empty:
        ax.text(
            0.5, 0.5, "No route", transform=ax.transAxes,
            ha="center", va="center", color="#7F8C8D", fontsize=12,
        )
    else:
        pts, _classes, new_pass = _route_pieces(geo, "RSRP", MAP_BIN_M)
        segments: list = []
        run = 0
        breaks = list(np.flatnonzero(new_pass)) + [len(pts)]
        for end in breaks:
            if end - run > 1:
                segments.append(pts[run:end])
            run = end
        if segments:
            ax.add_collection(LineCollection(
                segments, colors="#1B4F72", linewidths=MAP_ROUTE_LW + 2.6,
                capstyle="round", joinstyle="round", zorder=2,
            ))
            ax.add_collection(LineCollection(
                segments, colors="#F4F7FB", linewidths=MAP_ROUTE_LW,
                capstyle="round", joinstyle="round", zorder=3,
            ))
        if extent is not None:
            ax.set_xlim(extent[0], extent[1])
            ax.set_ylim(extent[2], extent[3])
        else:
            lon_min, lon_max, lat_min, lat_max = _coverage_extent_from_geo(geo)
            ax.set_xlim(lon_min, lon_max)
            ax.set_ylim(lat_min, lat_max)
    ax.set_aspect("equal", adjustable="box")
    _add_basemap(ax)
    add_map_scale_bar(ax)
    hide_map_axes(ax)
    fig.tight_layout()
    fig.savefig(path, dpi=MAP_DPI, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return path


def plot_rsrp_coverage_map(
    df: pd.DataFrame,
    path: Path,
    max_points: int = MAP_MAX_POINTS,
    *,
    layer: str | None = None,
    extent: tuple[float, float, float, float] | None = None,
    mode_label: str | None = None,
) -> Path:
    """Discrete RSRP coverage map with a range legend (no lon/lat chrome)."""
    title = "Bogura coverage map — RSRP"
    if mode_label:
        title = f"{title} — {mode_label}"
    if layer:
        title = f"{title} — {layer}"
    empty = f"No {mode_label} samples on this band" if mode_label else "No IDLE samples on this band"
    return plot_discrete_coverage_map(
        df,
        path,
        column="RSRP",
        class_series=rsrp_map_class(df["RSRP"]),
        labels=MAP_RSRP_LABELS,
        colors=MAP_RSRP_COLORS,
        title=title,
        legend_title="RSRP (dBm)",
        max_points=max_points,
        extent=extent,
        layer=layer,
        empty_message=empty,
    )


def plot_rsrq_coverage_map(
    df: pd.DataFrame,
    path: Path,
    max_points: int = MAP_MAX_POINTS,
    *,
    layer: str | None = None,
    extent: tuple[float, float, float, float] | None = None,
    mode_label: str | None = None,
) -> Path:
    title = "Bogura coverage map — RSRQ"
    if mode_label:
        title = f"{title} — {mode_label}"
    if layer:
        title = f"{title} — {layer}"
    empty = f"No {mode_label} samples on this band" if mode_label else "No IDLE samples on this band"
    return plot_discrete_coverage_map(
        df,
        path,
        column="RSRQ",
        class_series=rsrq_map_class(df["RSRQ"]),
        labels=MAP_RSRQ_LABELS,
        colors=MAP_RSRQ_COLORS,
        title=title,
        legend_title="RSRQ (dB)",
        max_points=max_points,
        extent=extent,
        layer=layer,
        empty_message=empty,
    )


def plot_sinr_coverage_map(
    df: pd.DataFrame,
    path: Path,
    max_points: int = MAP_MAX_POINTS,
    *,
    layer: str | None = None,
    extent: tuple[float, float, float, float] | None = None,
    mode_label: str | None = None,
) -> Path:
    title = "Bogura coverage map — SINR"
    if mode_label:
        title = f"{title} — {mode_label}"
    if layer:
        title = f"{title} — {layer}"
    empty = f"No {mode_label} samples on this band" if mode_label else "No IDLE samples on this band"
    return plot_discrete_coverage_map(
        df,
        path,
        column="SINR",
        class_series=sinr_map_class(df["SINR"]),
        labels=MAP_SINR_LABELS,
        colors=MAP_SINR_COLORS,
        title=title,
        legend_title="SINR (dB)",
        max_points=max_points,
        extent=extent,
        layer=layer,
        empty_message=empty,
    )


# Bad-spot contract from report v1.35.
# 1) Poor is RSRP < -120 dBm (red bin only; -120 through -115 is light red),
#    RSRQ <= -20 dB (inside the red bin), SINR <= 0 dB (orange and red).
# 2) A bin of road is 50 m. Its colour and its poor/good call both use the median
#    of the samples in it, and poor bins are drawn on top of good ones.
# 3) Every poor run is covered. A run stays one run across a gap of at most 150 m
#    and 60 s on the same pass. Runs within 2 km of each other join one circle
#    when the poor road inside that circle still spans at most 7 km, so the drawn
#    circle stays within 10 km and covers the road. A longer poor road is split
#    into several circles. Pockets more than 2 km apart stay separate circles
#    even when both sit inside the same 10 km neighbourhood, and each circle is
#    its own shape so it can be moved on its own.
# 4) Leftover poor bins that cover >= 1 km² stay a discrete (dotted) circle.
#    With rule 3 taking every run, that count is 0 on this drive; the rule stays.
# Idle best server is the strongest band in each scanner sweep (4 s or 40 m).
# Active best server is the strongest Robi carrier sharing one timestamp.
BAD_SPOT_MIN_CONSEC_M = 200.0
BAD_SPOT_DISCRETE_AREA_KM2 = 1.0
# Best-server samples are one scanner sweep apart (<= 40 m), so a poor stretch may
# skip a sweep or two and still be the same hole.
BAD_SPOT_MAX_SAMPLE_GAP_M = 150.0
BAD_SPOT_MAX_SAMPLE_GAP_S = 60.0
BAD_SPOT_GRID_DEG = 0.009  # ~1.0 km cells for discrete-area clustering
BAD_SPOT_MAX_CONSEC = 30
BAD_SPOT_MAX_DISCRETE = 15
BAD_SPOT_DISCRETE_MIN_POOR = 20
BAD_SPOT_DISCRETE_MAX_SPAN_KM = 3.2
# Calling a whole km² bad means the roads inside it are bad, not that a few
# scattered samples in it are.
BAD_SPOT_DISCRETE_MIN_SHARE = 0.50
# Poor-road span that still fits inside a drawn circle of BAD_SPOT_CIRCLE_MAX_KM.
BAD_SPOT_GROUP_SPAN_KM = 7.0
BAD_SPOT_CIRCLE_MAX_KM = 10.0
BAD_SPOT_CIRCLE_MIN_KM = 2.5
BAD_SPOT_MERGE_GAP_KM = 2.0
BAD_SPOT_CONSEC_MAX_SPAN_KM = BAD_SPOT_GROUP_SPAN_KM
BAD_SPOT_ZOOM_PAD_KM = 0.70
BAD_SPOT_ZOOM_MIN_KM = 1.2
BAD_SPOT_OUTLINE_PAD = 1.45  # with sqrt(2), oval covers bbox corners / red tails
BAD_SPOT_OUTLINE_MIN_MINOR_KM = 0.50
BAD_SPOT_OVERVIEW_MIN_KM = 2.6  # city-scale combined map must show a visible circle
BAD_SPOT_OVERVIEW_MAX_KM = 4.5
BAD_SPOT_OVERVIEW_LW = 2.15
# A few spots can each carry a big haloed ring. Hundreds cannot: the rings would
# cover the colour they are marking, so they shrink and thin as the count grows.
BAD_SPOT_OVERVIEW_TIERS = (
    (40, 2.6, 2.15, True),
    (250, 1.4, 0.70, False),
    (10**9, 0.8, 0.50, False),
)
BAD_SPOT_ZOOM_COLS = 4
BAD_SPOT_MAX_ZOOMS = 300
BAD_SPOT_MAX = BAD_SPOT_MAX_CONSEC + BAD_SPOT_MAX_DISCRETE
BAD_SPOT_RULES = {
    # strict_less: a value equal to the threshold is not poor. -120 is the top of the red bin.
    "RSRP": {"threshold": -120.0, "unit": "dBm", "min_consecutive_m": 0.0, "max_consecutive": 2000, "strict_less": True},
    "RSRQ": {"threshold": -20.0, "unit": "dB", "min_consecutive_m": 0.0, "max_consecutive": 2000, "strict_less": False},
    "SINR": {"threshold": 0.0, "unit": "dB", "min_consecutive_m": 0.0, "max_consecutive": 2000, "strict_less": False},
}


def min_consecutive_m(column: str) -> float:
    return float(BAD_SPOT_RULES[column].get("min_consecutive_m", BAD_SPOT_MIN_CONSEC_M))


# Expected values for the contract above. Kept separate from the live constants so
# a later edit that retunes a rule fails the report check instead of passing quietly.
# Spot counts are circles and discrete areas per sheet.
# Idle stays on the published Borga_Idle_V2.0 result. Active follows
# Borga_Active_mode_11Nov26.7z: Robi M-PCell only, strongest RSRP at each location.
LOGIC_LOCK = {
    "active_operator": "Robi",
    "sweep_s": 4.0,
    "sweep_m": 40.0,
    "bin_m": 50.0,
    "gap_m": 150.0,
    "gap_s": 60.0,
    "span_km": 7.0,
    "circle_max_km": 10.0,
    "circle_min_km": 2.5,
    "merge_gap_km": 2.0,
    "discrete_km2": 1.0,
    "discrete_share": 0.50,
    "poor": {"RSRP": -120.0, "RSRQ": -20.0, "SINR": 0.0},
    "poor_strict": {"RSRP": True, "RSRQ": False, "SINR": False},
    "min_consecutive_m": {"RSRP": 0.0, "RSRQ": 0.0, "SINR": 0.0},
    "rsrp_greens": ("#008000", "#00FF00"),
    "ring_tiers": ((40, 2.6, 2.15, True), (250, 1.4, 0.70, False), (10**9, 0.8, 0.50, False)),
    # RSRQ and SINR bad-spot sheets are not in the workbook. The RSRP sheets
    # keep these circle counts. RSRQ and SINR rules still colour the KPI maps.
    "counts": {
        "Bad Spot RSRP": (26, 0),
        "Bad Spot RSRP_Active": (99, 0),
    },
}


def logic_drift() -> list[str]:
    """Report any live constant that no longer matches LOGIC_LOCK."""
    problems = []
    if ACTIVE_OPERATOR != LOGIC_LOCK["active_operator"]:
        problems.append(f"active operator is {ACTIVE_OPERATOR}")
    live_windows = (
        BEST_SERVER_SWEEP_S, BEST_SERVER_SWEEP_M, MAP_BIN_M,
        BAD_SPOT_MAX_SAMPLE_GAP_M, BAD_SPOT_MAX_SAMPLE_GAP_S, BAD_SPOT_CONSEC_MAX_SPAN_KM,
        BAD_SPOT_CIRCLE_MAX_KM, BAD_SPOT_CIRCLE_MIN_KM, BAD_SPOT_MERGE_GAP_KM,
        BAD_SPOT_DISCRETE_AREA_KM2, BAD_SPOT_DISCRETE_MIN_SHARE,
    )
    locked_windows = (
        LOGIC_LOCK["sweep_s"], LOGIC_LOCK["sweep_m"], LOGIC_LOCK["bin_m"],
        LOGIC_LOCK["gap_m"], LOGIC_LOCK["gap_s"], LOGIC_LOCK["span_km"],
        LOGIC_LOCK["circle_max_km"], LOGIC_LOCK["circle_min_km"], LOGIC_LOCK["merge_gap_km"],
        LOGIC_LOCK["discrete_km2"], LOGIC_LOCK["discrete_share"],
    )
    if live_windows != locked_windows:
        problems.append(f"distance/time windows {live_windows} != {locked_windows}")
    for kpi, threshold in LOGIC_LOCK["poor"].items():
        rule = BAD_SPOT_RULES[kpi]
        if float(rule["threshold"]) != threshold or float(rule["min_consecutive_m"]) != LOGIC_LOCK["min_consecutive_m"][kpi]:
            problems.append(f"{kpi} rule {rule['threshold']}, min {rule['min_consecutive_m']}")
        if bool(rule.get("strict_less")) != bool(LOGIC_LOCK["poor_strict"][kpi]):
            problems.append(f"{kpi} strict_less is {bool(rule.get('strict_less'))}")
    if tuple(MAP_RSRP_COLORS[:2]) != LOGIC_LOCK["rsrp_greens"]:
        problems.append(f"RSRP greens are {MAP_RSRP_COLORS[:2]}")
    if BAD_SPOT_OVERVIEW_TIERS != LOGIC_LOCK["ring_tiers"]:
        problems.append("overview ring tiers changed")
    return problems


def consecutive_label(column: str) -> str:
    """Legend / table wording for the consecutive rule actually applied to this KPI."""
    minimum = min_consecutive_m(column)
    if minimum <= 0:
        return "Poor run (any length)"
    return f"Consecutive \u2265 {minimum:,.0f} m"


def _km_per_deg_lon(lat: float) -> float:
    return 111.32 * float(np.cos(np.radians(lat)))


def _equirect_km(lat1, lon1, lat2, lon2) -> np.ndarray:
    mid = np.radians((np.asarray(lat1) + np.asarray(lat2)) / 2.0)
    dlat = (np.asarray(lat2) - np.asarray(lat1)) * 111.32
    dlon = (np.asarray(lon2) - np.asarray(lon1)) * 111.32 * np.cos(mid)
    return np.sqrt(dlat**2 + dlon**2)


def _poor_mask(values: pd.Series, threshold: float, *, strict_less: bool = False) -> pd.Series:
    """Poor samples. RSRP is strictly below -120. Light red is not poor. RSRQ and SINR are at or below their threshold."""
    v = pd.to_numeric(values, errors="coerce")
    poor = v < threshold if strict_less else v <= threshold
    return v.notna() & poor


def _prepare_track(df: pd.DataFrame, column: str) -> pd.DataFrame:
    g = df.dropna(subset=["Time", "Longitude", "Latitude", column]).copy()
    g[column] = pd.to_numeric(g[column], errors="coerce")
    g["Longitude"] = pd.to_numeric(g["Longitude"], errors="coerce")
    g["Latitude"] = pd.to_numeric(g["Latitude"], errors="coerce")
    g = g.dropna(subset=[column, "Longitude", "Latitude", "Time"])
    sort_cols = ["Source", "Time"] if "Source" in g.columns else ["Time"]
    return g.sort_values(sort_cols).reset_index(drop=True)


def _convex_hull_xy(points: np.ndarray) -> np.ndarray:
    """Monotone-chain convex hull for (n, 2) points."""
    pts = np.asarray(points, dtype=float)
    if len(pts) <= 1:
        return pts
    pts = pts[np.lexsort((pts[:, 1], pts[:, 0]))]
    uniq = np.concatenate([pts[0:1], pts[1:][np.any(np.diff(pts, axis=0) != 0, axis=1)]])
    if len(uniq) <= 2:
        return uniq

    def cross(o, a, b) -> float:
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    lower: list[np.ndarray] = []
    for p in uniq:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    upper: list[np.ndarray] = []
    for p in uniq[::-1]:
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return np.vstack(lower[:-1] + upper[:-1])


def _hull_area_km2(lons: np.ndarray, lats: np.ndarray) -> float:
    if len(lons) < 3:
        return 0.0
    lat0 = float(np.mean(lats))
    lon0 = float(np.mean(lons))
    xy = np.column_stack(((lons - lon0) * _km_per_deg_lon(lat0), (lats - lat0) * 111.32))
    hull = _convex_hull_xy(xy)
    if len(hull) < 3:
        return 0.0
    x, y = hull[:, 0], hull[:, 1]
    return float(abs(np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1))) / 2.0)


def _path_length_m(lats: np.ndarray, lons: np.ndarray) -> float:
    if len(lats) < 2:
        return 0.0
    return float(np.sum(_equirect_km(lats[:-1], lons[:-1], lats[1:], lons[1:])) * 1000.0)


def _covering_ellipse(
    lons: np.ndarray,
    lats: np.ndarray,
    *,
    pad: float = BAD_SPOT_OUTLINE_PAD,
    min_minor_km: float = BAD_SPOT_OUTLINE_MIN_MINOR_KM,
    min_major_km: float = 0.55,
) -> tuple[float, float, float, float, float]:
    """Axis-aligned oval that contains the whole poor stretch, including corners and red tails."""
    lons = np.asarray(lons, dtype=float)
    lats = np.asarray(lats, dtype=float)
    lat0 = float((np.min(lats) + np.max(lats)) / 2.0)
    lon0 = float((np.min(lons) + np.max(lons)) / 2.0)
    km_lon = _km_per_deg_lon(lat0)
    span_lon_km = (float(np.max(lons)) - float(np.min(lons))) * km_lon
    span_lat_km = (float(np.max(lats)) - float(np.min(lats))) * 111.32
    # Inscribed ellipse (width=bbox) leaves diagonal tails outside. Scale so corners stay in.
    cover = float(np.sqrt(2.0)) * pad / 1.15
    width_km = max(span_lon_km * cover, min_major_km)
    height_km = max(span_lat_km * cover, min_minor_km)
    return lon0, lat0, width_km, height_km, 0.0


def _outline_fields(lons: np.ndarray, lats: np.ndarray, kind: str) -> dict:
    lon0, lat0, width_km, height_km, angle = _covering_ellipse(lons, lats)
    width_km = min(BAD_SPOT_CIRCLE_MAX_KM, max(width_km, BAD_SPOT_CIRCLE_MIN_KM))
    height_km = min(BAD_SPOT_CIRCLE_MAX_KM, max(height_km, BAD_SPOT_CIRCLE_MIN_KM))
    if kind == "discrete":
        diameter_km = min(BAD_SPOT_CIRCLE_MAX_KM, max(width_km, height_km))
        width_km = height_km = diameter_km
        angle = 0.0
    return {
        "outline_lon": lon0,
        "outline_lat": lat0,
        "outline_width_km": float(width_km),
        "outline_height_km": float(height_km),
        "outline_angle": float(angle),
    }


def _spot_from_points(
    part: pd.DataFrame,
    column: str,
    threshold: float,
    kind: str,
    length_m: float,
    area_km2: float,
) -> dict:
    lat = float(part["Latitude"].mean())
    lon = float(part["Longitude"].mean())
    lat_min = float(part["Latitude"].min())
    lat_max = float(part["Latitude"].max())
    lon_min = float(part["Longitude"].min())
    lon_max = float(part["Longitude"].max())
    vals = pd.to_numeric(part[column], errors="coerce")
    n = int(vals.notna().sum())
    n_poor = int(_poor_mask(vals, threshold, strict_less=bool(BAD_SPOT_RULES[column].get("strict_less"))).sum())
    mean = float(vals.mean()) if n else np.nan
    return {
        "kind": kind,
        "n": n,
        "n_poor": n_poor,
        "poor_pct": n_poor / max(n, 1),
        "mean": mean,
        "lat": lat,
        "lon": lon,
        "lat_min": lat_min,
        "lat_max": lat_max,
        "lon_min": lon_min,
        "lon_max": lon_max,
        "span_lat_km": (lat_max - lat_min) * 111.32,
        "span_lon_km": (lon_max - lon_min) * _km_per_deg_lon(lat),
        "length_m": float(length_m),
        "area_km2": float(area_km2),
        "score": (length_m if kind == "consecutive" else area_km2 * 1000.0),
        **_outline_fields(part["Longitude"].to_numpy(), part["Latitude"].to_numpy(), kind),
    }


def _span_km(lats: np.ndarray, lons: np.ndarray) -> float:
    if len(lats) == 0:
        return 0.0
    lat0 = float(np.mean(lats))
    return max(
        (float(np.max(lats)) - float(np.min(lats))) * 111.32,
        (float(np.max(lons)) - float(np.min(lons))) * _km_per_deg_lon(lat0),
    )


def binned_kpi(geo: pd.DataFrame, column: str, bin_m: float = MAP_BIN_M) -> pd.Series:
    """Median KPI of each bin of road, spread back over the samples in that bin.

    Bad spots are judged on this, not on single samples, and the maps are drawn
    from the same bins: a stretch is circled only where the map shows it poor.
    """
    lat = pd.to_numeric(geo["Latitude"], errors="coerce")
    lon = pd.to_numeric(geo["Longitude"], errors="coerce")
    val = pd.to_numeric(geo[column], errors="coerce")
    if val.notna().sum() == 0:
        return val
    dlat, dlon = _bin_steps(lat, bin_m)
    key = pd.MultiIndex.from_arrays(
        [np.floor(lat / dlat).astype("Int64"), np.floor(lon / dlon).astype("Int64")]
    )
    return val.groupby(key).transform("median")


def _poor_runs(
    poor_idx: np.ndarray,
    same_pass: np.ndarray,
    cum_m: np.ndarray,
    time_s: np.ndarray,
) -> list[list[int]]:
    """Chain poor samples of one pass into runs, tolerating a short good gap."""
    runs: list[list[int]] = []
    if len(poor_idx) == 0:
        return runs
    current = [int(poor_idx[0])]
    for prev, cur in zip(poor_idx[:-1], poor_idx[1:]):
        gap_m = float(cum_m[cur] - cum_m[prev])
        gap_s = float(time_s[cur] - time_s[prev])
        if (
            same_pass[prev] == same_pass[cur]
            and gap_m <= BAD_SPOT_MAX_SAMPLE_GAP_M
            and gap_s <= BAD_SPOT_MAX_SAMPLE_GAP_S
        ):
            current.append(int(cur))
        else:
            runs.append(current)
            current = [int(cur)]
    runs.append(current)
    return runs


def _split_by_span(run: list[int], lat: np.ndarray, lon: np.ndarray) -> list[list[int]]:
    """Cut a run that wanders further than one oval should cover."""
    pieces: list[list[int]] = []
    current: list[int] = []
    for i in run:
        trial = current + [i]
        if len(trial) > 1 and _span_km(lat[trial], lon[trial]) > BAD_SPOT_CONSEC_MAX_SPAN_KM:
            pieces.append(current)
            current = [i]
        else:
            current = trial
    if current:
        pieces.append(current)
    return pieces


def _bbox_of(idx: list[int], lat: np.ndarray, lon: np.ndarray) -> tuple[float, float, float, float]:
    sl = lat[idx]
    so = lon[idx]
    return float(sl.min()), float(sl.max()), float(so.min()), float(so.max())


def _bbox_gap_km(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    """Distance between two lat/lon boxes. Zero when they touch or overlap."""
    lat_gap = max(0.0, b[0] - a[1], a[0] - b[1])
    lon_gap = max(0.0, b[2] - a[3], a[2] - b[3])
    lat0 = (a[0] + a[1] + b[0] + b[1]) / 4.0
    return float(math.hypot(lat_gap * 111.32, lon_gap * _km_per_deg_lon(lat0)))


def _union_span_km(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    lat_min, lat_max = min(a[0], b[0]), max(a[1], b[1])
    lon_min, lon_max = min(a[2], b[2]), max(a[3], b[3])
    lat0 = (lat_min + lat_max) / 2.0
    return max((lat_max - lat_min) * 111.32, (lon_max - lon_min) * _km_per_deg_lon(lat0))


def _merge_circle_groups(
    pieces: list[list[int]],
    lengths: list[float],
    lat: np.ndarray,
    lon: np.ndarray,
) -> list[dict]:
    """Join poor runs that sit within BAD_SPOT_MERGE_GAP_KM when the union still fits one circle.

    Runs farther apart stay as their own circles, including several circles inside
    one 10 km neighbourhood.
    """
    groups = [
        {"idx": list(idx), "length": length, "bbox": _bbox_of(idx, lat, lon)}
        for idx, length in zip(pieces, lengths)
        if idx
    ]
    changed = True
    while changed and len(groups) > 1:
        changed = False
        best: tuple[tuple[float, float], int, int] | None = None
        for i in range(len(groups)):
            for j in range(i + 1, len(groups)):
                gap = _bbox_gap_km(groups[i]["bbox"], groups[j]["bbox"])
                if gap > BAD_SPOT_MERGE_GAP_KM:
                    continue
                span = _union_span_km(groups[i]["bbox"], groups[j]["bbox"])
                if span > BAD_SPOT_GROUP_SPAN_KM:
                    continue
                key = (gap, span)
                if best is None or key < best[0]:
                    best = (key, i, j)
        if best is None:
            break
        _, i, j = best
        groups[i]["idx"].extend(groups[j]["idx"])
        groups[i]["length"] += groups[j]["length"]
        groups[i]["bbox"] = (
            min(groups[i]["bbox"][0], groups[j]["bbox"][0]),
            max(groups[i]["bbox"][1], groups[j]["bbox"][1]),
            min(groups[i]["bbox"][2], groups[j]["bbox"][2]),
            max(groups[i]["bbox"][3], groups[j]["bbox"][3]),
        )
        groups.pop(j)
        changed = True
    return groups


def _consecutive_bad_spots(
    track: pd.DataFrame,
    column: str,
    threshold: float,
    detect_column: str | None = None,
    min_length_m: float = BAD_SPOT_MIN_CONSEC_M,
) -> tuple[pd.DataFrame, np.ndarray]:
    """Rule 2: a run of consecutive poor road of at least min_length_m.

    Returns (spots, used-row mask).
    """
    detect_column = detect_column or column
    n = len(track)
    used = np.zeros(n, dtype=bool)
    if n == 0:
        return pd.DataFrame(), used
    lat = track["Latitude"].to_numpy(dtype=float)
    lon = track["Longitude"].to_numpy(dtype=float)
    time_s = pd.to_datetime(track["Time"]).to_numpy(dtype="datetime64[ns]").astype("int64") / 1e9
    same_pass = (
        track["Source"].to_numpy()
        if "Source" in track.columns
        else np.zeros(n, dtype=np.int64)
    )
    poor = _poor_mask(
        track[detect_column], threshold, strict_less=bool(BAD_SPOT_RULES[column].get("strict_less"))
    ).to_numpy()
    step_m = np.zeros(n)
    step_m[1:] = _equirect_km(lat[:-1], lon[:-1], lat[1:], lon[1:]) * 1000.0
    cum_m = np.cumsum(step_m)

    rows: list[dict] = []
    longest_short_m = 0.0
    pieces: list[list[int]] = []
    piece_lengths: list[float] = []
    for run in _poor_runs(np.flatnonzero(poor), same_pass, cum_m, time_s):
        for piece in _split_by_span(run, lat, lon):
            if not piece:
                continue
            # A lone poor sample still stands for the 50 m bin it was judged on.
            length_m = max(float(cum_m[piece[-1]] - cum_m[piece[0]]), MAP_BIN_M)
            if length_m < min_length_m:
                longest_short_m = max(longest_short_m, length_m)
                continue
            pieces.append(piece)
            piece_lengths.append(length_m)
            used[piece] = True
    for group in _merge_circle_groups(pieces, piece_lengths, lat, lon):
        chunk = track.iloc[sorted(group["idx"])]
        rows.append(
            _spot_from_points(
                chunk,
                column,
                threshold,
                kind="consecutive",
                length_m=float(group["length"]),
                area_km2=_hull_area_km2(chunk["Longitude"].to_numpy(), chunk["Latitude"].to_numpy()),
            )
        )
    spots = pd.DataFrame(rows)
    if not spots.empty:
        spots = spots.sort_values("length_m", ascending=False).reset_index(drop=True)
    spots.attrs["longest_short_m"] = longest_short_m
    return spots, used


def _discrete_area_bad_spots(leftover: pd.DataFrame, column: str, threshold: float) -> pd.DataFrame:
    """Rule 4: leftover poor bins covering >= 1 km² that never joined a poor run.

    Cluster on a ~1 km grid and only keep compact patches (not a thin road network).
    """
    if leftover.empty:
        return pd.DataFrame()
    g = leftover.reset_index(drop=True)
    latb = np.floor(g["Latitude"].to_numpy(dtype=float) / BAD_SPOT_GRID_DEG).astype(int)
    lonb = np.floor(g["Longitude"].to_numpy(dtype=float) / BAD_SPOT_GRID_DEG).astype(int)
    keys = list(zip(latb.tolist(), lonb.tolist()))
    cell_idx: dict[tuple[int, int], list[int]] = {}
    for i, cell in enumerate(keys):
        cell_idx.setdefault(cell, []).append(i)
    cells = set(cell_idx)
    neigh = ((1, 0), (-1, 0), (0, 1), (0, -1))
    seen: set[tuple[int, int]] = set()
    rows = []
    for cell in cells:
        if cell in seen:
            continue
        queue = deque([cell])
        seen.add(cell)
        group = [cell]
        while queue:
            y, x = queue.popleft()
            for dy, dx in neigh:
                nb = (y + dy, x + dx)
                if nb in cells and nb not in seen:
                    trial = group + [nb]
                    idx_trial = [i for c in trial for i in cell_idx[c]]
                    trial_part = g.iloc[idx_trial]
                    span_lat_km = (
                        float(trial_part["Latitude"].max()) - float(trial_part["Latitude"].min())
                    ) * 111.32
                    span_lon_km = (
                        float(trial_part["Longitude"].max()) - float(trial_part["Longitude"].min())
                    ) * _km_per_deg_lon(float(trial_part["Latitude"].mean()))
                    if max(span_lat_km, span_lon_km) > BAD_SPOT_DISCRETE_MAX_SPAN_KM:
                        continue
                    seen.add(nb)
                    queue.append(nb)
                    group.append(nb)
        idx = [i for c in group for i in cell_idx[c]]
        part = g.iloc[idx]
        if len(part) < BAD_SPOT_DISCRETE_MIN_POOR:
            continue
        area = _hull_area_km2(part["Longitude"].to_numpy(), part["Latitude"].to_numpy())
        span_lat_km = (float(part["Latitude"].max()) - float(part["Latitude"].min())) * 111.32
        span_lon_km = (float(part["Longitude"].max()) - float(part["Longitude"].min())) * _km_per_deg_lon(float(part["Latitude"].mean()))
        bbox_area = span_lat_km * span_lon_km
        # A thin road of leftover holes is not a 1 km² area.
        if min(span_lat_km, span_lon_km) < 0.45:
            continue
        if area < BAD_SPOT_DISCRETE_AREA_KM2 and bbox_area < BAD_SPOT_DISCRETE_AREA_KM2:
            continue
        rows.append(
            _spot_from_points(
                part,
                column,
                threshold,
                kind="discrete",
                length_m=np.nan,
                area_km2=max(area, bbox_area),
            )
        )
    spots = pd.DataFrame(rows)
    if spots.empty:
        return spots
    return spots.sort_values("area_km2", ascending=False).head(BAD_SPOT_MAX_DISCRETE).reset_index(drop=True)


def _spot_extra_stats(df: pd.DataFrame, spot: pd.Series) -> dict:
    box = df[
        df["Latitude"].between(spot["lat_min"], spot["lat_max"])
        & df["Longitude"].between(spot["lon_min"], spot["lon_max"])
    ]
    top_cell = pd.NA
    if "Cell Id" in box.columns and box["Cell Id"].notna().any():
        top_cell = int(box["Cell Id"].dropna().astype("int64").value_counts().idxmax())
    return {
        "mean_rsrp": float(box["RSRP"].mean()) if "RSRP" in box and box["RSRP"].notna().any() else np.nan,
        "mean_rsrq": float(box["RSRQ"].mean()) if "RSRQ" in box and box["RSRQ"].notna().any() else np.nan,
        "mean_sinr": float(box["SINR"].mean()) if "SINR" in box and box["SINR"].notna().any() else np.nan,
        "top_cell": top_cell,
    }


def find_bad_spots(df: pd.DataFrame, column: str, max_spots: int | None = None) -> pd.DataFrame:
    """Bad spots on the same best-server all-band track used by the first coverage map."""
    rule = BAD_SPOT_RULES[column]
    threshold = rule["threshold"]
    if max_spots is None:
        max_spots = int(rule.get("max_consecutive", BAD_SPOT_MAX_CONSEC)) + BAD_SPOT_MAX_DISCRETE
    track = _prepare_track(df, column)
    binned_col = f"{column}__binned"
    track[binned_col] = binned_kpi(track, column)
    consec, used = _consecutive_bad_spots(
        track,
        column,
        threshold,
        detect_column=binned_col,
        min_length_m=float(rule.get("min_consecutive_m", BAD_SPOT_MIN_CONSEC_M)),
    )
    longest_short_m = float(consec.attrs.get("longest_short_m", 0.0))
    total_consec = len(consec)
    max_consec = int(rule.get("max_consecutive", BAD_SPOT_MAX_CONSEC))
    if total_consec > max_consec:
        consec = consec.head(max_consec).copy()
    strict_less = bool(rule.get("strict_less"))
    leftover = track.iloc[np.flatnonzero(_poor_mask(track[binned_col], threshold, strict_less=strict_less).to_numpy() & ~used)]
    discrete = _discrete_area_bad_spots(leftover, column, threshold)
    if not discrete.empty:
        dense = [
            _map_poor_share(track, row, column, threshold) >= BAD_SPOT_DISCRETE_MIN_SHARE
            for _, row in discrete.iterrows()
        ]
        discrete = discrete.loc[dense].reset_index(drop=True)
    total_discrete = len(discrete)
    spots = pd.concat([consec, discrete], ignore_index=True)
    if spots.empty:
        return spots
    if len(spots) > max_spots:
        spots = spots.head(max_spots).copy()
    extras = [_spot_extra_stats(df, row) for _, row in spots.iterrows()]
    spots = spots.copy()
    spots["spot"] = np.arange(1, len(spots) + 1)
    spots["kpi"] = column
    spots["threshold"] = threshold
    spots["unit"] = rule["unit"]
    spots["mean_rsrp"] = [e["mean_rsrp"] for e in extras]
    spots["mean_rsrq"] = [e["mean_rsrq"] for e in extras]
    spots["mean_sinr"] = [e["mean_sinr"] for e in extras]
    spots["top_cell"] = [e["top_cell"] for e in extras]
    # Reported, not filtered on: every sample below the threshold is painted on top
    # of the good ones, so a stretch that is poor on one pass out of five is still
    # visible on the map and must still be circled.
    spots["map_poor_share"] = [_map_poor_share(track, row, column, threshold) for _, row in spots.iterrows()]
    spots.attrs["total_consecutive"] = total_consec
    spots.attrs["total_discrete"] = total_discrete
    spots.attrs["longest_short_m"] = longest_short_m
    return spots


def _map_poor_share(df: pd.DataFrame, spot: pd.Series, column: str, threshold: float) -> float:
    """Share of all-band best-server samples in the (padded) spot box that are poor."""
    lat_min = float(spot["lat_min"])
    lat_max = float(spot["lat_max"])
    lon_min = float(spot["lon_min"])
    lon_max = float(spot["lon_max"])
    dlat = max(lat_max - lat_min, 1e-5)
    dlon = max(lon_max - lon_min, 1e-5)
    box = df[
        df["Latitude"].between(lat_min - 0.15 * dlat, lat_max + 0.15 * dlat)
        & df["Longitude"].between(lon_min - 0.15 * dlon, lon_max + 0.15 * dlon)
    ]
    vals = pd.to_numeric(box[column], errors="coerce")
    mask = _poor_mask(vals, threshold, strict_less=bool(BAD_SPOT_RULES[column].get("strict_less")))
    n = int(vals.notna().sum())
    if n == 0:
        return 0.0
    return float(mask.sum() / n)


def _ellipse_from_points(
    lons: np.ndarray,
    lats: np.ndarray,
    min_km: float = 1.6,
    max_km: float = 6.0,
    length_km: float | None = None,
) -> tuple[float, float, float, float, float]:
    """Oval around a consecutive poor stretch; sized from the actual span/length."""
    lon0 = float(np.mean(lons))
    lat0 = float(np.mean(lats))
    km_lon = _km_per_deg_lon(lat0)
    span_lon_km = float(np.max(lons) - np.min(lons)) * km_lon if len(lons) else 0.0
    span_lat_km = float(np.max(lats) - np.min(lats)) * 111.32 if len(lats) else 0.0
    angle = 0.0
    if len(lons) >= 8:
        x = (lons - lon0) * km_lon
        y = (lats - lat0) * 111.32
        cov = np.cov(np.vstack([x, y]))
        eigvals, eigvecs = np.linalg.eigh(cov)
        order = np.argsort(eigvals)[::-1]
        eigvecs = eigvecs[:, order]
        angle = float(np.degrees(np.arctan2(eigvecs[1, 0], eigvecs[0, 0])))
    width_km = max(span_lon_km * 1.4, min_km)
    height_km = max(span_lat_km * 1.4, min_km * 0.7)
    if length_km:
        major = max(width_km, height_km, float(length_km) * 1.15)
        if width_km >= height_km:
            width_km = major
        else:
            height_km = major
    width_km = min(width_km, max_km)
    height_km = min(height_km, max_km)
    return lon0, lat0, width_km / km_lon, height_km / 111.32, angle


def _circle_from_bbox(lon_min: float, lon_max: float, lat_min: float, lat_max: float, min_diameter_km: float = 2.8, max_diameter_km: float = 5.5) -> tuple[float, float, float, float]:
    """Large circle covering a discrete poor area (>= 1 km²)."""
    lat0 = (lat_min + lat_max) / 2.0
    lon0 = (lon_min + lon_max) / 2.0
    span_lon_km = (lon_max - lon_min) * _km_per_deg_lon(lat0)
    span_lat_km = (lat_max - lat_min) * 111.32
    diameter_km = float(np.clip(max(span_lon_km, span_lat_km) * 1.25, min_diameter_km, max_diameter_km))
    return lon0, lat0, diameter_km / _km_per_deg_lon(lat0), diameter_km / 111.32


def add_map_scale_bar(ax, km: float | None = None, corner: str = "upper left", boxed: bool = True) -> None:
    """Scale bar on a map. The length is a round fraction of the view.

    boxed is off on a bad-spot map so the length text has no filled background.
    """
    xlim = ax.get_xlim()
    ylim = ax.get_ylim()
    if xlim[1] <= xlim[0] or ylim[1] <= ylim[0]:
        return
    lat_mid = (ylim[0] + ylim[1]) / 2.0
    span_km = (xlim[1] - xlim[0]) * _km_per_deg_lon(lat_mid)
    if km is None:
        raw = span_km / 5.0
        km = 0.2
        for choice in (0.2, 0.5, 1.0, 2.0, 5.0, 10.0, 20.0, 40.0):
            if choice <= raw * 1.35:
                km = choice
    if corner == "lower left":
        lat = ylim[0] + 0.08 * (ylim[1] - ylim[0])
    else:
        lat = ylim[1] - 0.09 * (ylim[1] - ylim[0])
    lon0 = xlim[0] + 0.045 * (xlim[1] - xlim[0])
    width = km / _km_per_deg_lon(lat)
    tick = 0.012 * (ylim[1] - ylim[0])
    ax.plot([lon0, lon0 + width], [lat, lat], color="white", lw=5.5, solid_capstyle="butt", zorder=40, clip_on=True)
    ax.plot([lon0, lon0 + width], [lat, lat], color="#1B1B1B", lw=2.2, solid_capstyle="butt", zorder=41, clip_on=True)
    ax.plot([lon0, lon0], [lat - tick, lat + tick], color="#1B1B1B", lw=1.3, zorder=41, clip_on=True)
    ax.plot([lon0 + width, lon0 + width], [lat - tick, lat + tick], color="#1B1B1B", lw=1.3, zorder=41, clip_on=True)
    label = f"{km:g} km" if km >= 1 else f"{int(round(km * 1000))} m"
    style = {
        "ha": "center",
        "va": "bottom",
        "fontsize": 8,
        "color": "#F7FEFF" if not boxed else "#1B1B1B",
        "fontweight": "bold",
        "zorder": 42,
        "clip_on": True,
    }
    if boxed:
        style["bbox"] = {"boxstyle": "round,pad=0.15", "fc": "white", "ec": "#1B1B1B", "lw": 0.6, "alpha": 0.94}
    else:
        style["path_effects"] = _LABEL_STROKE
    ax.text(lon0 + width / 2.0, lat + tick, label, **style)


def _spot_outline(spot: pd.Series) -> tuple[float, float, float, float, float, tuple]:
    """Oval/circle that encloses the whole poor stretch, including red tails."""
    kind = str(spot.get("kind", "consecutive"))
    style = (0, (1.0, 1.4)) if kind == "discrete" else (0, (3.0, 2.0))
    lat0 = float(spot["outline_lat"]) if "outline_lat" in spot.index and pd.notna(spot.get("outline_lat")) else (float(spot["lat_min"]) + float(spot["lat_max"])) / 2.0
    lon0 = float(spot["outline_lon"]) if "outline_lon" in spot.index and pd.notna(spot.get("outline_lon")) else (float(spot["lon_min"]) + float(spot["lon_max"])) / 2.0
    km_lon = _km_per_deg_lon(lat0)
    if "outline_width_km" in spot.index and pd.notna(spot.get("outline_width_km")):
        width_km = float(spot["outline_width_km"])
        height_km = float(spot["outline_height_km"])
        angle = float(spot.get("outline_angle", 0.0) or 0.0)
        return lon0, lat0, width_km / km_lon, height_km / 111.32, angle, style
    span_lon_km = max((float(spot["lon_max"]) - float(spot["lon_min"])) * km_lon, 0.2)
    span_lat_km = max((float(spot["lat_max"]) - float(spot["lat_min"])) * 111.32, 0.2)
    cover = float(np.sqrt(2.0)) * BAD_SPOT_OUTLINE_PAD / 1.15
    width_km = max(span_lon_km * cover, 0.45)
    height_km = max(span_lat_km * cover, 0.38)
    if kind == "discrete":
        diameter_km = min(max(width_km, height_km, 1.15), 4.5)
        return lon0, lat0, diameter_km / km_lon, diameter_km / 111.32, 0.0, style
    return lon0, lat0, width_km / km_lon, height_km / 111.32, 0.0, style


def overview_ring_style(spot_count: int) -> tuple[float, float, bool]:
    """Ring floor (km), line width and halo flag for this many spots on one map."""
    for limit, min_km, lw, halo in BAD_SPOT_OVERVIEW_TIERS:
        if spot_count <= limit:
            return min_km, lw, halo
    return BAD_SPOT_OVERVIEW_TIERS[-1][1:]


def _add_spot_outline(
    ax,
    spot: pd.Series,
    *,
    overview: bool = False,
    style_tier: tuple[float, float, bool] = (BAD_SPOT_OVERVIEW_MIN_KM, BAD_SPOT_OVERVIEW_LW, True),
) -> None:
    cx, cy, w, h, angle, style = _spot_outline(spot)
    lat0 = cy
    km_lon = _km_per_deg_lon(lat0)
    if overview:
        min_diameter_km, lw, halo = style_tier
        width_km = abs(w) * km_lon
        height_km = abs(h) * 111.32
        diameter_km = float(np.clip(max(width_km, height_km, min_diameter_km), min_diameter_km, BAD_SPOT_OVERVIEW_MAX_KM))
        w = diameter_km / km_lon
        h = diameter_km / 111.32
        angle = 0.0
        if halo:
            ax.add_patch(
                Ellipse(
                    (cx, cy),
                    width=w,
                    height=h,
                    angle=angle,
                    fill=False,
                    edgecolor="#FFFFFF",
                    linestyle="solid",
                    linewidth=lw + 2.0,
                    zorder=19,
                )
            )
    else:
        lw = 1.0
    # Black under the orange stroke so the circle shows on both dark water and pale roofs.
    ax.add_patch(
        Ellipse(
            (cx, cy),
            width=w,
            height=h,
            angle=angle,
            fill=False,
            edgecolor="#1B1B1B",
            linestyle=style,
            linewidth=lw + 1.6,
            zorder=19,
        )
    )
    ax.add_patch(
        Ellipse(
            (cx, cy),
            width=w,
            height=h,
            angle=angle,
            fill=False,
            edgecolor="#FF6600",
            linestyle=style,
            linewidth=lw,
            zorder=20,
        )
    )


def _zoom_extent(spot: pd.Series) -> tuple[float, float, float, float]:
    lat0 = float(spot["outline_lat"]) if "outline_lat" in getattr(spot, "index", []) and pd.notna(spot.get("outline_lat")) else float(spot["lat"])
    lon0 = float(spot["outline_lon"]) if "outline_lon" in getattr(spot, "index", []) and pd.notna(spot.get("outline_lon")) else float(spot["lon"])
    km_lon = _km_per_deg_lon(lat0)
    span_lat_km = max((float(spot["lat_max"]) - float(spot["lat_min"])) * 111.32, 0.2)
    span_lon_km = max((float(spot["lon_max"]) - float(spot["lon_min"])) * km_lon, 0.2)
    if "outline_width_km" in getattr(spot, "index", []) and pd.notna(spot.get("outline_width_km")):
        span_lon_km = max(span_lon_km, float(spot["outline_width_km"]))
        span_lat_km = max(span_lat_km, float(spot["outline_height_km"]))
    half = max(BAD_SPOT_ZOOM_MIN_KM, span_lat_km, span_lon_km) / 2.0 + BAD_SPOT_ZOOM_PAD_KM
    return (
        lon0 - half / km_lon,
        lon0 + half / km_lon,
        lat0 - half / 111.32,
        lat0 + half / 111.32,
    )


def plot_bad_spot_map(
    df: pd.DataFrame,
    path: Path,
    column: str,
    class_series: pd.Series,
    labels: list[str],
    colors: list[str],
    spots: pd.DataFrame,
    title: str,
    legend_title: str,
    max_points: int = MAP_MAX_POINTS,
    extent: tuple[float, float, float, float] | None = None,
    basemap: bool = True,
    shapes: bool = False,
    outlines: bool = False,
    basemap_style: str = "satellite",
    route_casing: bool = True,
    legend_inside: bool = False,
    legend_counts: list[int] | None = None,
) -> Path:
    """First all-band coverage map.

    A .svg path writes a pure vector file: no basemap, transparent background,
    and text kept as text. shapes=True leaves the circles out of the picture and
    records one geometry per circle so the workbook can place a separate shape
    for each of them. outlines=True paints the circles into the picture. The
    combined view uses neither: its circles live only on the individual bad-spot maps.
    """
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    vector = path.suffix.lower() == ".svg"
    if vector:
        basemap = False
    geo = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    fig, ax = plt.subplots(figsize=(8.0, 9.2))
    _draw_coverage_points(
        ax, geo, class_series, labels, colors, max_points=max_points,
        bin_m=MAP_BIN_M, column=column, route_casing=route_casing,
    )
    if extent is not None:
        ax.set_xlim(extent[0], extent[1])
        ax.set_ylim(extent[2], extent[3])
    elif not geo.empty:
        lon_min, lon_max, lat_min, lat_max = _coverage_extent_from_geo(geo)
        ax.set_xlim(lon_min, lon_max)
        ax.set_ylim(lat_min, lat_max)
    ax.set_aspect("equal", adjustable="box")
    if basemap and basemap_style == "boundary":
        _add_boundary_basemap(ax)
    elif basemap:
        _add_basemap(ax)
    if not vector:
        add_map_scale_bar(ax, corner="lower left" if legend_inside else "upper left")
    if legend_inside:
        counts = legend_counts if legend_counts is not None else [0] * len(labels)
        legend = _place_count_legend(ax, _count_legend_handles(labels, colors, counts), legend_title)
    else:
        handles = _coverage_legend_handles(labels, colors)
        handles.extend(
            [
                Line2D([0], [0], color="#FF6600", lw=2.2, linestyle=(0, (3, 2)), label=consecutive_label(column)),
                Line2D([0], [0], color="#FF6600", lw=2.2, linestyle=(0, (1, 1.4)), label="Discrete area ≥ 1 km²"),
            ]
        )
        legend = ax.legend(
            handles=handles,
            title=legend_title,
            loc="center left",
            bbox_to_anchor=(1.02, 0.5),
            frameon=True,
            fontsize=8,
            title_fontsize=10,
        )
    if spots is not None and not spots.empty and outlines and not shapes:
        for _, spot in spots.iterrows():
            _add_spot_outline(ax, spot, overview=False)
    if not legend_inside:
        ax.set_title(title)
    hide_map_axes(ax)
    fig.tight_layout()
    _save_map(fig, path, legends=[legend], ax=ax if shapes else None, spots=spots if shapes else None)
    return path


def _record_circle_pixels(
    fig, ax, path: Path, spots: pd.DataFrame, tight, axes=None, name_start: int = 1,
) -> None:
    """Store each circle's box in pixels of the PNG that savefig is about to write.

    axes, when set, is one axes per spot, in the same order. name_start keeps a
    single-spot map named Circle 7 when it is bad spot 7.
    """
    dpi = float(fig.dpi)
    save_dpi = float(MAP_DPI)
    circles = []
    for i, (_, spot) in enumerate(spots.iterrows()):
        n = name_start + i
        use_ax = axes[i] if axes is not None else ax
        cx, cy, w, h, _angle, _style = _spot_outline(spot)
        x0, y_lo = use_ax.transData.transform((cx - w / 2.0, cy - h / 2.0))
        x1, y_hi = use_ax.transData.transform((cx + w / 2.0, cy + h / 2.0))

        def to_saved_x(px: float) -> float:
            return (px / dpi - tight.x0) * save_dpi

        def to_saved_y_top(py: float) -> float:
            return (tight.y1 - py / dpi) * save_dpi

        left, right = to_saved_x(x0), to_saved_x(x1)
        top, bottom = to_saved_y_top(y_hi), to_saved_y_top(y_lo)
        circles.append({
            "name": f"Circle {n}",
            "x": float(min(left, right)),
            "y": float(min(top, bottom)),
            "w": float(abs(right - left)),
            "h": float(abs(bottom - top)),
            "dotted": str(spot.get("kind", "consecutive")) == "discrete",
        })
    MAP_CIRCLE_PIXELS[str(path.resolve())] = {
        "size": (tight.width * save_dpi, tight.height * save_dpi),
        "circles": circles,
    }


def _save_map(
    fig, path: Path, legends=(), ax=None, spots: pd.DataFrame | None = None, axes=None, name_start: int = 1,
) -> None:
    """PNG at MAP_DPI on white, or for .svg a transparent vector file with text kept as text.

    When spots is set, each circle is recorded against the PNG and is not drawn
    into it. The workbook turns those records into separate shapes. axes is one
    axes per spot for a grid of zoom panels.
    """
    pad = 0.15
    if path.suffix.lower() == ".svg":
        fig.patch.set_alpha(0.0)
        for axis in fig.axes:
            axis.patch.set_alpha(0.0)
        for legend in legends:
            legend.get_frame().set_facecolor("none")
        with plt.rc_context({"svg.fonttype": "none", "svg.hashsalt": "bogura"}):
            fig.savefig(path, format="svg", transparent=True, bbox_inches="tight", pad_inches=pad)
    else:
        bbox_arg: object = "tight"
        if spots is not None and not spots.empty and (ax is not None or axes is not None):
            fig.canvas.draw()
            tight = fig.get_tightbbox(fig.canvas.get_renderer()).padded(pad)
            _record_circle_pixels(fig, ax, path, spots, tight, axes=axes, name_start=name_start)
            bbox_arg = tight
            pad = 0.0
        fig.savefig(path, dpi=MAP_DPI, facecolor="white", bbox_inches=bbox_arg, pad_inches=pad)
    plt.close(fig)


def plot_bad_spot_zoom_grid(
    df: pd.DataFrame,
    spots: pd.DataFrame,
    path: Path,
    column: str,
    class_series: pd.Series,
    labels: list[str],
    colors: list[str],
    title: str,
    shapes: bool = False,
) -> Path | None:
    """Zoomed all-band coverage around each identified spot, with no numbers.

    A .svg path writes a pure vector file without the basemap, like plot_bad_spot_map.
    shapes=True leaves each circle out of the picture and records it so the workbook
    can place a separate shape on that panel.
    """
    if spots is None or spots.empty:
        return None
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    vector = path.suffix.lower() == ".svg"
    work = spots.head(BAD_SPOT_MAX_ZOOMS)
    n = len(work)
    cols = BAD_SPOT_ZOOM_COLS
    rows = max(1, (n + cols - 1) // cols)
    fig, axes = plt.subplots(rows, cols, figsize=(3.15 * cols, 3.35 * rows))
    axes_flat = np.atleast_1d(axes).ravel()
    geo = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    panel_axes = []
    for i, ax in enumerate(axes_flat):
        if i >= n:
            ax.set_axis_off()
            continue
        spot = work.iloc[i]
        zext = _zoom_extent(spot)
        local = geo[
            geo["Longitude"].between(zext[0], zext[1])
            & geo["Latitude"].between(zext[2], zext[3])
        ]
        _draw_coverage_points(
            ax,
            local,
            class_series.reindex(local.index),
            labels,
            colors,
            max_points=8000,
            poor_on_top=True,
            bin_m=MAP_BIN_M,
            column=column,
        )
        ax.set_xlim(zext[0], zext[1])
        ax.set_ylim(zext[2], zext[3])
        ax.set_aspect("equal", adjustable="box")
        if not vector:
            _add_basemap(ax)
            add_map_scale_bar(ax)
        if not vector:
            draw_bad_spot_sector_pies(ax, spot)
        if shapes:
            panel_axes.append(ax)
        else:
            _add_spot_outline(ax, spot)
        hide_map_axes(ax)
        kind = str(spot.get("kind", "consecutive"))
        if kind == "discrete":
            caption = f"Discrete {float(spot['area_km2']):.1f} km²"
        else:
            word = "Poor run" if min_consecutive_m(column) <= 0 else "Consecutive"
            caption = f"{word} {float(spot['length_m']):.0f} m"
        ax.set_title(caption, fontsize=8)
    fig.suptitle(title, fontsize=11, fontweight="bold")
    fig.tight_layout()
    if shapes and panel_axes:
        _save_map(fig, path, axes=panel_axes, spots=work.reset_index(drop=True))
    else:
        _save_map(fig, path)
    return path


def _extent_around(extent: tuple[float, float, float, float], lat: float, lon: float, pad_km: float = 0.95) -> tuple[float, float, float, float]:
    """Widen a view so a site centre and its sector arms stay inside the frame."""
    pad_lon = pad_km / _km_per_deg_lon(lat)
    pad_lat = pad_km / 111.32
    return (
        min(extent[0], lon - pad_lon),
        max(extent[1], lon + pad_lon),
        min(extent[2], lat - pad_lat),
        max(extent[3], lat + pad_lat),
    )


def _extent_with_nearest_site(
    extent: tuple[float, float, float, float],
    lat: float,
    lon: float,
    overrides: dict | None = None,
    names: list[str] | None = None,
) -> tuple[float, float, float, float]:
    """Keep at least the nearest site on the map, plus any sector the plan moves."""
    nearest = _sectors_near(lat, lon, 5.0)
    if nearest:
        best = min(nearest, key=lambda rec: float(_equirect_km(rec["lat"], rec["lon"], lat, lon)))
        extent = _extent_around(extent, float(best["lat"]), float(best["lon"]))
    override_names = {str(name).strip() for name in (overrides or {})}
    wanted = [str(name).strip() for name in (names or []) if str(name).strip()]
    wanted.extend(override_names)
    if not wanted:
        return extent
    sectors = load_site_sectors()
    if sectors.empty:
        return extent
    sector_name = sectors["Sector"].astype(str).str.strip()
    site_name = sectors["SiteName"].astype(str).str.strip()
    for name in wanted:
        hit = sectors[sector_name == name]
        if hit.empty:
            hit = sectors[site_name == name]
        if hit.empty:
            continue
        slat, slon = float(hit.iloc[0]["lat"]), float(hit.iloc[0]["lon"])
        if float(_equirect_km(slat, slon, lat, lon)) > 5.0:
            continue
        # A turned sector grows a dashed arm about twice the cyan arm. Keep that tip in frame.
        pad = 2.3 if name in override_names else 0.95
        extent = _extent_around(extent, slat, slon, pad_km=pad)
    return extent


def plot_bad_spot_serial_maps(
    df: pd.DataFrame,
    spots: pd.DataFrame,
    directory: Path,
    stem: str,
    column: str,
    class_series: pd.Series,
    labels: list[str],
    colors: list[str],
    *,
    basemap_style: str = "satellite",
    route_casing: bool = True,
    mode: str = "",
    only: set[int] | None = None,
) -> list[Path]:
    """One large map per bad spot, longest first.

    The circle is recorded and not painted, so the workbook can place one movable
    shape on that map. Sector pies and distance lines stay in the picture. The
    combined view does not use this function.
    """
    if spots is None or spots.empty:
        return []
    directory.mkdir(parents=True, exist_ok=True)
    work = spots.head(BAD_SPOT_MAX_ZOOMS)
    geo = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    paths: list[Path] = []
    total = len(work)
    for i, (_, spot) in enumerate(work.iterrows(), start=1):
        path = directory / f"{stem}_{i:02d}.png"
        if only is not None and i not in only:
            continue
        print(f"  {stem} bad spot {i}/{total}", flush=True)
        fig, ax = plt.subplots(figsize=(9.4, 8.6))
        zext = _zoom_extent(spot)
        # Room for every sector arm when the site sits near the edge of the spot.
        arm_km = 0.9
        mid_lat = (zext[2] + zext[3]) / 2.0
        arm_lon = arm_km / _km_per_deg_lon(mid_lat)
        arm_lat = arm_km / 111.32
        zext = (zext[0] - arm_lon, zext[1] + arm_lon, zext[2] - arm_lat, zext[3] + arm_lat)
        plan = bad_spot_advice(df, column, spot, mode=mode)
        proposal = plan.get("draw")
        overshoot = plan.get("overshoot_marks") or {}
        if proposal:
            zext = _extent_with_proposal(zext, proposal)
        plan_names = []
        for row in plan.get("actions") or []:
            plan_names.append(str(row.get("site") or ""))
            plan_names.append(str(row.get("sector") or ""))
        zext = _extent_with_nearest_site(
            zext,
            float(spot["lat"]),
            float(spot["lon"]),
            plan.get("azimuth_overrides") or {},
            plan_names,
        )
        local = geo[
            geo["Longitude"].between(zext[0], zext[1])
            & geo["Latitude"].between(zext[2], zext[3])
        ]
        _draw_coverage_points(
            ax,
            local,
            class_series.reindex(local.index),
            labels,
            colors,
            max_points=8000,
            poor_on_top=True,
            bin_m=MAP_BIN_M,
            column=column,
            route_casing=route_casing,
        )
        ax.set_xlim(zext[0], zext[1])
        ax.set_ylim(zext[2], zext[3])
        ax.set_aspect("equal", adjustable="box")
        kind = str(spot.get("kind", "consecutive"))
        blanks = _blank_segments(df, spot)
        blank_m = sum(gap["length_m"] for gap in blanks)
        if kind == "discrete":
            caption = f"Discrete {float(spot['area_km2']):.1f} km²"
        else:
            word = "Poor run" if min_consecutive_m(column) <= 0 else "Consecutive"
            caption = f"{word} {float(spot['length_m']):.0f} m"
        if blank_m >= 80:
            caption = f"{caption}  +  blank {blank_m:,.0f} m"
        ax.set_title(caption, fontsize=14, fontweight="bold")
        hide_map_axes(ax)
        # Lock the axes size before any label is measured in screen pixels.
        fig.tight_layout()
        fig.canvas.draw()
        if basemap_style == "boundary":
            _add_boundary_basemap(ax)
        else:
            _add_basemap(ax)
        draw_bad_spot_sector_pies(
            ax, spot, local, column, proposal, overshoot, plan.get("azimuth_overrides") or {},
        )
        _draw_proposal(ax, proposal)
        for gap in blanks:
            ax.plot(
                [gap["lon1"], gap["lon2"]],
                [gap["lat1"], gap["lat2"]],
                color="#111111",
                linewidth=3.2,
                linestyle=(0, (1.2, 1.4)),
                solid_capstyle="round",
                zorder=24,
            )
            ax.plot(
                [gap["lon1"], gap["lon2"]],
                [gap["lat1"], gap["lat2"]],
                color="#FFFFFF",
                linewidth=1.5,
                linestyle=(0, (1.2, 1.4)),
                solid_capstyle="round",
                zorder=25,
            )
        add_map_scale_bar(ax, boxed=False)
        one = work.iloc[[i - 1]].reset_index(drop=True)
        _save_map(fig, path, ax=ax, spots=one, name_start=i)
        paths.append(path)
    return paths


def plot_rsrp_bad_spot_map(
    df: pd.DataFrame,
    spots: pd.DataFrame,
    path: Path,
    *,
    extent: tuple[float, float, float, float] | None = None,
    title: str = "Combined view — RSRP",
    shapes: bool = False,
    outlines: bool = False,
    legend_title: str = "Idle by RSRP",
) -> Path:
    """Satellite map coloured by the four-bin RSRP legend, counts on the left.

    The basemap is the same Esri imagery as the RSRP sheet. The coloured trace
    is the KPI, with the same white casing the coverage maps use.
    """
    counts = _legend_sample_counts(df, "RSRP", rsrp_map_class, MAP_RSRP_LABELS)
    return plot_bad_spot_map(
        df,
        path,
        column="RSRP",
        class_series=rsrp_map_class(df["RSRP"]),
        labels=MAP_RSRP_LABELS,
        colors=MAP_RSRP_COLORS,
        spots=spots,
        title=title,
        legend_title=legend_title,
        extent=extent,
        shapes=shapes,
        outlines=outlines,
        basemap_style="satellite",
        route_casing=True,
        legend_inside=True,
        legend_counts=counts,
    )


def export_rsrp_map_svg(df: pd.DataFrame, path: Path | None = None) -> Path:
    """Vector copy of the RSRP sheet's combined Idle map, from the same track and spots."""
    track = best_server(df)
    spots = find_bad_spots(track, "RSRP")
    out = path or (PLOTS_DIR / "svg" / "rsrp_combined_idle.svg")
    out.parent.mkdir(parents=True, exist_ok=True)
    return plot_rsrp_bad_spot_map(track, spots, out, extent=coverage_extent(track))


def plot_rsrq_bad_spot_map(
    df: pd.DataFrame,
    spots: pd.DataFrame,
    path: Path,
    *,
    extent: tuple[float, float, float, float] | None = None,
    title: str = "Combined view — RSRQ",
    shapes: bool = False,
    outlines: bool = False,
    legend_title: str = "Idle by RSRQ",
) -> Path:
    """Satellite map coloured by the three-bin RSRQ legend, counts on the left.

    The basemap is the same Esri imagery as the RSRP sheet. The coloured trace
    is the KPI, with the same white casing the coverage maps use.
    """
    counts = _legend_sample_counts(df, "RSRQ", rsrq_map_class, MAP_RSRQ_LABELS)
    return plot_bad_spot_map(
        df,
        path,
        column="RSRQ",
        class_series=rsrq_map_class(df["RSRQ"]),
        labels=MAP_RSRQ_LABELS,
        colors=MAP_RSRQ_COLORS,
        spots=spots,
        title=title,
        legend_title=legend_title,
        extent=extent,
        shapes=shapes,
        outlines=outlines,
        basemap_style="satellite",
        route_casing=True,
        legend_inside=True,
        legend_counts=counts,
    )


def plot_sinr_bad_spot_map(
    df: pd.DataFrame,
    spots: pd.DataFrame,
    path: Path,
    *,
    extent: tuple[float, float, float, float] | None = None,
    title: str = "Combined view — SINR",
    shapes: bool = False,
    outlines: bool = False,
) -> Path:
    return plot_bad_spot_map(
        df,
        path,
        column="SINR",
        class_series=sinr_map_class(df["SINR"]),
        labels=MAP_SINR_LABELS,
        colors=MAP_SINR_COLORS,
        spots=spots,
        title=title,
        legend_title="SINR (dB)",
        extent=extent,
        shapes=shapes,
        outlines=outlines,
    )


def extract_7z(archive: Path, dest: Path) -> None:
    """Extract a 7z archive. Uses the 7z binary when it is installed, else py7zr."""
    from shutil import which

    dest.mkdir(parents=True, exist_ok=True)
    exe = which("7z") or which("7za")
    if exe:
        subprocess.run([exe, "x", f"-o{dest}", str(archive), "-y"], check=True)
        return
    import py7zr

    with py7zr.SevenZipFile(archive) as zf:
        zf.extractall(path=dest)


def _numbered_csvs(directory: Path) -> list[Path]:
    if not directory.exists():
        return []
    return sorted(directory.glob("*.csv"), key=lambda p: int(p.stem) if p.stem.isdigit() else p.stem)


def idle_csv_files() -> list[Path]:
    return _numbered_csvs(IDLE_DIR)


def active_csv_files() -> list[Path]:
    return _numbered_csvs(ACTIVE_DIR)


def extract_archives() -> None:
    EXTRACT_DIR.mkdir(parents=True, exist_ok=True)
    if not IDLE_ARCHIVE.exists():
        sys.exit(f"Missing idle archive: {IDLE_ARCHIVE.name}")
    print(f"Extracting {IDLE_ARCHIVE.name} -> {IDLE_DIR}")
    extract_7z(IDLE_ARCHIVE, EXTRACT_DIR)
    files = idle_csv_files()
    if not files:
        sys.exit(f"Extraction finished but no CSV files were found in {IDLE_DIR}")
    print(f"  {len(files)} CSV files")


def _cgi_eci(value: object) -> object:
    """Last number of a CGI such as 470-2-168962068, which is the E-UTRAN cell identity."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return pd.NA
    text = str(value).strip()
    if not text or text.lower() in {"nan", "<na>", "none"}:
        return pd.NA
    token = text.replace(" ", "").split("-")[-1]
    if not token.isdigit():
        return pd.NA
    return int(token)


def _parse_active_time(series: pd.Series) -> pd.Series:
    """Parse either a full timestamp or a clock with no date."""
    raw = series.astype("string").str.strip()
    clock = raw.str.match(r"^\d{1,2}:\d{2}:\d{2}\s*[AaPp][Mm]$", na=False)
    if float(clock.mean()) > 0.9:
        parsed = pd.to_datetime(raw.where(clock), format="%I:%M:%S %p", errors="coerce")
        return ACTIVE_CLOCK_ANCHOR + (parsed - parsed.dt.normalize())
    return pd.to_datetime(series, errors="coerce")


def earfcn_to_layer(earfcn: object) -> str:
    """Map DL EARFCN to L900 / L1800 / L2100 / L2600. Unknown EARFCNs return ''."""
    if pd.isna(earfcn):
        return ""
    try:
        value = int(earfcn)
    except (TypeError, ValueError):
        return ""
    for layer, ranges in LAYER_EARFCN_RANGES.items():
        for lo, hi in ranges:
            if lo <= value <= hi:
                return layer
    return ""


def filter_by_layer(df: pd.DataFrame, layer: str) -> pd.DataFrame:
    if "Layer" in df.columns:
        return df.loc[df["Layer"] == layer].copy()
    earfcn = pd.to_numeric(df["DL EARFCN"], errors="coerce")
    return df.loc[earfcn.map(earfcn_to_layer) == layer].copy()


def _sweep_group_ids(
    source: np.ndarray,
    time_s: np.ndarray,
    lat: np.ndarray,
    lon: np.ndarray,
    max_s: float = BEST_SERVER_SWEEP_S,
    max_m: float = BEST_SERVER_SWEEP_M,
) -> np.ndarray:
    """Split a pass into scanner sweeps: a new group every max_s seconds or max_m metres."""
    n = len(time_s)
    ids = np.empty(n, dtype=np.int64)
    if n == 0:
        return ids
    step_m = np.zeros(n)
    step_m[1:] = _equirect_km(lat[:-1], lon[:-1], lat[1:], lon[1:]) * 1000.0
    cum_m = np.cumsum(step_m)
    gid = 0
    t0 = time_s[0]
    d0 = cum_m[0]
    src0 = source[0]
    for i in range(n):
        if source[i] != src0 or time_s[i] - t0 > max_s or cum_m[i] - d0 > max_m:
            gid += 1
            t0 = time_s[i]
            d0 = cum_m[i]
            src0 = source[i]
        ids[i] = gid
    return ids


def strongest_fix(df: pd.DataFrame) -> pd.DataFrame:
    """One row per latitude, longitude and timestamp: the strongest RSRP.

    RSRQ and SINR stay on that same row. Any other row at that fix is dropped.
    """
    work = df.dropna(subset=["RSRP", "Time", "Latitude", "Longitude"]).copy()
    if work.empty:
        return work
    work = work.reset_index(drop=True)
    work["RSRP"] = pd.to_numeric(work["RSRP"], errors="coerce")
    work["Latitude"] = pd.to_numeric(work["Latitude"], errors="coerce")
    work["Longitude"] = pd.to_numeric(work["Longitude"], errors="coerce")
    work["Time"] = pd.to_datetime(work["Time"], errors="coerce")
    work = work.dropna(subset=["RSRP", "Latitude", "Longitude", "Time"])
    if work.empty:
        return work
    idx = work.groupby(["Time", "Latitude", "Longitude"], sort=False)["RSRP"].idxmax()
    return work.loc[idx].reset_index(drop=True)


def best_server(df: pd.DataFrame) -> pd.DataFrame:
    """All-band best server: strongest RSRP per scanner sweep.

    Rows that share a latitude, longitude and timestamp are reduced to the
    strongest RSRP first. The scanner still writes one row per band with its own
    millisecond timestamp, so the remaining rows are grouped into short sweep
    windows and the strongest band in that window is kept.
    """
    work = strongest_fix(df)
    if work.empty:
        return work
    sort_cols = ["Source", "Time"] if "Source" in work.columns else ["Time"]
    work = work.sort_values(sort_cols).reset_index(drop=True)
    source = (
        work["Source"].to_numpy()
        if "Source" in work.columns
        else np.zeros(len(work), dtype=np.int64)
    )
    time_s = work["Time"].to_numpy(dtype="datetime64[ns]").astype("int64") / 1e9
    work["_sweep"] = _sweep_group_ids(
        source,
        time_s,
        work["Latitude"].to_numpy(dtype=float),
        work["Longitude"].to_numpy(dtype=float),
    )
    idx = work.groupby("_sweep", sort=False)["RSRP"].idxmax()
    return work.loc[idx].drop(columns=["_sweep"]).reset_index(drop=True)


def strongest_location(df: pd.DataFrame) -> pd.DataFrame:
    """Active auto mode: one band at each latitude and longitude.

    Rows that share a latitude, longitude and Cell Id collapse to the strongest
    RSRP, so one cell keeps one band. A secondary carrier with no CGI has no
    Cell Id; those rows still compete, one per EARFCN at that location. The
    strongest RSRP remains. RSRQ and SINR stay on that row.
    """
    work = df.dropna(subset=["RSRP", "Latitude", "Longitude"]).copy()
    if work.empty:
        return work
    work = work.reset_index(drop=True)
    work["RSRP"] = pd.to_numeric(work["RSRP"], errors="coerce")
    work["Latitude"] = pd.to_numeric(work["Latitude"], errors="coerce")
    work["Longitude"] = pd.to_numeric(work["Longitude"], errors="coerce")
    work["Cell Id"] = pd.to_numeric(work["Cell Id"], errors="coerce")
    work = work.dropna(subset=["RSRP", "Latitude", "Longitude"])
    if work.empty:
        return work
    identified = work.loc[work["Cell Id"].notna()]
    anonymous = work.loc[work["Cell Id"].isna()]
    parts = []
    if not identified.empty:
        per_cell = identified.groupby(["Latitude", "Longitude", "Cell Id"], sort=False)["RSRP"].idxmax()
        parts.append(identified.loc[per_cell])
    if not anonymous.empty:
        keys = ["Latitude", "Longitude"]
        if "DL EARFCN" in anonymous.columns:
            keys.append("DL EARFCN")
        per_carrier = anonymous.groupby(keys, sort=False, dropna=False)["RSRP"].idxmax()
        parts.append(anonymous.loc[per_carrier])
    if not parts:
        return work.iloc[0:0].copy()
    work = pd.concat(parts, ignore_index=True)
    per_loc = work.groupby(["Latitude", "Longitude"], sort=False)["RSRP"].idxmax()
    return work.loc[per_loc].reset_index(drop=True)


def select_active_report(df: pd.DataFrame) -> pd.DataFrame:
    """Active rows used on the Excel and web reports.

    Home Operator is stored as Operator. Only Robi primary cells (Type M-PCell)
    are drawn. Other operators and secondary cells stay in the merged CSV.
    """
    if df is None or df.empty:
        return df
    if "Type" not in df.columns:
        sys.exit("Active log has no Type column, so M-PCell cannot be selected")
    work = df
    if "Operator" in work.columns:
        operator = work["Operator"].astype("string").str.strip()
        work = work.loc[operator.eq(ACTIVE_OPERATOR)]
    kind = work["Type"].astype("string").str.strip()
    return work.loc[kind.eq(ACTIVE_CELL_TYPE)].copy()


def best_carrier(df: pd.DataFrame) -> pd.DataFrame:
    """Active Mode: one Robi row per latitude and longitude.

    The UE is in auto mode, so one band is plotted at a location. That band is
    the strongest RSRP, and its RSRQ and SINR are the values that are plotted.
    """
    return strongest_location(df)


def coverage_extent(df: pd.DataFrame) -> tuple[float, float, float, float]:
    """Padded lon/lat box so all-band and per-band maps share the same view."""
    geo = df.dropna(subset=["Longitude", "Latitude"])
    lon_min, lon_max = float(geo["Longitude"].min()), float(geo["Longitude"].max())
    lat_min, lat_max = float(geo["Latitude"].min()), float(geo["Latitude"].max())
    pad_lon = 0.02 * (lon_max - lon_min)
    pad_lat = 0.02 * (lat_max - lat_min)
    return lon_min - pad_lon, lon_max + pad_lon, lat_min - pad_lat, lat_max + pad_lat


def earfcn_to_band(earfcn: float) -> str:
    if pd.isna(earfcn):
        return "Unknown"
    value = int(earfcn)
    for band, (lo, hi, label) in EARFCN_BANDS.items():
        if lo <= value <= hi:
            return f"{label} ({band})"
    return f"EARFCN {value}"


def load_idle_csv(path: Path) -> pd.DataFrame:
    print(f"Loading {path.name} ...")
    df = pd.read_csv(path)
    df = df.rename(columns=SCANNER_RENAME)
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        sys.exit(f"{path.name} is missing columns: {missing}")
    keep = [c for c in df.columns if c in set(COLUMNS) | {
        "Operator", "Technology Band", "eNB-Id", "Sector Id",
        "Bin Name", "Collection", "Campaign",
    }]
    df = df[keep].copy()
    df["Time"] = pd.to_datetime(df["Time"], errors="coerce")
    for col in ("RSRP", "RSRQ", "SINR", "Longitude", "Latitude", "DL EARFCN"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["Cell Id"] = pd.to_numeric(df["Cell Id"], errors="coerce").astype("Int64")
    df["DL EARFCN"] = df["DL EARFCN"].astype("Int64")
    df["Source"] = path.name
    df["Band"] = df["DL EARFCN"].map(earfcn_to_band)
    df["Layer"] = df["DL EARFCN"].map(earfcn_to_layer)
    print(f"  {len(df):,} rows")
    return df


def merge_parts() -> pd.DataFrame:
    files = idle_csv_files()
    if not files:
        sys.exit(f"No idle CSV files in {IDLE_DIR}; run extract first")
    frames = [load_idle_csv(p) for p in files]
    merged = pd.concat(frames, ignore_index=True)
    merged = merged.drop_duplicates(subset=["Time", "Cell Id", "DL EARFCN", "RSRP"], keep="first")
    merged = merged.sort_values(["Time", "Source"], kind="mergesort").reset_index(drop=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    merged.to_csv(MERGED_CSV, index=False, compression="gzip")
    print(f"Wrote merged file: {MERGED_CSV} ({len(merged):,} rows)")
    return merged


def load_active_csv(path: Path) -> pd.DataFrame:
    """Active Mode UE export. KPIs are already named RSRP / RSRQ / SINR.

    Older exports name the operator Home Operator and carry the cell as CGI.
    Borga_Active_mode_11Nov26 already has Operator and Cell Id. Cell Id is the
    E-UTRAN cell identity. When only a CGI is present, Cell Id is its last number.
    A secondary cell often has no CGI, so its Cell Id stays empty.
    """
    print(f"Loading active {path.name} ...")
    df = pd.read_csv(path)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
    if "Home Operator" in df.columns and "Operator" not in df.columns:
        df = df.rename(columns={"Home Operator": "Operator"})
    if "Cell Id" not in df.columns:
        if "CGI" not in df.columns:
            sys.exit(f"{path.name} has no Cell Id and no CGI")
        df["Cell Id"] = df["CGI"].map(_cgi_eci)
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        sys.exit(f"{path.name} is missing columns: {missing}")
    keep = [c for c in df.columns if c in set(COLUMNS) | {
        "Operator", "Data Technology", "Technology Band", "Collection", "Campaign", "Type",
    }]
    df = df[keep].copy()
    if "Operator" in df.columns:
        df["Operator"] = df["Operator"].astype("string").str.strip()
    df["Time"] = _parse_active_time(df["Time"])
    for col in ("RSRP", "RSRQ", "SINR", "Longitude", "Latitude", "DL EARFCN"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["Cell Id"] = pd.to_numeric(df["Cell Id"], errors="coerce").astype("Int64")
    df["DL EARFCN"] = df["DL EARFCN"].astype("Int64")
    df["Source"] = path.name
    df["Band"] = df["DL EARFCN"].map(earfcn_to_band)
    df["Layer"] = df["DL EARFCN"].map(earfcn_to_layer)
    print(f"  {len(df):,} rows")
    return df


def merge_active() -> pd.DataFrame:
    if not ACTIVE_ARCHIVE.exists():
        sys.exit(f"Missing active log: {ACTIVE_ARCHIVE.name}")
    print(f"Extracting {ACTIVE_ARCHIVE.name} -> {ACTIVE_DIR}")
    extract_7z(ACTIVE_ARCHIVE, EXTRACT_DIR)
    files = active_csv_files()
    if not files:
        sys.exit(f"Extraction finished but no CSV files were found in {ACTIVE_DIR}")
    print(f"  {len(files)} active CSV files")
    frames = [load_active_csv(path) for path in files]
    merged = pd.concat(frames, ignore_index=True)
    merged = merged.drop_duplicates(subset=["Time", "Operator", "Cell Id", "DL EARFCN", "RSRP"], keep="first")
    sort_cols = ["Time", "Source"]
    if "Operator" in merged.columns:
        sort_cols = ["Time", "Operator", "Source"]
    merged = merged.sort_values(sort_cols, kind="mergesort").reset_index(drop=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    merged.to_csv(ACTIVE_MERGED, index=False, compression="gzip")
    print(f"Wrote merged active file: {ACTIVE_MERGED} ({len(merged):,} rows)")
    return merged


def format_stats(series: pd.Series) -> str:
    s = series.dropna()
    if s.empty:
        return "n=0"
    return (
        f"n={len(s):,}  min={s.min():.1f}  p5={s.quantile(0.05):.1f}  "
        f"mean={s.mean():.1f}  median={s.median():.1f}  "
        f"p95={s.quantile(0.95):.1f}  max={s.max():.1f}"
    )


def write_summary(df: pd.DataFrame) -> None:
    lines = [
        "# Bogura Drive-Test Analysis Summary",
        "",
        "## What the source files are",
        "",
        "- `Borga_Idle_V2.0.7z` is the IDLE Mode scanner export (folder `Borga_Idle_V2.0/` with `1.csv` … `15.csv`).",
        "- Columns include Time, Longitude, Latitude, Cell Id, DL EARFCN, and scanner KPIs `NB RSRP` / `NB RSRQ` / `NB RS SINR` (renamed to RSRP / RSRQ / SINR).",
        "- Concatenating the 15 CSVs produces the **final log**: `output/Bogura_merged.csv.gz`.",
        "",
        "## Row counts",
        "",
        f"- CSV files: {df['Source'].nunique()}",
        f"- Merged scanner detections: {len(df):,} samples",
        "",
        "## Time and area",
        "",
    ]
    for source, part in df.groupby("Source", sort=False):
        tmin, tmax = part["Time"].min(), part["Time"].max()
        lines.append(
            f"- {source}: {tmin} → {tmax}; "
            f"lon {part['Longitude'].min():.5f}–{part['Longitude'].max():.5f}, "
            f"lat {part['Latitude'].min():.5f}–{part['Latitude'].max():.5f}"
        )
    lines += [
        "",
        "## Radio KPIs (merged)",
        "",
        f"- RSRP: {format_stats(df['RSRP'])}",
        f"- RSRQ: {format_stats(df['RSRQ'])}",
        f"- SINR: {format_stats(df['SINR'])} (blank SINR samples are skipped)",
        "",
        "## RSRP ranges (dBm)",
        "",
    ]
    rsrp_counts = rsrp_range_counts(df["RSRP"])
    for label, n in rsrp_counts.items():
        lines.append(f"- {label}: {int(n):,} ({n / len(df) * 100:.1f}%)")

    lines += ["", "## Bands (from DL EARFCN)", ""]
    band_counts = df["Band"].value_counts()
    for band, n in band_counts.items():
        lines.append(f"- {band}: {int(n):,} ({n / len(df) * 100:.1f}%)")

    lines += [
        "",
        f"- Unique Cell Ids: {df['Cell Id'].nunique():,}",
        "",
        "## Notes",
        "",
        "- Campaign is **Scanner** (NB RSRP / NB RSRQ / NB RS SINR). Coverage maps use the strongest RSRP per scanner sweep (~4 s / 40 m) so every band measured at a location competes for best server.",
        "- Files 1.csv–15.csv cover 11–21 May 2026. Operator is Robi. Layers: L900 (B8), L1800 (B3), L2100 (B1), L2600 (B41).",
        "",
        "Plots are written to `output/plots/`.",
        "",
    ]
    SUMMARY_MD.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {SUMMARY_MD}")


def downsample(df: pd.DataFrame, max_points: int = MAP_MAX_POINTS) -> pd.DataFrame:
    if len(df) <= max_points:
        return df
    step = max(1, len(df) // max_points)
    return df.iloc[::step]


def style_axes(ax, title: str, xlabel: str, ylabel: str) -> None:
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, alpha=0.3)


def plot_stacked_map_hist(ax, values, classifier, labels, colors, bins, title, xlabel, legend_title, xlim) -> None:
    v = pd.to_numeric(values, errors="coerce")
    cls = classifier(v)
    bottom = np.zeros(len(bins) - 1)
    widths = np.diff(bins)
    for lab, color in zip(labels, colors):
        counts, _ = np.histogram(v[cls == lab].dropna(), bins=bins)
        ax.bar(
            bins[:-1],
            counts,
            width=widths,
            align="edge",
            bottom=bottom,
            color=color,
            edgecolor="none",
            label=lab,
        )
        bottom += counts
    ax.legend(title=legend_title, fontsize=6, loc="upper right")
    style_axes(ax, title, xlabel, "Samples")
    ax.set_xlim(*xlim)


def plot_all(df: pd.DataFrame) -> None:
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    geo = best_server(df.dropna(subset=["Longitude", "Latitude", "RSRP"]))
    plot_rsrp_coverage_map(geo, PLOTS_DIR / "01_route_rsrp.png")
    plot_rsrq_coverage_map(geo, PLOTS_DIR / "01_route_rsrq.png")
    plot_sinr_coverage_map(geo, PLOTS_DIR / "01_route_sinr.png")
    rsrp_spots = find_bad_spots(geo, "RSRP")
    rsrq_spots = find_bad_spots(geo, "RSRQ")
    sinr_spots = find_bad_spots(geo, "SINR")
    plot_rsrp_bad_spot_map(geo, rsrp_spots, PLOTS_DIR / "08_bad_spots_rsrp.png")
    plot_rsrq_bad_spot_map(geo, rsrq_spots, PLOTS_DIR / "08_bad_spots_rsrq.png")
    plot_sinr_bad_spot_map(geo, sinr_spots, PLOTS_DIR / "08_bad_spots_sinr.png")

    fig, axes = plt.subplots(1, 3, figsize=(16, 5.2))
    plot_stacked_map_hist(
        axes[0], df["RSRP"], rsrp_map_class, MAP_RSRP_LABELS, MAP_RSRP_COLORS,
        np.arange(-140, -40, 2), "RSRP distribution", "RSRP (dBm)", "RSRP (dBm)", (-140, -50),
    )
    plot_stacked_map_hist(
        axes[1], df["RSRQ"], rsrq_map_class, MAP_RSRQ_LABELS, MAP_RSRQ_COLORS,
        np.arange(-30, 0, 0.5), "RSRQ distribution", "RSRQ (dB)", "RSRQ (dB)", (-24, -3),
    )
    plot_stacked_map_hist(
        axes[2], df["SINR"], sinr_map_class, MAP_SINR_LABELS, MAP_SINR_COLORS,
        np.arange(-20, 32, 1), "SINR distribution", "SINR (dB)", "SINR (dB)", (-15, 30),
    )
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "02_kpi_histograms.png", dpi=140)
    plt.close(fig)

    fig, axes = plt.subplots(3, 1, figsize=(12, 9), sharex=True)
    ts = downsample(df.dropna(subset=["Time"]), max_points=40000)
    for ax, col in zip(axes, ("RSRP", "RSRQ", "SINR")):
        ax.scatter(ts["Time"], ts[col], s=2, alpha=0.25, color="#1d3557", linewidths=0)
        style_axes(ax, f"{col} vs time", "Time", col)
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "03_kpi_vs_time.png", dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(9, 5.5))
    counts = rsrp_range_counts(df["RSRP"])
    ax.barh(list(counts.index[::-1]), counts.values[::-1], color=list(reversed(RSRP_COLORS)))
    style_axes(ax, "RSRP ranges (dBm)", "Samples", "")
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "04_rsrp_quality_bins.png", dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8, 5))
    band_counts = df["Band"].value_counts()
    ax.bar(band_counts.index.astype(str), band_counts.values, color="#457b9d")
    ax.tick_params(axis="x", rotation=20)
    style_axes(ax, "Samples by LTE band (DL EARFCN)", "Band", "Samples")
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "05_band_counts.png", dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 6))
    top_cells = df["Cell Id"].dropna().astype("int64").value_counts().head(15)
    ax.barh(top_cells.index.astype(str)[::-1], top_cells.values[::-1], color="#1d3557")
    style_axes(ax, "Top 15 scanner Cell Ids", "Samples", "Cell Id")
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "06_top_cells.png", dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 5))
    source_counts = df["Source"].value_counts()
    order = sorted(source_counts.index, key=lambda n: int(str(n).replace(".csv", "")) if str(n).replace(".csv", "").isdigit() else str(n))
    source_counts = source_counts.reindex(order)
    ax.bar(source_counts.index.astype(str), source_counts.values, color="#2a6f97")
    ax.tick_params(axis="x", rotation=45)
    style_axes(ax, "Merged idle log: samples by CSV", "Source file", "Samples")
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "07_csv_counts.png", dpi=140)
    plt.close(fig)

    print(f"Wrote plots to {PLOTS_DIR}")


def _finalize_merged(df: pd.DataFrame) -> pd.DataFrame:
    df["Time"] = pd.to_datetime(df["Time"], errors="coerce")
    df["Cell Id"] = pd.to_numeric(df["Cell Id"], errors="coerce").astype("Int64")
    df["DL EARFCN"] = pd.to_numeric(df["DL EARFCN"], errors="coerce").astype("Int64")
    df["Band"] = df["DL EARFCN"].map(earfcn_to_band)
    df["Layer"] = df["DL EARFCN"].map(earfcn_to_layer)
    return df


def load_merged_csv() -> pd.DataFrame:
    print(f"Loading existing merged file {MERGED_CSV}")
    return _finalize_merged(pd.read_csv(MERGED_CSV))


def load_active_merged() -> pd.DataFrame:
    print(f"Loading existing merged file {ACTIVE_MERGED}")
    return _finalize_merged(pd.read_csv(ACTIVE_MERGED))


def main() -> None:
    if "--from-merged" in sys.argv:
        if not MERGED_CSV.exists():
            sys.exit(f"{MERGED_CSV} not found; run without --from-merged first")
        merged = load_merged_csv()
        if not ACTIVE_MERGED.exists():
            sys.exit(f"{ACTIVE_MERGED} not found; run without --from-merged first")
        active = load_active_merged()
    else:
        extract_archives()
        merged = merge_parts()
        active = merge_active()
    write_summary(merged)
    plot_all(merged)
    from generate_excel_report import build_report, verify_report

    report_path = build_report(merged, active)
    result = verify_report(report_path)
    if result["problems"]:
        raise SystemExit("Excel view checks failed: " + "; ".join(result["problems"]))
    print("Excel report verified: no freeze panes, no gridlines.")
    print("Done.")


if __name__ == "__main__":
    main()
