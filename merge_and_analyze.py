#!/usr/bin/env python3
"""Extract split Bogura drive-test archives, merge P1+P2, and plot KPIs."""

from __future__ import annotations

import subprocess
import sys
from collections import deque
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Ellipse, PathPatch
from matplotlib.path import Path as MPath
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
EXTRACT_DIR = ROOT / "extracted"
OUTPUT_DIR = ROOT / "output"
PLOTS_DIR = OUTPUT_DIR / "plots"

P1_PARTS = [ROOT / f"Bogura_P1.part{i}.rar" for i in range(1, 6)]
P2_RAR = ROOT / "Bogura_P2.rar"
P1_XLSX = EXTRACT_DIR / "Bogura_P1.xlsx"
P2_XLSX = EXTRACT_DIR / "Bogura_P2.xlsx"
MERGED_CSV = OUTPUT_DIR / "Bogura_merged.csv.gz"
SUMMARY_MD = OUTPUT_DIR / "analysis_summary.md"
SITE_DB_PATH = ROOT / "Physical_Site_Database_V1.0.xlsb"

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

MAP_RSRP_LABELS = [
    "-90 <= X < Max",
    "-100 <= X < -90",
    "-110 <= X < -100",
    "-115 <= X < -110",
    "-120 <= X < -115",
    "-Min <= X < -120",
]
MAP_RSRP_THRESHOLDS = [-90, -100, -110, -115, -120]
MAP_RSRP_COLORS = [
    LEGEND_BLUE,
    LEGEND_DARK_GREEN,
    LEGEND_LIGHT_GREEN,
    LEGEND_YELLOW,
    LEGEND_MAGENTA,
    LEGEND_RED,
]

# Source table listed RSRQ worst-first and wrote "X < -5" for the blue bin.
# Blue is the best colour, so that row is X >= -5; red is X < -20.
MAP_RSRQ_LABELS = [
    "X >= -5",
    "-10 <= X < -5",
    "-15 <= X < -10",
    "-20 <= X < -15",
    "X < -20",
]
MAP_RSRQ_THRESHOLDS = [-5, -10, -15, -20]
MAP_RSRQ_COLORS = [
    LEGEND_BLUE,
    LEGEND_DARK_GREEN,
    LEGEND_LIGHT_GREEN,
    LEGEND_YELLOW,
    LEGEND_RED,
]

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
    """Classify RSRP into the coverage-map legend bins."""
    return _map_class(series, MAP_RSRP_THRESHOLDS, MAP_RSRP_LABELS)


def rsrq_map_class(series: pd.Series) -> pd.Series:
    return _map_class(series, MAP_RSRQ_THRESHOLDS, MAP_RSRQ_LABELS)


def sinr_map_class(series: pd.Series) -> pd.Series:
    return _map_class(series, MAP_SINR_THRESHOLDS, MAP_SINR_LABELS)


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
        cells = cells.rename(columns={c: c.strip() for c in cells.columns})
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


def plot_discrete_coverage_map(
    df: pd.DataFrame,
    path: Path,
    column: str,
    class_series: pd.Series,
    labels: list[str],
    colors: list[str],
    title: str,
    legend_title: str,
    max_points: int = 25000,
    extent: tuple[float, float, float, float] | None = None,
    layer: str | None = None,
) -> Path:
    """Drive route colored by discrete KPI ranges, with an RSRP-style side legend."""
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    geo = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    fig, ax = plt.subplots(figsize=(8.0, 9.2))
    if not geo.empty:
        sample = downsample(geo, max_points=max_points)
        classes = class_series.reindex(sample.index)
        for label, color in zip(reversed(labels), reversed(colors)):
            part = sample[classes == label]
            if part.empty:
                continue
            ax.scatter(part["Longitude"], part["Latitude"], s=5, c=color, linewidths=0, rasterized=True, zorder=2)
    if extent is not None:
        ax.set_xlim(extent[0], extent[1])
        ax.set_ylim(extent[2], extent[3])
    elif not geo.empty:
        lon_min, lon_max = float(geo["Longitude"].min()), float(geo["Longitude"].max())
        lat_min, lat_max = float(geo["Latitude"].min()), float(geo["Latitude"].max())
        pad_lon = 0.02 * (lon_max - lon_min)
        pad_lat = 0.02 * (lat_max - lat_min)
        ax.set_xlim(lon_min - pad_lon, lon_max + pad_lon)
        ax.set_ylim(lat_min - pad_lat, lat_max + pad_lat)
    if not geo.empty:
        draw_site_pies(ax, geo, layer=layer)
    else:
        ax.text(
            0.5,
            0.5,
            "No IDLE samples on this band",
            transform=ax.transAxes,
            ha="center",
            va="center",
            color="#7F8C8D",
            fontsize=12,
        )
    handles = [
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
    ax.legend(
        handles=handles,
        title=legend_title,
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=True,
        fontsize=8,
        title_fontsize=10,
    )
    ax.set_aspect("equal", adjustable="box")
    ax.set_title(title)
    hide_map_axes(ax)
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return path


def plot_rsrp_coverage_map(
    df: pd.DataFrame,
    path: Path,
    max_points: int = 25000,
    *,
    layer: str | None = None,
    extent: tuple[float, float, float, float] | None = None,
) -> Path:
    """Discrete RSRP coverage map with a range legend (no lon/lat chrome)."""
    title = "Bogura coverage map — RSRP"
    if layer:
        title = f"{title} — {layer}"
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
    )


def plot_rsrq_coverage_map(
    df: pd.DataFrame,
    path: Path,
    max_points: int = 25000,
    *,
    layer: str | None = None,
    extent: tuple[float, float, float, float] | None = None,
) -> Path:
    title = "Bogura coverage map — RSRQ"
    if layer:
        title = f"{title} — {layer}"
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
    )


def plot_sinr_coverage_map(
    df: pd.DataFrame,
    path: Path,
    max_points: int = 25000,
    *,
    layer: str | None = None,
    extent: tuple[float, float, float, float] | None = None,
) -> Path:
    title = "Bogura coverage map — SINR"
    if layer:
        title = f"{title} — {layer}"
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
    )


# Poor-sample clustering. Use the full drive map and keep every dense
# yellow/magenta/red stretch — not just the three worst nearby clusters.
BAD_SPOT_GRID_DEG = 0.01  # ~1.1 km cells
BAD_SPOT_MIN_POOR = 180
BAD_SPOT_MAX = 40
BAD_SPOT_RULES = {
    "RSRP": {"threshold": -115.0, "unit": "dBm"},
    "RSRQ": {"threshold": -15.0, "unit": "dB"},
    "SINR": {"threshold": 0.0, "unit": "dB"},
}


def _km_per_deg_lon(lat: float) -> float:
    return 111.32 * float(np.cos(np.radians(lat)))


def _equirect_km(lat1, lon1, lat2, lon2) -> np.ndarray:
    mid = np.radians((np.asarray(lat1) + np.asarray(lat2)) / 2.0)
    dlat = (np.asarray(lat2) - np.asarray(lat1)) * 111.32
    dlon = (np.asarray(lon2) - np.asarray(lon1)) * 111.32 * np.cos(mid)
    return np.sqrt(dlat**2 + dlon**2)


def cluster_poor_cells(
    df: pd.DataFrame,
    column: str,
    threshold: float,
    grid_deg: float = BAD_SPOT_GRID_DEG,
    min_poor: int = BAD_SPOT_MIN_POOR,
) -> pd.DataFrame:
    """Cluster poor samples on a ~1 km grid and merge adjacent cells.

    Every dense yellow/magenta/red stretch becomes a spot. Nearby cells join;
    separate holes on other roads stay separate.
    """
    g = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    values = pd.to_numeric(g[column], errors="coerce")
    g = g.loc[values.notna() & (values < threshold)].copy()
    g[column] = values.loc[g.index]
    if g.empty:
        return pd.DataFrame()
    g["latb"] = np.floor(g["Latitude"] / grid_deg).astype(int)
    g["lonb"] = np.floor(g["Longitude"] / grid_deg).astype(int)
    agg = g.groupby(["latb", "lonb"], sort=False).agg(
        n_poor=(column, "size"),
        mean=(column, "mean"),
        lat=("Latitude", "mean"),
        lon=("Longitude", "mean"),
        lat_min=("Latitude", "min"),
        lat_max=("Latitude", "max"),
        lon_min=("Longitude", "min"),
        lon_max=("Longitude", "max"),
    )
    hot = agg[agg["n_poor"] >= min_poor]
    if hot.empty:
        return pd.DataFrame()

    cells = set(hot.index)
    neigh = ((1, 0), (-1, 0), (0, 1), (0, -1), (1, 1), (1, -1), (-1, 1), (-1, -1))
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
                    seen.add(nb)
                    queue.append(nb)
                    group.append(nb)
        part = hot.loc[group]
        n_poor = int(part["n_poor"].sum())
        weights = part["n_poor"].to_numpy(dtype=float)
        mean = float(np.average(part["mean"], weights=weights))
        lat = float(np.average(part["lat"], weights=weights))
        lon = float(np.average(part["lon"], weights=weights))
        lat_min = float(part["lat_min"].min())
        lat_max = float(part["lat_max"].max())
        lon_min = float(part["lon_min"].min())
        lon_max = float(part["lon_max"].max())
        rows.append(
            {
                "n": n_poor,
                "n_poor": n_poor,
                "poor_pct": 1.0,
                "mean": mean,
                "lat": lat,
                "lon": lon,
                "lat_min": lat_min,
                "lat_max": lat_max,
                "lon_min": lon_min,
                "lon_max": lon_max,
                "span_lat_km": (lat_max - lat_min) * 111.32,
                "span_lon_km": (lon_max - lon_min) * _km_per_deg_lon(lat),
                "score": n_poor * max(0.0, threshold - mean),
            }
        )
    return pd.DataFrame(rows).sort_values("score", ascending=False).reset_index(drop=True)


def select_all_spots(clusters: pd.DataFrame, max_spots: int = BAD_SPOT_MAX) -> pd.DataFrame:
    """Keep every significant cluster on the full drive, not one local window."""
    if clusters.empty:
        return clusters
    return clusters.sort_values("score", ascending=False).head(max_spots).reset_index(drop=True)


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


def find_bad_spots(df: pd.DataFrame, column: str, max_spots: int = BAD_SPOT_MAX) -> pd.DataFrame:
    """Return every significant bad spot for one KPI on the full drive, numbered 1..N."""
    rule = BAD_SPOT_RULES[column]
    clusters = cluster_poor_cells(df, column, rule["threshold"])
    spots = select_all_spots(clusters, max_spots=max_spots)
    if spots.empty:
        return spots
    extras = [_spot_extra_stats(df, row) for _, row in spots.iterrows()]
    spots = spots.copy()
    bbox_n = []
    bbox_poor = []
    for _, row in spots.iterrows():
        box = df[
            df["Latitude"].between(row["lat_min"], row["lat_max"])
            & df["Longitude"].between(row["lon_min"], row["lon_max"])
        ]
        vals = pd.to_numeric(box[column], errors="coerce")
        n = int(vals.notna().sum())
        n_poor = int((vals < rule["threshold"]).sum())
        bbox_n.append(n)
        bbox_poor.append(n_poor)
    spots["spot"] = np.arange(1, len(spots) + 1)
    spots["kpi"] = column
    spots["threshold"] = rule["threshold"]
    spots["unit"] = rule["unit"]
    spots["n"] = bbox_n
    spots["n_poor"] = bbox_poor
    spots["poor_pct"] = np.array(bbox_poor) / np.maximum(np.array(bbox_n), 1)
    spots["mean_rsrp"] = [e["mean_rsrp"] for e in extras]
    spots["mean_rsrq"] = [e["mean_rsrq"] for e in extras]
    spots["mean_sinr"] = [e["mean_sinr"] for e in extras]
    spots["top_cell"] = [e["top_cell"] for e in extras]
    return spots


def _ellipse_from_points(lons: np.ndarray, lats: np.ndarray, min_km: float = 4.0, max_km: float = 14.0) -> tuple[float, float, float, float, float]:
    """Return ellipse center/width/height/angle so a dashed oval follows the poor samples."""
    lon0 = float(np.mean(lons))
    lat0 = float(np.mean(lats))
    km_lon = _km_per_deg_lon(lat0)
    if len(lons) < 12:
        lon_span = max(float(np.max(lons) - np.min(lons)) if len(lons) else 0.0, min_km / km_lon)
        lat_span = max(float(np.max(lats) - np.min(lats)) if len(lats) else 0.0, min_km / 111.32)
        width_deg = min(lon_span * 1.5, max_km / km_lon)
        height_deg = min(lat_span * 1.5, max_km / 111.32)
        return lon0, lat0, width_deg, height_deg, 0.0
    x = (lons - lon0) * km_lon
    y = (lats - lat0) * 111.32
    cov = np.cov(np.vstack([x, y]))
    eigvals, eigvecs = np.linalg.eigh(cov)
    order = np.argsort(eigvals)[::-1]
    eigvals = np.clip(eigvals[order], 0.05, None)
    eigvecs = eigvecs[:, order]
    width_km = min(max(2.2 * 2.0 * np.sqrt(eigvals[0]), min_km), max_km)
    height_km = min(max(2.2 * 2.0 * np.sqrt(eigvals[1]), min_km * 0.55), max_km * 0.7)
    angle = float(np.degrees(np.arctan2(eigvecs[1, 0], eigvecs[0, 0])))
    width_deg = width_km / km_lon
    height_deg = height_km / 111.32
    return lon0, lat0, width_deg, height_deg, angle


def add_map_scale_bar(ax, km: float = 4.0) -> None:
    xlim = ax.get_xlim()
    ylim = ax.get_ylim()
    lat = ylim[1] - 0.06 * (ylim[1] - ylim[0])
    lon0 = xlim[0] + 0.05 * (xlim[1] - xlim[0])
    width = km / _km_per_deg_lon(lat)
    ax.plot([lon0, lon0 + width], [lat, lat], color="#1B1B1B", lw=2.8, solid_capstyle="butt", zorder=6)
    ax.text(lon0 + width / 2, lat, f"  {km:.0f} km", ha="center", va="bottom", fontsize=8, color="#1B1B1B", zorder=6)


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
    max_points: int = 25000,
) -> Path:
    """Full-drive coverage map with every significant poor stretch circled."""
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    geo = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    classes_all = class_series.reindex(geo.index)
    sample = downsample(geo, max_points=max_points)
    classes = classes_all.reindex(sample.index)
    lat_min = float(geo["Latitude"].min())
    lat_max = float(geo["Latitude"].max())
    lon_min = float(geo["Longitude"].min())
    lon_max = float(geo["Longitude"].max())
    pad_lat = 0.02 * (lat_max - lat_min)
    pad_lon = 0.02 * (lon_max - lon_min)
    lat_min -= pad_lat
    lat_max += pad_lat
    lon_min -= pad_lon
    lon_max += pad_lon

    fig, ax = plt.subplots(figsize=(8.0, 9.2))
    for label, color in zip(reversed(labels), reversed(colors)):
        part = sample[classes == label]
        if part.empty:
            continue
        ax.scatter(part["Longitude"], part["Latitude"], s=5, c=color, linewidths=0, rasterized=True, zorder=2)
    handles = [
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
    ax.set_xlim(lon_min, lon_max)
    ax.set_ylim(lat_min, lat_max)
    ax.set_aspect("equal", adjustable="box")
    site_handles = draw_site_pies(ax, geo)
    ax.legend(
        handles=handles + site_handles,
        title=legend_title,
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=True,
        fontsize=8,
        title_fontsize=9,
    )

    if spots is not None and not spots.empty:
        threshold = float(spots["threshold"].iloc[0])
        poor = geo[pd.to_numeric(geo[column], errors="coerce") < threshold]
        for _, spot in spots.iterrows():
            box_poor = poor[
                poor["Latitude"].between(spot["lat_min"], spot["lat_max"])
                & poor["Longitude"].between(spot["lon_min"], spot["lon_max"])
            ]
            if box_poor.empty:
                cx, cy, w, h, angle = float(spot["lon"]), float(spot["lat"]), 0.012, 0.01, 0.0
            else:
                cx, cy, w, h, angle = _ellipse_from_points(
                    box_poor["Longitude"].to_numpy(),
                    box_poor["Latitude"].to_numpy(),
                )
            ax.add_patch(
                Ellipse(
                    (cx, cy),
                    width=w,
                    height=h,
                    angle=angle,
                    fill=False,
                    edgecolor="#E74C3C",
                    linestyle=(0, (6, 4)),
                    linewidth=1.5,
                    zorder=4,
                )
            )
            ax.text(
                cx,
                cy,
                str(int(spot["spot"])),
                fontsize=7,
                fontweight="bold",
                color="#922B21",
                ha="center",
                va="center",
                zorder=5,
                bbox={
                    "boxstyle": "round,pad=0.12",
                    "facecolor": "#FDEBD0",
                    "edgecolor": "#922B21",
                    "linewidth": 0.7,
                },
            )
    add_map_scale_bar(ax, km=20)
    ax.set_title(title)
    hide_map_axes(ax)
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return path


def plot_rsrp_bad_spot_map(df: pd.DataFrame, spots: pd.DataFrame, path: Path) -> Path:
    return plot_bad_spot_map(
        df,
        path,
        column="RSRP",
        class_series=rsrp_map_class(df["RSRP"]),
        labels=MAP_RSRP_LABELS,
        colors=MAP_RSRP_COLORS,
        spots=spots,
        title="Bad spot analysis — RSRP",
        legend_title="RSRP (dBm)",
    )


def plot_rsrq_bad_spot_map(df: pd.DataFrame, spots: pd.DataFrame, path: Path) -> Path:
    return plot_bad_spot_map(
        df,
        path,
        column="RSRQ",
        class_series=rsrq_map_class(df["RSRQ"]),
        labels=MAP_RSRQ_LABELS,
        colors=MAP_RSRQ_COLORS,
        spots=spots,
        title="Bad spot analysis — RSRQ",
        legend_title="RSRQ (dB)",
    )


def plot_sinr_bad_spot_map(df: pd.DataFrame, spots: pd.DataFrame, path: Path) -> Path:
    return plot_bad_spot_map(
        df,
        path,
        column="SINR",
        class_series=sinr_map_class(df["SINR"]),
        labels=MAP_SINR_LABELS,
        colors=MAP_SINR_COLORS,
        spots=spots,
        title="Bad spot analysis — SINR",
        legend_title="SINR (dB)",
    )


def require_unrar() -> str:
    from shutil import which

    exe = which("unrar")
    if not exe:
        sys.exit("unrar is required. Install it with: sudo apt-get install unrar")
    return exe


def extract_archives() -> None:
    EXTRACT_DIR.mkdir(parents=True, exist_ok=True)
    unrar = require_unrar()
    missing = [p for p in P1_PARTS if not p.exists()]
    if missing:
        sys.exit(f"Missing P1 split volumes: {', '.join(p.name for p in missing)}")
    if not P2_RAR.exists():
        sys.exit(f"Missing {P2_RAR.name}")

    print("Extracting Bogura_P1.part1–part5 -> Bogura_P1.xlsx")
    subprocess.run(
        [unrar, "x", "-o+", str(P1_PARTS[0]), str(EXTRACT_DIR) + "/"],
        check=True,
    )
    print("Extracting Bogura_P2.rar -> Bogura_P2.xlsx")
    subprocess.run(
        [unrar, "x", "-o+", str(P2_RAR), str(EXTRACT_DIR) + "/"],
        check=True,
    )
    if not P1_XLSX.exists() or not P2_XLSX.exists():
        sys.exit("Extraction finished but expected xlsx files were not found.")


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


def load_part(path: Path, source: str) -> pd.DataFrame:
    print(f"Loading {path.name} ...")
    df = pd.read_excel(path, sheet_name="Sheet1", engine="openpyxl")
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        sys.exit(f"{path.name} is missing columns: {missing}")
    df = df[COLUMNS].copy()
    df["Time"] = pd.to_datetime(df["Time"], errors="coerce")
    for col in ("RSRP", "RSRQ", "SINR", "Longitude", "Latitude", "DL EARFCN"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["Cell Id"] = pd.to_numeric(df["Cell Id"], errors="coerce").astype("Int64")
    df["DL EARFCN"] = df["DL EARFCN"].astype("Int64")
    df["Source"] = source
    df["Band"] = df["DL EARFCN"].map(earfcn_to_band)
    print(f"  {len(df):,} rows")
    return df


def merge_parts() -> pd.DataFrame:
    p1 = load_part(P1_XLSX, "P1")
    p2 = load_part(P2_XLSX, "P2")
    merged = pd.concat([p1, p2], ignore_index=True)
    merged = merged.sort_values(["Time", "Source"], kind="mergesort").reset_index(drop=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    merged.to_csv(MERGED_CSV, index=False, compression="gzip")
    print(f"Wrote merged file: {MERGED_CSV} ({len(merged):,} rows)")
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
        "- `Bogura_P1.part1.rar` … `Bogura_P1.part5.rar` are one split RAR volume set. Extracting them yields **Bogura_P1.xlsx** (Part 1 of the drive test).",
        "- `Bogura_P2.rar` is already a complete archive. Extracting it yields **Bogura_P2.xlsx** (Part 2).",
        "- Both workbooks use the same 8 columns: Time, RSRP, RSRQ, SINR, Longitude, Latitude, Cell Id, DL EARFCN.",
        "- Concatenating P1 then P2 produces the **final log**: `output/Bogura_merged.csv.gz`.",
        "",
        "## Row counts",
        "",
        f"- P1: {int((df['Source'] == 'P1').sum()):,} samples",
        f"- P2: {int((df['Source'] == 'P2').sum()):,} samples",
        f"- Merged: {len(df):,} samples",
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
        "- P1 covers multiple drive days (11–20 May 2026). P2 is a shorter session on 21 May 2026.",
        "- About 24% of samples have no SINR (typical when the logger recorded neighbor cells).",
        "- Band 41 (TDD 2500) appears as EARFCN ~40742 / 40940 / 41138.",
        "",
        "Plots are written to `output/plots/`.",
        "",
    ]
    SUMMARY_MD.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {SUMMARY_MD}")


def downsample(df: pd.DataFrame, max_points: int = 25000) -> pd.DataFrame:
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
    geo = df.dropna(subset=["Longitude", "Latitude", "RSRP"])
    plot_rsrp_coverage_map(geo, PLOTS_DIR / "01_route_rsrp.png")
    plot_rsrq_coverage_map(geo, PLOTS_DIR / "01_route_rsrq.png")
    plot_sinr_coverage_map(geo, PLOTS_DIR / "01_route_sinr.png")
    rsrp_spots = find_bad_spots(df, "RSRP")
    rsrq_spots = find_bad_spots(df, "RSRQ")
    sinr_spots = find_bad_spots(df, "SINR")
    plot_rsrp_bad_spot_map(df, rsrp_spots, PLOTS_DIR / "08_bad_spots_rsrp.png")
    plot_rsrq_bad_spot_map(df, rsrq_spots, PLOTS_DIR / "08_bad_spots_rsrq.png")
    plot_sinr_bad_spot_map(df, sinr_spots, PLOTS_DIR / "08_bad_spots_sinr.png")

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
    style_axes(ax, "Top 15 serving Cell Ids", "Samples", "Cell Id")
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "06_top_cells.png", dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(7, 5))
    source_counts = df["Source"].value_counts().reindex(["P1", "P2"])
    ax.bar(source_counts.index, source_counts.values, color=["#2a6f97", "#e07a3d"])
    style_axes(ax, "Merged log: P1 vs P2 samples", "Source file", "Samples")
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / "07_p1_vs_p2_counts.png", dpi=140)
    plt.close(fig)

    print(f"Wrote plots to {PLOTS_DIR}")


def load_merged_csv() -> pd.DataFrame:
    print(f"Loading existing merged file {MERGED_CSV}")
    df = pd.read_csv(MERGED_CSV)
    df["Time"] = pd.to_datetime(df["Time"], errors="coerce")
    df["Cell Id"] = pd.to_numeric(df["Cell Id"], errors="coerce").astype("Int64")
    df["DL EARFCN"] = pd.to_numeric(df["DL EARFCN"], errors="coerce").astype("Int64")
    df["Band"] = df["DL EARFCN"].map(earfcn_to_band)
    return df


def main() -> None:
    if "--from-merged" in sys.argv:
        if not MERGED_CSV.exists():
            sys.exit(f"{MERGED_CSV} not found; run without --from-merged first")
        merged = load_merged_csv()
    else:
        extract_archives()
        merged = merge_parts()
    write_summary(merged)
    plot_all(merged)
    from generate_excel_report import build_report, verify_report

    report_path = build_report(merged)
    result = verify_report(report_path)
    if result["problems"]:
        raise SystemExit("Excel view checks failed: " + "; ".join(result["problems"]))
    print("Excel report verified: no freeze panes, no gridlines.")
    print("Done.")


if __name__ == "__main__":
    main()
