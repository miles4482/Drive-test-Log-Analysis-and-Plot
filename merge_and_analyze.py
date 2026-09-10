#!/usr/bin/env python3
"""Extract split Bogura drive-test archives, merge P1+P2, and plot KPIs."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
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
) -> Path:
    """Drive route colored by discrete KPI ranges, with an RSRP-style side legend."""
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    geo = df.dropna(subset=["Longitude", "Latitude", column]).copy()
    sample = downsample(geo, max_points=max_points)
    classes = class_series.reindex(sample.index)
    fig, ax = plt.subplots(figsize=(8.0, 9.2))
    for label, color in zip(reversed(labels), reversed(colors)):
        part = sample[classes == label]
        if part.empty:
            continue
        ax.scatter(part["Longitude"], part["Latitude"], s=5, c=color, linewidths=0, rasterized=True)
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
        fontsize=9,
        title_fontsize=10,
    )
    ax.set_aspect("equal", adjustable="box")
    ax.set_title(title)
    hide_map_axes(ax)
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return path


def plot_rsrp_coverage_map(df: pd.DataFrame, path: Path, max_points: int = 25000) -> Path:
    """Discrete RSRP coverage map with a range legend (no lon/lat chrome)."""
    return plot_discrete_coverage_map(
        df,
        path,
        column="RSRP",
        class_series=rsrp_map_class(df["RSRP"]),
        labels=MAP_RSRP_LABELS,
        colors=MAP_RSRP_COLORS,
        title="Bogura coverage map — RSRP",
        legend_title="RSRP (dBm)",
        max_points=max_points,
    )


def plot_rsrq_coverage_map(df: pd.DataFrame, path: Path, max_points: int = 25000) -> Path:
    return plot_discrete_coverage_map(
        df,
        path,
        column="RSRQ",
        class_series=rsrq_map_class(df["RSRQ"]),
        labels=MAP_RSRQ_LABELS,
        colors=MAP_RSRQ_COLORS,
        title="Bogura coverage map — RSRQ",
        legend_title="RSRQ (dB)",
        max_points=max_points,
    )


def plot_sinr_coverage_map(df: pd.DataFrame, path: Path, max_points: int = 25000) -> Path:
    return plot_discrete_coverage_map(
        df,
        path,
        column="SINR",
        class_series=sinr_map_class(df["SINR"]),
        labels=MAP_SINR_LABELS,
        colors=MAP_SINR_COLORS,
        title="Bogura coverage map — SINR",
        legend_title="SINR (dB)",
        max_points=max_points,
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
