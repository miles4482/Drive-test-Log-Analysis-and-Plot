#!/usr/bin/env python3
"""Build a Bogura drive-test Excel report with RSRP/RSRQ/SINR charts.

Worksheet rules for every sheet:
- freeze panes off
- gridlines off (screen and print)
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import shutil
import zipfile

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import MultipleLocator
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from openpyxl.chart.text import RichText
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from merge_and_analyze import (
    OUTPUT_DIR,
    PLOTS_DIR,
    MAP_RSRP_COLORS,
    MAP_RSRP_LABELS,
    downsample,
    earfcn_to_band,
    plot_rsrp_coverage_map,
    plot_rsrq_coverage_map,
    plot_sinr_coverage_map,
    rsrp_map_class,
    rsrp_range_counts,
)

REPORT_VERSION = "1.6"
REPORT_XLSX = OUTPUT_DIR / "Bogura_DriveTest_Report.xlsx"
VERSIONED_XLSX = OUTPUT_DIR / f"Bogura_DriveTest_Report_v{REPORT_VERSION}.xlsx"

NAVY = "1B4F72"
WHITE = "FFFFFF"
LIGHT = "F4F6F7"
CARD = "EAF2F8"
RSRP_COLOR = "C0392B"
RSRQ_COLOR = "1F618D"
SINR_COLOR = "117A65"
THIN = Border(
    left=Side(style="thin", color="D5D8DC"),
    right=Side(style="thin", color="D5D8DC"),
    top=Side(style="thin", color="D5D8DC"),
    bottom=Side(style="thin", color="D5D8DC"),
)

RSRQ_BINS = [-np.inf, -20, -15, -10, np.inf]
RSRQ_LABELS = [
    "Poor (<-20)",
    "Fair (-20 to -15)",
    "Good (-15 to -10)",
    "Excellent (>-10)",
]
SINR_BINS = [-np.inf, 0, 13, 20, np.inf]
SINR_LABELS = [
    "Poor (<0)",
    "Fair (0 to 13)",
    "Good (13 to 20)",
    "Excellent (>20)",
]
RSRQ_STACK_COLORS = ["9B2226", "E07A3D", "90BE6D", "2A9D8F"]
SINR_STACK_COLORS = ["9B2226", "E9C46A", "2A9D8F", "1A9850"]


def apply_sheet_view(ws: Worksheet) -> None:
    """No freeze panes, no gridlines (on-screen or printed)."""
    ws.freeze_panes = None
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.view = "normal"


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def font(size=11, bold=False, color=NAVY, name="Calibri") -> Font:
    return Font(name=name, size=size, bold=bold, color=color)


def write_cell(ws, row, col, value, *, size=11, bold=False, color=NAVY, fill_color=None, align="left", num_fmt=None, wrap=False):
    cell = ws.cell(row, col, value)
    cell.font = font(size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill_color:
        cell.fill = fill(fill_color)
    if num_fmt:
        cell.number_format = num_fmt
    return cell


def set_widths(ws, widths: dict[str, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def kpi_stats(series: pd.Series) -> dict[str, float]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return {k: np.nan for k in ("Count", "Min", "P5", "Median", "Mean", "P95", "Max", "Std")}
    return {
        "Count": float(len(s)),
        "Min": float(s.min()),
        "P5": float(s.quantile(0.05)),
        "Median": float(s.median()),
        "Mean": float(s.mean()),
        "P95": float(s.quantile(0.95)),
        "Max": float(s.max()),
        "Std": float(s.std()),
    }


def quality_counts(series: pd.Series, bins, labels) -> pd.Series:
    binned = pd.cut(pd.to_numeric(series, errors="coerce"), bins=bins, labels=labels, right=False)
    return binned.value_counts(dropna=False).reindex(labels, fill_value=0)


def add_image(ws: Worksheet, path: Path, anchor: str, width: int, height: int) -> None:
    img = XLImage(str(path))
    img.width = width
    img.height = height
    ws.add_image(img, anchor)


def banner(ws: Worksheet, title: str, subtitle: str, last_col: int = 12) -> None:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    write_cell(ws, 1, 1, title, size=20, bold=True, color=WHITE, fill_color=NAVY, align="left")
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    write_cell(ws, 2, 1, subtitle, size=12, color=WHITE, fill_color=NAVY)
    ws.row_dimensions[2].height = 20
    for col in range(1, last_col + 1):
        ws.cell(1, col).fill = fill(NAVY)
        ws.cell(2, col).fill = fill(NAVY)
        ws.cell(1, col).font = font(20, bold=True, color=WHITE)
        ws.cell(2, col).font = font(12, color=WHITE)


def write_table(ws, start_row: int, start_col: int, headers: list[str], rows: list[list], header_fill=NAVY, num_formats: dict[int, str] | None = None):
    for i, h in enumerate(headers):
        cell = write_cell(ws, start_row, start_col + i, h, size=11, bold=True, color=WHITE, fill_color=header_fill, align="center")
        cell.border = THIN
    for r_i, row in enumerate(rows, start=start_row + 1):
        bg = LIGHT if (r_i - start_row) % 2 == 0 else WHITE
        for c_i, value in enumerate(row):
            fmt = (num_formats or {}).get(c_i)
            cell = write_cell(ws, r_i, start_col + c_i, value, fill_color=bg, align="center" if c_i else "left", num_fmt=fmt)
            cell.border = THIN
    return start_row + len(rows)


def _title_outside(title_obj) -> None:
    if title_obj is None:
        return
    try:
        title_obj.overlay = False
    except Exception:
        pass


def reserve_plot_area(chart, left=0.14, top=0.16, width=0.78, height=0.60) -> None:
    """Keep plot area inside the frame so axis titles and tick levels stay visible."""
    chart.layout = Layout(
        manualLayout=ManualLayout(
            xMode="edge",
            yMode="edge",
            x=left,
            y=top,
            w=width,
            h=height,
        )
    )


def apply_xy_levels(
    chart,
    x_title: str,
    y_title: str,
    y_min=None,
    y_max=None,
    show_all_x_ticks=True,
    rotate_x=False,
    y_major_unit=None,
    y_ax_pos="l",
    layout=True,
    y_num_fmt=None,
) -> None:
    """Put category levels on the bottom (X) and value levels on the left (Y).

    openpyxl defaults both axes to axPos='l', which Excel then draws without
    horizontal tick labels or a usable vertical scale — the bug in the v1.3 charts.
    """
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    _title_outside(getattr(chart, "title", None))
    _title_outside(chart.x_axis.title)
    _title_outside(chart.y_axis.title)
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = y_ax_pos
    chart.x_axis.tickLblPos = "nextTo"
    chart.y_axis.tickLblPos = "nextTo"
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"
    chart.x_axis.minorTickMark = "none"
    chart.y_axis.minorTickMark = "none"
    if y_num_fmt:
        chart.y_axis.numFmt = y_num_fmt
    if show_all_x_ticks:
        try:
            chart.x_axis.tickLblSkip = 1
            chart.x_axis.tickMarkSkip = 1
        except Exception:
            pass
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    if y_max is not None:
        chart.y_axis.scaling.max = y_max
    if y_major_unit is not None:
        try:
            chart.y_axis.majorUnit = y_major_unit
        except Exception:
            pass
    if rotate_x:
        # Rotation only — do not add an empty text run, which blanks tick labels.
        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(rot="-5400000", wrap="square", anchor="ctr"),
            p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=700)))],
        )
    if layout:
        reserve_plot_area(chart)


def style_chart(chart, color: str, title: str, y_title: str, x_title: str, width=14, height=8, show_legend=False) -> None:
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = x_title
    chart.style = 10
    chart.width = width
    chart.height = height
    if show_legend:
        chart.legend.position = "b"
    else:
        chart.legend = None
    if chart.series:
        chart.series[0].graphicalProperties.solidFill = color
        chart.series[0].graphicalProperties.line.solidFill = color
    apply_xy_levels(chart, x_title, y_title, show_all_x_ticks=False, rotate_x=False)


def add_col_chart(ws_data, ws_dest, anchor, title, color, cat_col, data_col, min_row, max_row, y_title, x_title, width=14, height=8, show_legend=False):
    chart = BarChart()
    chart.type = "col"
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_col, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    style_chart(chart, color, title, y_title, x_title, width=width, height=height, show_legend=show_legend)
    chart.shape = 4
    ws_dest.add_chart(chart, anchor)
    return chart


def add_stacked_col_chart(ws_data, ws_dest, anchor, title, cat_col, data_min, data_max, min_row, max_row, y_title, x_title, colors, width=14, height=8):
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_min, max_col=data_max, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = x_title
    chart.style = 10
    chart.width = width
    chart.height = height
    chart.legend.position = "b"
    apply_xy_levels(chart, x_title, y_title, show_all_x_ticks=False, rotate_x=False)
    for i, color in enumerate(colors):
        if i < len(chart.series):
            hexcol = color.lstrip("#")
            chart.series[i].graphicalProperties.solidFill = hexcol
            chart.series[i].graphicalProperties.line.solidFill = hexcol
    ws_dest.add_chart(chart, anchor)
    return chart


def add_line_chart(ws_data, ws_dest, anchor, title, color, cat_col, data_col, min_row, max_row, y_title, x_title, width=14, height=8, legend=False):
    chart = LineChart()
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_col, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    style_chart(chart, color, title, y_title, x_title, width=width, height=height)
    if legend:
        chart.legend.position = "b"
    chart.series[0].graphicalProperties.line.solidFill = color
    chart.series[0].graphicalProperties.line.width = 18000
    ws_dest.add_chart(chart, anchor)
    return chart


def write_stacked_hist(ws, start_col, values, fine_bins, class_series, class_labels):
    values = pd.to_numeric(values, errors="coerce")
    mask = values.notna()
    values = values[mask]
    classes = class_series.reindex(values.index).astype(str)
    ws.cell(1, start_col, "bin")
    for j, lab in enumerate(class_labels):
        ws.cell(1, start_col + 1 + j, str(lab))
    lefts = fine_bins[:-1]
    for i, left in enumerate(lefts, start=2):
        ws.cell(i, start_col, float(left))
    for j, lab in enumerate(class_labels):
        subset = values[classes == str(lab)]
        counts, _ = np.histogram(subset.to_numpy(), bins=fine_bins)
        for i, n in enumerate(counts, start=2):
            ws.cell(i, start_col + 1 + j, int(n))
    return {
        "cat_col": start_col,
        "data_min": start_col + 1,
        "data_max": start_col + len(class_labels),
        "n": len(lefts),
    }


def write_chart_data(ws: Worksheet, df: pd.DataFrame) -> dict:
    """Hidden sheet holding all series used by Excel charts."""
    rsrp_bins = np.arange(-140, -48, 2)
    rsrq_bins = np.arange(-24, -2.5, 0.5)
    sinr_bins = np.arange(-15, 31, 1)

    rsrp_cls = rsrp_map_class(df["RSRP"])
    rsrq_cls = pd.cut(df["RSRQ"], bins=RSRQ_BINS, labels=RSRQ_LABELS, right=False)
    sinr_cls = pd.cut(df["SINR"], bins=SINR_BINS, labels=SINR_LABELS, right=False)

    blocks = {
        "RSRP": write_stacked_hist(ws, 26, df["RSRP"], rsrp_bins, rsrp_cls, MAP_RSRP_LABELS),
        "RSRQ": write_stacked_hist(ws, 35, df["RSRQ"], rsrq_bins, rsrq_cls, RSRQ_LABELS),
        "SINR": write_stacked_hist(ws, 41, df["SINR"], sinr_bins, sinr_cls, SINR_LABELS),
    }

    percentiles = list(range(0, 101, 2))
    ws.cell(1, 7, "Percentile")
    ws.cell(1, 8, "RSRP_CDF")
    ws.cell(1, 9, "RSRQ_CDF")
    ws.cell(1, 10, "SINR_CDF")
    rsrp = df["RSRP"].dropna()
    rsrq = df["RSRQ"].dropna()
    sinr = df["SINR"].dropna()
    rsrp_p = np.nanpercentile(rsrp, percentiles)
    rsrq_p = np.nanpercentile(rsrq, percentiles)
    sinr_p = np.nanpercentile(sinr, percentiles)
    for i, p in enumerate(percentiles, start=2):
        ws.cell(i, 7, p)
        ws.cell(i, 8, float(rsrp_p[i - 2]))
        ws.cell(i, 9, float(rsrq_p[i - 2]))
        ws.cell(i, 10, float(sinr_p[i - 2]))
    blocks["cdf_n"] = len(percentiles)

    ws.cell(1, 12, "RSRP_quality")
    ws.cell(1, 13, "RSRP_quality_n")
    rsrp_q = rsrp_range_counts(df["RSRP"])
    for i, (label, n) in enumerate(rsrp_q.items(), start=2):
        ws.cell(i, 12, str(label))
        ws.cell(i, 13, int(n))
    blocks["rsrp_q_n"] = len(rsrp_q)
    ws.cell(1, 15, "RSRQ_quality")
    ws.cell(1, 16, "RSRQ_quality_n")
    for i, (label, n) in enumerate(quality_counts(df["RSRQ"], RSRQ_BINS, RSRQ_LABELS).items(), start=2):
        ws.cell(i, 15, str(label))
        ws.cell(i, 16, int(n))
    ws.cell(1, 18, "SINR_quality")
    ws.cell(1, 19, "SINR_quality_n")
    for i, (label, n) in enumerate(quality_counts(df["SINR"], SINR_BINS, SINR_LABELS).items(), start=2):
        ws.cell(i, 18, str(label))
        ws.cell(i, 19, int(n))

    ts = (
        df.dropna(subset=["Time"])
        .set_index("Time")[["RSRP", "RSRQ", "SINR"]]
        .resample("10min")
        .mean()
        .dropna(how="all")
    )
    if len(ts) > 900:
        ts = ts.iloc[:: max(1, len(ts) // 900)]
    ws.cell(1, 21, "Time")
    ws.cell(1, 22, "RSRP")
    ws.cell(1, 23, "RSRQ")
    ws.cell(1, 24, "SINR")
    for i, (t, row) in enumerate(ts.iterrows(), start=2):
        cell = ws.cell(i, 21, pd.Timestamp(t).to_pydatetime())
        cell.number_format = "yyyy-mm-dd hh:mm"
        ws.cell(i, 22, None if pd.isna(row["RSRP"]) else float(row["RSRP"]))
        ws.cell(i, 23, None if pd.isna(row["RSRQ"]) else float(row["RSRQ"]))
        ws.cell(i, 24, None if pd.isna(row["SINR"]) else float(row["SINR"]))
    blocks["ts_n"] = len(ts)

    rsrp_vs_bins = list(range(-80, -126, -1))
    pair = df.copy()
    pair["rsrp_bin"] = pd.to_numeric(pair["RSRP"], errors="coerce").round().astype("Int64")
    ws.cell(1, 50, "RSRP_dBm")
    ws.cell(1, 51, "SINR")
    ws.cell(1, 52, "RSRQ")
    all_mean = pair.groupby("rsrp_bin")[["SINR", "RSRQ"]].mean()
    for i, b in enumerate(rsrp_vs_bins, start=2):
        ws.cell(i, 50, int(b))
        if b in all_mean.index:
            ws.cell(i, 51, None if pd.isna(all_mean.loc[b, "SINR"]) else float(all_mean.loc[b, "SINR"]))
            ws.cell(i, 52, None if pd.isna(all_mean.loc[b, "RSRQ"]) else float(all_mean.loc[b, "RSRQ"]))
    blocks["vs_n"] = len(rsrp_vs_bins)
    return blocks


def build_cover(ws: Worksheet, df: pd.DataFrame) -> None:
    banner(
        ws,
        f"  Bogura LTE Drive-Test Report  (v{REPORT_VERSION})",
        f"  Version {REPORT_VERSION}  |  Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Gridlines off  |  Freeze panes off",
        last_col=10,
    )
    set_widths(ws, {"A": 28, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 18, "J": 18})

    write_cell(ws, 4, 1, "Dataset", size=14, bold=True)
    info_rows = [
        ["Report version", f"v{REPORT_VERSION}"],
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["P1 samples", int((df["Source"] == "P1").sum())],
        ["P2 samples", int((df["Source"] == "P2").sum())],
        ["Merged samples", int(len(df))],
        ["Start time", df["Time"].min()],
        ["End time", df["Time"].max()],
        ["Unique Cell Ids", int(df["Cell Id"].nunique())],
        ["Longitude range", f"{df['Longitude'].min():.5f} – {df['Longitude'].max():.5f}"],
        ["Latitude range", f"{df['Latitude'].min():.5f} – {df['Latitude'].max():.5f}"],
    ]
    write_table(ws, 5, 1, ["Item", "Value"], info_rows, num_formats={1: "#,##0"})
    ws.cell(11, 2).number_format = "yyyy-mm-dd hh:mm:ss"
    ws.cell(12, 2).number_format = "yyyy-mm-dd hh:mm:ss"

    write_cell(ws, 4, 4, "Radio KPIs (all samples)", size=14, bold=True)
    stats_headers = ["KPI", "Count", "Min", "P5", "Median", "Mean", "P95", "Max"]
    stats_rows = []
    for name in ("RSRP", "RSRQ", "SINR"):
        st = kpi_stats(df[name])
        stats_rows.append(
            [name, st["Count"], st["Min"], st["P5"], st["Median"], st["Mean"], st["P95"], st["Max"]]
        )
    write_table(
        ws,
        5,
        4,
        stats_headers,
        stats_rows,
        num_formats={1: "#,##0", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00", 6: "0.00", 7: "0.00"},
    )

    write_cell(ws, 15, 1, "How to read this workbook", size=14, bold=True)
    notes = [
        "Cover — dataset size, KPI scorecard, band mix, and overview histograms.",
        "RSRP / RSRQ / SINR — statistics table and coverage map (discrete range legend).",
        "RSRP vs KPIs — merged P1+P2 as one SINR line and one RSRQ line vs RSRP.",
        "Sample Log — evenly spaced subset of the merged samples (full 1.07M rows stay in output/Bogura_merged.csv.gz).",
        "P1 is the split archive (part1–part5). P2 is the standalone archive. This report uses the concatenated final file.",
    ]
    for i, text in enumerate(notes):
        ws.merge_cells(start_row=16 + i, start_column=1, end_row=16 + i, end_column=10)
        write_cell(ws, 16 + i, 1, text, wrap=True)
        ws.row_dimensions[16 + i].height = 18

    write_cell(ws, 22, 1, "Band mix (DL EARFCN)", size=14, bold=True)
    band = df["Band"].value_counts()
    band_rows = [[str(k), int(v), v / len(df)] for k, v in band.items()]
    write_table(ws, 23, 1, ["Band", "Samples", "Share"], band_rows, num_formats={1: "#,##0", 2: "0.0%"})

    write_cell(ws, 22, 5, "RSRP ranges (dBm)", size=14, bold=True)
    rsrp_q = rsrp_range_counts(df["RSRP"])
    rsrp_rows = [[str(label), int(n), n / len(df)] for label, n in rsrp_q.items()]
    write_table(
        ws,
        23,
        5,
        ["Range", "Samples", "Share"],
        rsrp_rows,
        num_formats={1: "#,##0", 2: "0.0%"},
    )
    write_cell(
        ws,
        33,
        5,
        "≥ -85 holds samples stronger than -85. Other bins: lower ≤ RSRP < upper (<-125 is RSRP < -125).",
        size=9,
        wrap=True,
    )
    ws.merge_cells(start_row=33, start_column=5, end_row=33, end_column=8)
    # Histograms are added after _ChartData exists; see add_cover_charts().


def build_kpi_sheet(ws, df, name, color, unit, map_path: Path | None):
    banner(ws, f"  {name} report  (v{REPORT_VERSION})", f"  Unit: {unit}  |  Statistics and coverage map  |  Gridlines off", last_col=10)
    set_widths(ws, {get_column_letter(i): 16 for i in range(1, 11)})
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18

    write_cell(ws, 4, 1, f"{name} statistics", size=14, bold=True)
    st = kpi_stats(df[name])
    write_table(
        ws,
        5,
        1,
        ["Metric", "Value"],
        [[k, st[k]] for k in ("Count", "Min", "P5", "Median", "Mean", "P95", "Max", "Std")],
        header_fill=color,
        num_formats={1: "0.00"},
    )
    ws.cell(6, 2).number_format = "#,##0"
    if map_path and map_path.exists():
        write_cell(ws, 4, 4, f"{name} coverage map", size=14, bold=True)
        add_image(ws, map_path, "D5", width=780, height=580)


def add_vs_line_chart(ws_data, ws_dest, anchor, title, cat_col, data_min, data_max, min_row, max_row, y_title, colors, width=16, height=8):
    chart = LineChart()
    chart.style = 10
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = "RSRP (dBm)"
    chart.width = width
    chart.height = height
    chart.legend.position = "b"
    cats = Reference(ws_data, min_col=cat_col, min_row=min_row, max_row=max_row)
    data = Reference(ws_data, min_col=data_min, max_col=data_max, min_row=min_row - 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    for i, color in enumerate(colors):
        if i >= len(chart.series):
            break
        hexcol = color.lstrip("#")
        chart.series[i].graphicalProperties.line.solidFill = hexcol
        chart.series[i].graphicalProperties.line.width = 18000
        chart.series[i].marker = Marker(symbol="circle", size=6)
        chart.series[i].marker.graphicalProperties.solidFill = hexcol
    if "SINR" in y_title:
        apply_xy_levels(
            chart, "RSRP (dBm)", y_title,
            y_min=-10, y_max=35, show_all_x_ticks=True, rotate_x=True, y_major_unit=5,
            y_num_fmt="0",
        )
    elif "RSRQ" in y_title:
        apply_xy_levels(
            chart, "RSRP (dBm)", y_title,
            y_min=-22, y_max=0, show_all_x_ticks=True, rotate_x=True, y_major_unit=2,
            y_num_fmt="0.0",
        )
    else:
        apply_xy_levels(chart, "RSRP (dBm)", y_title, show_all_x_ticks=True, rotate_x=True)
    ws_dest.add_chart(chart, anchor)
    return chart


def add_dual_axis_vs_chart(ws_data, ws_dest, anchor, n_rows, width=22, height=9):
    cats = Reference(ws_data, min_col=50, min_row=2, max_row=1 + n_rows)
    sinr = LineChart()
    sinr.style = 10
    sinr.title = "RSRP vs SINR vs RSRQ"
    sinr.y_axis.title = "SINR (dB)"
    sinr.x_axis.title = "RSRP (dBm)"
    sinr.width = width
    sinr.height = height
    sinr.add_data(Reference(ws_data, min_col=51, min_row=1, max_row=1 + n_rows), titles_from_data=True)
    sinr.set_categories(cats)
    sinr.series[0].graphicalProperties.line.solidFill = SINR_COLOR
    sinr.series[0].graphicalProperties.line.width = 20000
    sinr.legend.position = "b"

    rsrq = LineChart()
    rsrq.y_axis.axId = 200
    rsrq.y_axis.title = "RSRQ (dB)"
    rsrq.add_data(Reference(ws_data, min_col=52, min_row=1, max_row=1 + n_rows), titles_from_data=True)
    rsrq.series[0].graphicalProperties.line.solidFill = RSRQ_COLOR
    rsrq.series[0].graphicalProperties.line.width = 20000
    sinr.y_axis.crosses = "min"
    rsrq.y_axis.crosses = "max"
    apply_xy_levels(
        sinr, "RSRP (dBm)", "SINR (dB)",
        y_min=-10, y_max=35, show_all_x_ticks=True, rotate_x=True, y_major_unit=5,
        layout=False, y_num_fmt="0",
    )
    reserve_plot_area(sinr, left=0.12, top=0.14, width=0.74, height=0.58)
    rsrq.y_axis.delete = False
    rsrq.y_axis.axPos = "r"
    rsrq.y_axis.tickLblPos = "nextTo"
    rsrq.y_axis.majorTickMark = "out"
    rsrq.y_axis.minorTickMark = "none"
    rsrq.y_axis.title = "RSRQ (dB)"
    _title_outside(rsrq.y_axis.title)
    rsrq.y_axis.scaling.min = -22
    rsrq.y_axis.scaling.max = 0
    rsrq.y_axis.numFmt = "0.0"
    try:
        rsrq.y_axis.majorUnit = 2
    except Exception:
        pass
    sinr += rsrq
    ws_dest.add_chart(sinr, anchor)


def save_rsrp_scatter(df: pd.DataFrame, ycol: str, path: Path, ylabel: str, title: str, ylim=None) -> Path:
    sample = downsample(df.dropna(subset=["RSRP", ycol]), 18000)
    fig, ax = plt.subplots(figsize=(9.2, 6.2))
    ax.scatter(sample["RSRP"], sample[ycol], s=6, alpha=0.22, c="#1F618D", linewidths=0)
    x = sample["RSRP"].to_numpy()
    y = sample[ycol].to_numpy()
    if len(x) > 20:
        coeff = np.polyfit(x, y, 1)
        xs = np.linspace(np.nanmin(x), np.nanmax(x), 80)
        ax.plot(xs, np.polyval(coeff, xs), color="#17202A", lw=1.8, label="Trend")
        ax.legend(fontsize=8)
    ax.set_xlabel("RSRP (dBm)")
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.set_xlim(-50, -130)
    if ylim:
        ax.set_ylim(*ylim)
    ax.xaxis.set_major_locator(MultipleLocator(5))
    ax.yaxis.set_major_locator(MultipleLocator(5 if "SINR" in ylabel else 2))
    ax.tick_params(axis="both", labelsize=8)
    ax.grid(True, alpha=0.35)
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white")
    plt.close(fig)
    return path


def save_binned_vs_preview(df: pd.DataFrame, path: Path) -> Path:
    pair = df.dropna(subset=["RSRP", "SINR"]).copy()
    pair["rsrp_bin"] = pair["RSRP"].round().astype(int)
    fig, ax = plt.subplots(figsize=(11, 5.5))
    g = pair.groupby("rsrp_bin")["SINR"].mean().sort_index(ascending=False)
    g = g.loc[(g.index <= -80) & (g.index >= -125)]
    ax.plot(g.index, g.values, marker="o", ms=4, color=f"#{SINR_COLOR}", label="SINR")
    ax.set_title("RSRP vs SINR")
    ax.set_xlabel("RSRP (dBm)")
    ax.set_ylabel("SINR (dB)")
    ax.set_xlim(-80, -125)
    ax.set_ylim(-10, 35)
    ax.set_xticks(range(-80, -126, -5))
    ax.yaxis.set_major_locator(MultipleLocator(5))
    ax.tick_params(axis="x", rotation=90, labelsize=7)
    ax.tick_params(axis="y", labelsize=8)
    ax.legend()
    ax.grid(True, alpha=0.35)
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white")
    plt.close(fig)
    return path


def build_rsrp_vs_sheet(ws, data_ws, n_rows: int, scatter_sinr: Path, scatter_rsrq: Path) -> None:
    banner(
        ws,
        "  RSRP vs SINR / RSRQ",
        f"  RSRP on the horizontal axis  |  Version {REPORT_VERSION}  |  Binned mean lines and scatter with trend  |  Gridlines off",
        last_col=12,
    )
    set_widths(ws, {get_column_letter(i): 14 for i in range(1, 13)})
    write_cell(
        ws,
        4,
        1,
        "Binned 1 dBm means of the merged P1+P2 log (one SINR line, one RSRQ line). Worse RSRP is to the right. Scatter plots below include a trend line.",
        wrap=True,
    )
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=12)
    add_vs_line_chart(
        data_ws, ws, "A6", "RSRP vs SINR",
        50, 51, 51, 2, 1 + n_rows, "SINR (dB)",
        [SINR_COLOR], width=15, height=8,
    )
    add_vs_line_chart(
        data_ws, ws, "I6", "RSRP vs RSRQ",
        50, 52, 52, 2, 1 + n_rows, "RSRQ (dB)",
        [RSRQ_COLOR], width=15, height=8,
    )
    add_dual_axis_vs_chart(data_ws, ws, "A24", n_rows, width=24, height=10)
    write_cell(ws, 40, 1, "RSRP vs SINR scatter (trend line)", size=14, bold=True)
    write_cell(ws, 40, 8, "RSRP vs RSRQ scatter (trend line)", size=14, bold=True)
    if scatter_sinr.exists():
        add_image(ws, scatter_sinr, "A42", width=520, height=350)
    if scatter_rsrq.exists():
        add_image(ws, scatter_rsrq, "H42", width=520, height=350)


def build_sample_log(ws: Worksheet, df: pd.DataFrame, n: int = 8000) -> None:
    banner(ws, "  Sample drive log", "  Evenly spaced subset for inspection  |  Full merged file is CSV  |  No freeze / no gridlines", last_col=10)
    sample = downsample(df, max_points=n).copy()
    cols = ["Time", "RSRP", "RSRQ", "SINR", "Longitude", "Latitude", "Cell Id", "DL EARFCN", "Source", "Band"]
    sample = sample[cols]
    for c, w in zip(cols, (22, 12, 12, 12, 14, 14, 14, 14, 10, 22)):
        ws.column_dimensions[get_column_letter(cols.index(c) + 1)].width = w
    headers = cols
    for i, h in enumerate(headers, start=1):
        write_cell(ws, 4, i, h, bold=True, color=WHITE, fill_color=NAVY, align="center")
    for r_i, row in enumerate(sample.itertuples(index=False), start=5):
        bg = LIGHT if r_i % 2 == 0 else WHITE
        values = list(row)
        for c_i, value in enumerate(values, start=1):
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            if hasattr(value, "item"):
                try:
                    value = value.item()
                except Exception:
                    pass
            if pd.isna(value):
                value = None
            cell = write_cell(ws, r_i, c_i, value, fill_color=bg, align="center")
            cell.border = THIN
            if c_i == 1 and value is not None:
                cell.number_format = "yyyy-mm-dd hh:mm:ss.000"
            if c_i in (2, 3, 4, 5, 6):
                cell.number_format = "0.000"


def _chart_title_text(title) -> str | None:
    if title is None:
        return None
    try:
        return str(title.tx.rich.p[0].r[0].t)
    except Exception:
        return str(title)


def verify_report(path: Path) -> dict:
    wb = load_workbook(path)
    info = {"sheets": {}, "chart_titles": [], "axis_check": [], "problems": []}
    for name in wb.sheetnames:
        ws = wb[name]
        freeze = ws.freeze_panes
        grid = ws.sheet_view.showGridLines
        print_grid = ws.print_options.gridLines
        n_charts = len(ws._charts)
        n_images = len(ws._images)
        info["sheets"][name] = {
            "freeze_panes": freeze,
            "showGridLines": grid,
            "print_gridLines": print_grid,
            "charts": n_charts,
            "images": n_images,
        }
        if freeze not in (None, "A1"):
            info["problems"].append(f"{name}: freeze_panes={freeze}")
        if grid is not False:
            info["problems"].append(f"{name}: showGridLines={grid}")
        if print_grid:
            info["problems"].append(f"{name}: print gridlines on")
        for ch in ws._charts:
            title = _chart_title_text(getattr(ch, "title", None))
            x_title = _chart_title_text(getattr(ch.x_axis, "title", None))
            y_title = _chart_title_text(getattr(ch.y_axis, "title", None))
            x_pos = getattr(ch.x_axis, "axPos", None)
            y_pos = getattr(ch.y_axis, "axPos", None)
            info["chart_titles"].append((name, title, type(ch).__name__))
            info["axis_check"].append((name, title, x_title, y_title, x_pos, y_pos))
            if x_pos != "b":
                info["problems"].append(f"{name}/{title}: x axPos={x_pos} (expected b)")
            if y_pos not in ("l", "r"):
                info["problems"].append(f"{name}/{title}: y axPos={y_pos} (expected l/r)")
            if not x_title:
                info["problems"].append(f"{name}/{title}: missing horizontal axis title")
            if not y_title:
                info["problems"].append(f"{name}/{title}: missing vertical axis title")
    wb.close()
    with zipfile.ZipFile(path) as zf:
        for item in zf.namelist():
            if not item.startswith("xl/charts/chart"):
                continue
            xml = zf.read(item).decode("utf-8")
            for cat in re.findall(r"<catAx>.*?</catAx>", xml):
                pos = re.search(r'<axPos val="([^"]+)"', cat)
                if not pos or pos.group(1) != "b":
                    info["problems"].append(f"{item}: category axis axPos is {pos.group(1) if pos else 'missing'} (need b)")
                if 'tickLblPos val="none"' in cat:
                    info["problems"].append(f"{item}: category tick labels hidden")
            for val in re.findall(r"<valAx>.*?</valAx>", xml):
                pos = re.search(r'<axPos val="([^"]+)"', val)
                if not pos or pos.group(1) not in ("l", "r"):
                    info["problems"].append(f"{item}: value axis axPos is {pos.group(1) if pos else 'missing'}")
                if 'tickLblPos val="none"' in val:
                    info["problems"].append(f"{item}: value tick labels hidden")
            if "<a:t></a:t>" in xml or "<a:t/>" in xml:
                info["problems"].append(f"{item}: empty text run would blank axis labels")
    return info


def add_cover_charts(cover, data_ws, blocks) -> None:
    write_cell(cover, 35, 1, "RSRP, RSRQ and SINR plots (native Excel charts with legends)", size=14, bold=True)
    add_stacked_col_chart(
        data_ws, cover, "A37", "RSRP histogram",
        blocks["RSRP"]["cat_col"], blocks["RSRP"]["data_min"], blocks["RSRP"]["data_max"],
        2, 1 + blocks["RSRP"]["n"],
        "Samples", "RSRP (dBm)", MAP_RSRP_COLORS, width=14, height=8,
    )
    add_stacked_col_chart(
        data_ws, cover, "H37", "RSRQ histogram",
        blocks["RSRQ"]["cat_col"], blocks["RSRQ"]["data_min"], blocks["RSRQ"]["data_max"],
        2, 1 + blocks["RSRQ"]["n"],
        "Samples", "RSRQ (dB)", RSRQ_STACK_COLORS, width=14, height=8,
    )
    add_stacked_col_chart(
        data_ws, cover, "A56", "SINR histogram",
        blocks["SINR"]["cat_col"], blocks["SINR"]["data_min"], blocks["SINR"]["data_max"],
        2, 1 + blocks["SINR"]["n"],
        "Samples", "SINR (dB)", SINR_STACK_COLORS, width=14, height=8,
    )
    add_col_chart(
        data_ws, cover, "H56", "RSRP ranges (dBm)", RSRP_COLOR,
        12, 13, 2, 1 + blocks["rsrp_q_n"],
        "Samples", "RSRP range", width=14, height=8, show_legend=True,
    )


def save_report_preview(df: pd.DataFrame, path: Path) -> Path:
    fig, axes = plt.subplots(1, 3, figsize=(16, 5.2))
    rsrp_bins = np.arange(-140, -48, 2)
    cls = rsrp_map_class(df["RSRP"])
    bottom = np.zeros(len(rsrp_bins) - 1)
    widths = np.diff(rsrp_bins)
    for lab, color in zip(MAP_RSRP_LABELS, MAP_RSRP_COLORS):
        counts, _ = np.histogram(df.loc[cls == lab, "RSRP"].dropna(), bins=rsrp_bins)
        axes[0].bar(
            rsrp_bins[:-1],
            counts,
            width=widths,
            align="edge",
            bottom=bottom,
            color=color,
            edgecolor="none",
            label=lab,
        )
        bottom += counts
    axes[0].legend(title="RSRP (dBm)", fontsize=7, loc="upper right")
    axes[0].set_title("RSRP histogram")
    axes[0].set_xlabel("RSRP (dBm)")
    axes[0].set_ylabel("Samples")
    axes[0].set_xlim(-140, -50)
    axes[0].grid(False)

    for ax, col, bins, xlim, color, unit in (
        (axes[1], "RSRQ", np.arange(-24, -2.5, 0.5), (-24, -3), RSRQ_COLOR, "dB"),
        (axes[2], "SINR", np.arange(-15, 31, 1), (-15, 30), SINR_COLOR, "dB"),
    ):
        ax.hist(df[col].dropna(), bins=bins, color=f"#{color}", edgecolor="none", label=f"{col} samples")
        ax.legend(fontsize=8)
        ax.set_title(f"{col} histogram")
        ax.set_xlabel(f"{col} ({unit})")
        ax.set_ylabel("Samples")
        ax.set_xlim(*xlim)
        ax.grid(False)
        ax.set_facecolor("white")
    fig.suptitle("Bogura Drive-Test Report — RSRP / RSRQ / SINR", fontsize=13, fontweight="bold")
    fig.tight_layout()
    fig.savefig(path, dpi=140, facecolor="white")
    plt.close(fig)
    return path


def build_report(df: pd.DataFrame, out_path: Path = REPORT_XLSX) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    PLOTS_DIR.mkdir(parents=True, exist_ok=True)
    if "Band" not in df.columns:
        df = df.copy()
        df["Band"] = df["DL EARFCN"].map(earfcn_to_band)

    rsrp_map = plot_rsrp_coverage_map(df, PLOTS_DIR / "excel_route_rsrp.png")
    rsrq_map = plot_rsrq_coverage_map(df, PLOTS_DIR / "excel_route_rsrq.png")
    sinr_map = plot_sinr_coverage_map(df, PLOTS_DIR / "excel_route_sinr.png")
    scatter_sinr = save_rsrp_scatter(df, "SINR", PLOTS_DIR / "excel_scatter_rsrp_sinr.png", "SINR (dB)", "RSRP vs SINR", ylim=(-10, 32))
    scatter_rsrq = save_rsrp_scatter(df, "RSRQ", PLOTS_DIR / "excel_scatter_rsrp_rsrq.png", "RSRQ (dB)", "RSRP vs RSRQ", ylim=(-24, 0))
    save_binned_vs_preview(df, PLOTS_DIR / "excel_rsrp_vs_sinr_lines.png")
    save_report_preview(df, PLOTS_DIR / "excel_rsrp_rsrq_sinr_histograms.png")

    wb = Workbook()
    cover = wb.active
    cover.title = "Cover"
    rsrp_ws = wb.create_sheet("RSRP")
    rsrq_ws = wb.create_sheet("RSRQ")
    sinr_ws = wb.create_sheet("SINR")
    time_ws = wb.create_sheet("RSRP vs KPIs")
    sample_ws = wb.create_sheet("Sample Log")
    data_ws = wb.create_sheet("_ChartData")
    data_ws.sheet_state = "hidden"

    blocks = write_chart_data(data_ws, df)
    build_cover(cover, df)
    add_cover_charts(cover, data_ws, blocks)
    build_kpi_sheet(rsrp_ws, df, "RSRP", RSRP_COLOR, "dBm", rsrp_map)
    build_kpi_sheet(rsrq_ws, df, "RSRQ", RSRQ_COLOR, "dB", rsrq_map)
    build_kpi_sheet(sinr_ws, df, "SINR", SINR_COLOR, "dB", sinr_map)
    build_rsrp_vs_sheet(time_ws, data_ws, blocks["vs_n"], scatter_sinr, scatter_rsrq)
    build_sample_log(sample_ws, df)

    for ws in wb.worksheets:
        apply_sheet_view(ws)

    wb.properties.title = f"Bogura LTE Drive-Test Report v{REPORT_VERSION}"
    wb.properties.subject = f"Version {REPORT_VERSION}"
    wb.properties.description = (
        f"Bogura drive-test report v{REPORT_VERSION}. "
        "KPI sheets: statistics and coverage map. Analysis: RSRP vs SINR/RSRQ."
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    shutil.copy2(out_path, VERSIONED_XLSX)
    print(f"Wrote Excel report: {out_path}")
    print(f"Wrote versioned copy: {VERSIONED_XLSX}")
    return out_path


if __name__ == "__main__":
    from merge_and_analyze import MERGED_CSV, extract_archives, load_merged_csv, merge_parts

    if MERGED_CSV.exists():
        frame = load_merged_csv()
    else:
        extract_archives()
        frame = merge_parts()
    path = build_report(frame)
    result = verify_report(path)
    print(result)
    if result["problems"]:
        raise SystemExit("Excel view checks failed: " + "; ".join(result["problems"]))
    print("Verified: no freeze panes, no gridlines, RSRP/RSRQ/SINR charts present.")
