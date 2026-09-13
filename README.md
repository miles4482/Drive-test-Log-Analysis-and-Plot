# Drive-test Log Analysis and Plot

LTE **IDLE Mode scanner** logs from **Bogura** (Robi). The current source is `Borga_Idle_V2.0.7z`.

## What these files are

| File(s) | Role | After extract |
| --- | --- | --- |
| `Borga_Idle_V2.0.7z` | IDLE Mode scanner export | `extracted/Borga_Idle_V2.0/1.csv` … `15.csv` |
| `Physical_Site_Database_V1.0.xlsb` | Physical site / sector database | Kept in the repo; not overlaid on current maps |

Scanner columns include `Time`, `Longitude`, `Latitude`, `Cell Id`, `DL EARFCN`, `NB RSRP`, `NB RSRQ`, `NB RS SINR` (mapped to RSRP / RSRQ / SINR). Campaign is **Scanner**. Files 1–15 cover 11–21 May 2026.

```
Borga_Idle_V2.0.7z  -->  1.csv … 15.csv
1.csv + … + 15.csv  -->  output/Bogura_merged.csv.gz  (final file)
```

## How to rebuild the final file and plots

Needs `p7zip-full` plus Python packages in `requirements.txt`.

```bash
sudo apt-get install p7zip-full
pip install -r requirements.txt
python merge_and_analyze.py
```

The script:

1. Extracts `Borga_Idle_V2.0.7z` into `extracted/Borga_Idle_V2.0/`
2. Merges the 15 CSVs into `output/Bogura_merged.csv.gz`
3. Writes KPI plots under `output/plots/` and a text summary in `output/analysis_summary.md`
4. Writes `output/Bogura_DriveTest_Report.xlsx` — freeze panes off and gridlines off on every sheet

You can rebuild only the Excel report from the merged CSV:

```bash
python generate_excel_report.py
```

## Plots

After a successful run:

- `output/plots/01_route_rsrp.png` — best-server route colored by RSRP
- `output/plots/01_route_rsrq.png` — best-server route colored by RSRQ
- `output/plots/01_route_sinr.png` — best-server route colored by SINR
- `output/plots/02_kpi_histograms.png` — RSRP / RSRQ / SINR histograms
- `output/plots/03_kpi_vs_time.png` — KPIs over time
- `output/plots/04_rsrp_quality_bins.png` — coverage quality bins
- `output/plots/05_band_counts.png` — LTE band from DL EARFCN
- `output/plots/06_top_cells.png` — most-seen Cell Ids
- `output/plots/07_csv_counts.png` — sample counts by CSV file (`1.csv` … `15.csv`)
- `output/plots/08_bad_spots_rsrp.png` / `_rsrq.png` / `_sinr.png` — bad-spot maps (no site pies)

If `output/Bogura_merged.csv.gz` already exists, skip extract with:

```bash
python merge_and_analyze.py --from-merged
```

## Merged-log snapshot

| Item | Value |
| --- | --- |
| Archive | Borga_Idle_V2.0.7z (15 CSVs) |
| Scanner detections | 1,178,276 (11–21 May 2026) |
| Operator | Robi |
| Layers | L900 / L1800 / L2100 / L2600 |
| Dominant bands | B8 900 (28.4%), B1 2100 (25.4%), B3 1800 (23.5%), B41 2500 (22.7%) |
| Unique Cell Ids | 5,132 |

The all-band coverage maps show the **best server per scanner sweep**. The scanner writes one row per
band with its own millisecond timestamp and revisits each band about every 3 s, so picking the
strongest RSRP at an exact timestamp never lets the bands compete — it just keeps whichever band was
sampled at that instant, and B41 2500 MHz is polled about four times as often as B8 900 MHz. Samples
are therefore grouped into sweep windows (at most 4 s and 40 m of the same pass) before the strongest
band is taken.

RSRP / RSRQ / SINR coverage maps use the Page 1 legend:

**RSRP (dBm):** `-90 <= X < Max` blue, `-100 <= X < -90` dark green, `-110 <= X < -100` light green, `-115 <= X < -110` yellow, `-120 <= X < -115` magenta, `-Min <= X < -120` red.

**RSRQ (dB):** `X >= -5` blue, `-10 <= X < -5` dark green, `-15 <= X < -10` light green, `-20 <= X < -15` yellow, `X < -20` red.

**SINR (dB):** `15 <= X < Max` blue, `10 <= X < 15` dark green, `5 <= X < 10` cyan, `0 <= X < 5` yellow, `-5 <= X < 0` orange, `Min < X < -5` red.

Throughput (PDCP DL) and CQI are on the same legend sheet but are not in these IDLE logs, so those maps are not drawn yet.

**RSRP / RSRQ / SINR sheets** show the **combined all-band IDLE map on top with bad spots marked**, then band-wise IDLE maps for L900, L1800, L2100 and L2600. Those maps do not overlay site pies. Active Mode cells stay blank until those logs are provided.

**Bad Spot RSRP / RSRQ / SINR** are separate sheets (one KPI each). Each sheet puts the **combined all-band view with marks on top**, then zoomed local inspection. Map numbers are omitted.

- **RSRP** < −115 dBm (magenta and red only; yellow is not a bad-spot colour), **RSRQ** < −20 dB (red), **SINR** < 0 dB
- Maps and rules read the same **50 m road bins**: a bin is coloured by the median of the samples measured in it and counts as poor only when that median is below the threshold, and poor bins are drawn on top of good ones — so a circle can only sit on colour the map actually shows
- **Consecutive:** at least 200 m of poor coverage along the drive (thin black dashed oval covering the whole poor stretch)
- **Discrete area:** leftover poor bins covering at least 1 km² where that area is at least half poor (thin black dotted circle)
- Poor patches shorter than 200 m are left unmarked by design; each sheet states how many stretches the rules found and how long the longest one that missed the rule was

## Download Excel report (v1.28)

**v1.28 (current):** https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.28.xlsx

Latest copy: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report.xlsx

GitHub file page: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/blob/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.28.xlsx

## Excel report

`output/Bogura_DriveTest_Report_v1.28.xlsx` (also saved as `output/Bogura_DriveTest_Report.xlsx`)

Sheet banners do not include the file version. The filename still carries the version when the report is updated.

| Sheet | Contents |
| --- | --- |
| Cover | IDLE Mode bar, dataset/KPI tables on the left, histograms on the right, Active Mode bar at the bottom |
| RSRP / RSRQ / SINR | Combined all-band IDLE map on top with bad spots marked, then L900 / L1800 / L2100 / L2600 IDLE maps (no site pies); Active Mode remains a placeholder |
| Bad Spot RSRP / RSRQ / SINR | One sheet per KPI: combined all-band view with marks on top, zoomed local views, and that KPI’s table |
| RSRP vs KPIs | RSRP vs SINR, RSRP vs RSRQ, dual-axis combined chart, and scatter plots with trend |
| Sample Log | Evenly spaced subset of the merged samples |

Every sheet has freeze panes disabled and worksheet gridlines turned off (screen and print). Charts are native Excel objects, not just pictures.
