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

RSRP / RSRQ / SINR coverage maps use the Page 1 legend:

**RSRP (dBm):** `-90 <= X < Max` blue, `-100 <= X < -90` dark green, `-110 <= X < -100` light green, `-115 <= X < -110` yellow, `-120 <= X < -115` magenta, `-Min <= X < -120` red.

**RSRQ (dB):** `X >= -5` blue, `-10 <= X < -5` dark green, `-15 <= X < -10` light green, `-20 <= X < -15` yellow, `X < -20` red.

**SINR (dB):** `15 <= X < Max` blue, `10 <= X < 15` dark green, `5 <= X < 10` cyan, `0 <= X < 5` yellow, `-5 <= X < 0` orange, `Min < X < -5` red.

Throughput (PDCP DL) and CQI are on the same legend sheet but are not in these IDLE logs, so those maps are not drawn yet.

**RSRP / RSRQ / SINR sheets** show the all-band IDLE map, then band-wise IDLE maps for L900, L1800, L2100 and L2600. Those maps do not overlay site pies. Active Mode cells stay blank until those logs are provided.

**Bad Spot Analysis** has no site pies. Poor samples are:

- **RSRP** ≤ −115 dBm, **RSRQ** ≤ −20 dB, **SINR** ≤ 0 dB
- **Consecutive:** at least 200 m of poor coverage along the drive (thin black dashed oval)
- **Discrete area:** leftover poor samples covering at least 1 km² that never form a 200 m consecutive stretch (thin black dotted circle)

Numbered marks match the table.

## Download Excel report (v1.22)

**v1.22 (current):** https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.22.xlsx

Latest copy: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report.xlsx

GitHub file page: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/blob/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.22.xlsx

## Excel report

`output/Bogura_DriveTest_Report_v1.22.xlsx` (also saved as `output/Bogura_DriveTest_Report.xlsx`)

Sheet banners do not include the file version. The filename still carries the version when the report is updated.

| Sheet | Contents |
| --- | --- |
| Cover | IDLE Mode bar, dataset/KPI tables on the left, histograms on the right, Active Mode bar at the bottom |
| RSRP / RSRQ / SINR | All-band IDLE map, then L900 / L1800 / L2100 / L2600 IDLE maps (no site pies); Active Mode remains a placeholder |
| Bad Spot Analysis | First all-band IDLE maps (same as RSRP / RSRQ / SINR sheets) with consecutive 200 m ovals and discrete ≥1 km² circles; no site pies |
| RSRP vs KPIs | RSRP vs SINR, RSRP vs RSRQ, dual-axis combined chart, and scatter plots with trend |
| Sample Log | Evenly spaced subset of the merged samples |

Every sheet has freeze panes disabled and worksheet gridlines turned off (screen and print). Charts are native Excel objects, not just pictures.
