# Drive-test Log Analysis and Plot

LTE **IDLE Mode scanner** and **Active Mode UE** logs from **Bogura** (Robi). Idle source is `Borga_Idle_V2.0.7z`. Active source is the replacement `Borga_Active_V2.1.zip.001` + `Borga_Active_V2.1.zip.002` (campaign AUTO DATA). The report draws **Robi** only.

## What these files are

| File(s) | Role | After extract |
| --- | --- | --- |
| `Borga_Idle_V2.0.7z` | IDLE Mode scanner export | `extracted/Borga_Idle_V2.0/1.csv` … `15.csv` |
| `Borga_Active_V2.1.zip.001` + `.002` | Active Mode UE export (Robi and Grameenphone). Join the two parts before unzip | `extracted/Borga_Active_V2.1/1.csv` … `15.csv` |
| `System DUMP.xlsb` | Cell-level site dump. Several EUtranCellId values share one Sector name | Sector name on every arm, plus the distance line, on the individual bad-spot maps only |
| `Physical_Site_Database_V2.0.xlsx` | Earlier physical site database | Not used on the current maps |

Idle scanner columns include `Time`, `Longitude`, `Latitude`, `Cell Id`, `DL EARFCN`, `NB RSRP`, `NB RSRQ`, `NB RS SINR` (mapped to RSRP / RSRQ / SINR). Campaign is **Scanner**. Files 1–15 cover 11–21 May 2026.

Active columns are already named `RSRP`, `RSRQ`, `SINR`, plus `Operator`, `Data Technology`, `DL EARFCN` and `Cell Id`. Campaign is **AUTO DATA**. The Operator column also contains Grameenphone. The report keeps rows whose Operator is exactly **Robi** (surrounding spaces are stripped) and does not draw Grameenphone. LTE and LTE CA both stay in when the operator is Robi. Those other-operator rows stay in the merged CSV.

Where several rows share a latitude, longitude and timestamp, both modes keep **one row**: the strongest RSRP. RSRQ and SINR on the maps and on the RSRP vs KPIs charts are the values from that same row. On an Active band map, a carrier that lost that comparison is left blank at that spot. Idle still groups the remaining rows into scanner sweeps (at most 4 s and 40 m) and keeps the strongest band in the sweep.

```
Borga_Idle_V2.0.7z              -->  1.csv … 15.csv                         -->  output/Bogura_merged.csv.gz
Borga_Active_V2.1.zip.001+.002  -->  Borga_Active_V2.1/1.csv … 15.csv       -->  output/Bogura_active_merged.csv.gz
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

Active log (replacement `Borga_Active_V2.1`, campaign AUTO DATA): 2,185,081 rows, of which 1,075,048 are Robi and 1,110,033 are the other operator. After keeping the strongest RSRP at each latitude, longitude and timestamp, 561,553 Robi rows are plotted. That winning row has one layer, so an Active L1800 map is blank wherever L2100 (or another layer) was stronger at the same fix. RSRQ and SINR come from that same row.

The Idle all-band coverage maps show the **best server per scanner sweep**. Rows that already share a latitude, longitude and timestamp are reduced to the strongest RSRP first (395,121 of 1,178,276). The scanner still writes the bands at their own millisecond timestamps and revisits each band about every 3 s, so the remaining rows are grouped into sweep windows (at most 4 s and 40 m of the same pass) before the strongest band is taken. Idle band maps are built from those surviving rows, then the same sweep.

RSRP / RSRQ / SINR coverage maps use the Page 1 legend:

**RSRP (dBm):** `-90 <= X < Max` blue, `-100 <= X < -90` dark green, `-110 <= X < -100` light green, `-115 <= X < -110` yellow, `-120 <= X < -115` magenta, `-Min <= X < -120` red.

**RSRQ (dB):** `X >= -5` blue, `-10 <= X < -5` dark green, `-15 <= X < -10` light green, `-20 <= X < -15` yellow, `X < -20` red.

**SINR (dB):** `15 <= X < Max` blue, `10 <= X < 15` dark green, `5 <= X < 10` cyan, `0 <= X < 5` yellow, `-5 <= X < 0` orange, `Min < X < -5` red.

Throughput (PDCP DL) and CQI are on the same legend sheet but are not in these IDLE logs, so those maps are not drawn yet.

**RSRP / RSRQ / SINR sheets** show the **combined all-band IDLE map on the left**, then band-wise IDLE maps for L900, L1800, L2100 and L2600. The right side is the same view for **Active Mode (Robi UE)**. These sheets have no bad-spot circles. Every map sits on satellite imagery. Those maps do not overlay site pies.

**Bad Spot RSRP / RSRQ / SINR** are the Idle sheets. **Bad Spot RSRP_Active / RSRQ_Active / SINR_Active** apply the same rules to the Active Robi track. Each sheet puts the **combined all-band view on top with no circles**, then the bad spots one under another: Bad Spot 1, Bad Spot 2, and so on. Each of those maps is large, sits on the left, and leaves the columns to its right empty. Map numbers are omitted.

Active Mode plots the **strongest RSRP at each latitude, longitude and timestamp** (Robi only). RSRQ and SINR follow that row, so two bands are not drawn on top of each other at the same fix. Idle does the same exact-key collapse, then keeps the strongest band per scanner sweep.

- Poor is **RSRP ≤ −115 dBm** (magenta and red only; yellow is not a bad-spot colour), **RSRQ ≤ −20 dB** (red), **SINR ≤ 0 dB** (orange and red)
- Maps and rules read the same **50 m road bins**: a bin is coloured by the median of the samples measured in it and counts as poor only when that median is at or below the threshold, and poor bins are drawn on top of good ones — so a circle can only sit on colour the map actually shows
- **Every poor run is covered by a circle.** Runs within 2 km join one circle when the poor road inside it spans at most 7 km, so the drawn circle stays within 10 km and covers that road. A longer poor road is split into several circles. Pockets more than 2 km apart stay separate circles even inside the same 10 km neighbourhood
- **The combined view has no circles**, on the Bad Spot sheet and on the RSRP, RSRQ and SINR sheets. Each of the 16 longest bad spots has its own large map, in that same order, and **one orange Excel shape** on that map (named Circle 1, Circle 2, …). The shape is not painted into the picture, so it can be dragged. The columns to the right of each map are left empty
- **Sector marks are only on those individual bad-spot maps.** Cells in `System DUMP.xlsb` are rolled up by the Sector column: one sector name covers many EUtranCellId values. If any sector of a site falls in the view, every sector of that site is drawn. The mark is a transparent cyan outline with no fill, and the Sector name is written at the end of each arm. A yellow line, with the distance, is drawn when a sector shoots toward that bad spot
- **Discrete area:** leftover poor bins covering at least 1 km² where that area is at least half poor (dotted circle). With every run already covered, this drive has none
- Each sheet draws the 16 longest circles as separate maps and lists up to 300 in the table, longest first, and states the full count
- Published circle counts, with 0 discrete areas: Idle RSRP 41, RSRQ 45, SINR 204; Active Robi RSRP 163, RSRQ 52, SINR 190. These rules are locked in `LOGIC_LOCK`. `verify_report` fails if a rebuild moves a threshold, a colour, the circle size, or a count

### Maps are satellite, and each bad-spot circle is its own shape (v1.41)

Every coverage map and bad-spot map is drawn on Esri World Imagery. The combined view is the satellite route only: no circle is painted into it and no circle shape is placed on it. Bad Spot 1, Bad Spot 2, and the rest follow down the left side of the sheet, each a larger map with one movable circle. A site that falls in that view is drawn as a transparent cyan outline, one arm per sector, with the Sector name at the end of the arm and no fill. A circle is at most 10 km across. Where the poor road is in separate pockets, several circles sit inside that 10 km neighbourhood.

## Download Excel report (v1.41)

**v1.41 (current):** https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.41.xlsx

Latest copy: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report.xlsx

GitHub file page: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/blob/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.41.xlsx

## Excel report

`output/Bogura_DriveTest_Report_v1.41.xlsx` (also saved as `output/Bogura_DriveTest_Report.xlsx`)

Sheet banners do not include the file version. The filename still carries the version when the report is updated.

| Sheet | Contents |
| --- | --- |
| Cover | IDLE dataset and KPI tables, histograms, and an Active Mode dataset for the Robi UE |
| RSRP / RSRQ / SINR | Idle combined and L900 / L1800 / L2100 / L2600 maps on the left; Active Robi maps on the right |
| Bad Spot RSRP / RSRQ / SINR | Idle bad spots, one sheet per KPI |
| Bad Spot RSRP_Active / RSRQ_Active / SINR_Active | Active Robi bad spots, same rules, one sheet per KPI |
| RSRP vs KPIs | Idle: RSRP vs SINR, RSRP vs RSRQ, dual-axis chart, scatter with trend |
| RSRP vs KPIs_Active | The same chart set for the Active Robi UE |
| Sample Log | Evenly spaced subset of the Idle samples |

Every sheet has freeze panes disabled and worksheet gridlines turned off (screen and print). Charts are native Excel objects, not just pictures.
