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

Active auto mode keeps **one band at each latitude and longitude**. Robi rows that share a latitude, longitude and Cell Id collapse to the strongest RSRP, so that cell keeps one band. Then one Cell Id remains at that location, again the strongest RSRP. RSRQ and SINR come from that row. Any other Cell Id at that location is not plotted, and a row with no Cell Id is not plotted. Idle still collapses rows that share a latitude, longitude and timestamp, then keeps the strongest band in each scanner sweep (at most 4 s and 40 m).

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

Active log (replacement `Borga_Active_V2.1`, campaign AUTO DATA): 2,185,081 rows, of which 1,075,048 are Robi and 1,110,033 are the other operator. Robi rows with no Cell Id are not plotted. At each latitude and longitude one Cell Id and one band remain, the strongest RSRP: 280,640 plotted rows. An Active L1800 map is blank wherever L2100 (or another layer) was stronger at that same coordinate. RSRQ and SINR come from that same row.

The Idle all-band coverage maps show the **best server per scanner sweep**. Rows that already share a latitude, longitude and timestamp are reduced to the strongest RSRP first (395,121 of 1,178,276). The scanner still writes the bands at their own millisecond timestamps and revisits each band about every 3 s, so the remaining rows are grouped into sweep windows (at most 4 s and 40 m of the same pass) before the strongest band is taken. Idle band maps are built from those surviving rows, then the same sweep.

RSRP / RSRQ / SINR coverage maps use the Page 1 legend:

**RSRP (dBm):** `-95 to 0` dark green, `-110 to -95` green, `-120 to -110` yellow, `-150 to -120` red.

**RSRQ (dB):** `-10 to 0` green, `-17 to -10` yellow, `-30 to -17` red.

**SINR (dB):** `15 <= X < Max` blue, `10 <= X < 15` dark green, `5 <= X < 10` cyan, `0 <= X < 5` yellow, `-5 <= X < 0` orange, `Min < X < -5` red. SINR is unchanged. The RSRP and RSRQ coverage maps, their Cover histograms, and the RSRP and RSRQ bad-spot sheets use the four-colour and three-colour legends above. Each bad-spot legend sits inside the map on the left and shows that sheet's own sample count. SINR bad-spot sheets keep the previous SINR colours.

Throughput (PDCP DL) and CQI are on the same legend sheet but are not in these IDLE logs, so those maps are not drawn yet.

**RSRP / RSRQ / SINR sheets** show the **combined all-band IDLE map on the left**, then band-wise IDLE maps for L900, L1800, L2100 and L2600. The right side is the same view for **Active Mode (Robi UE)**. These sheets have no bad-spot circles. Every map sits on satellite imagery. Those maps do not overlay site pies.

**Bad Spot RSRP / RSRQ / SINR** are the Idle sheets. **Bad Spot RSRP_Active / RSRQ_Active / SINR_Active** apply the same rules to the Active Robi track. Each sheet puts the **combined all-band view on top with no circles**, then the bad spots one under another: Bad Spot 1, Bad Spot 2, and so on. Each of those maps is large, sits on the left, and leaves the columns to its right empty. Map numbers are omitted.

Active Mode plots **one band per latitude and longitude** (Robi only): the strongest RSRP, and the Cell Id that carried it. RSRQ and SINR follow that row, so L1800 and L2100 are not drawn on top of each other at the same coordinate. Idle collapses the same timestamp first, then keeps the strongest band per scanner sweep.

- Poor is **RSRP < −120 dBm** (red only; −120 is yellow and is not a bad spot), **RSRQ ≤ −20 dB** (inside the red bin; red above −20 is not poor), **SINR ≤ 0 dB** (orange and red)
- Maps and rules read the same **50 m road bins**: a bin is coloured by the median of the samples measured in it and counts as poor only when that median meets the rule above, and poor bins are drawn on top of good ones — so a circle can only sit on colour the map actually shows
- **Every poor run is covered by a circle.** Runs within 2 km join one circle when the poor road inside it spans at most 7 km, so the drawn circle stays within 10 km and covers that road. A longer poor road is split into several circles. Pockets more than 2 km apart stay separate circles even inside the same 10 km neighbourhood
- **The combined view has no circles**, on the Bad Spot sheet and on the RSRP, RSRQ and SINR sheets. Each of the 16 longest bad spots has its own large map, in that same order, and **one orange Excel shape** on that map (named Circle 1, Circle 2, …). The shape is not painted into the picture, so it can be dragged. The columns to the right of each map are left empty
- **Sector marks are only on those individual bad-spot maps.** Cells in `System DUMP.xlsb` are rolled up by the Sector column: one sector name covers many EUtranCellId values. If any sector of a site falls in the view, every sector of that site is drawn. The mark is a transparent cyan outline with no fill, and the Sector name is written at the end of each arm. A yellow distance line is drawn only from a sector whose azimuth points at that bad spot (within 32°). A sector that faces another way has no line
- **Discrete area:** leftover poor bins covering at least 1 km² where that area is at least half poor (dotted circle). With every run already covered, this drive has none
- Each sheet draws the 16 longest circles as separate maps and lists up to 300 in the table, longest first, and states the full count
- Published circle counts, with 0 discrete areas: Idle RSRP 26, RSRQ 45, SINR 204; Active Robi RSRP 81, RSRQ 50, SINR 190. These rules are locked in `LOGIC_LOCK`. `verify_report` fails if a rebuild moves a threshold, a colour, the circle size, or a count

### Cover shows the Active histograms and the DT route (v1.45)

The Cover sheet keeps the Idle dataset, KPI table, and Idle histograms. Under Active Mode it adds three native histograms (RSRP, RSRQ, SINR) for the plotted Robi rows, using the same legend colours as the Idle charts. Below that, drive details are on the left and the DT route is on the right. The route is only the road on satellite imagery: one colour, no legend, and no mode title. RSRP, RSRQ, SINR, bad-spot, and vs-KPI sheets are unchanged from v1.43, including the 96 movable circles.

### Maps are satellite, and each bad-spot circle is its own shape (v1.43)

Every coverage map and bad-spot map is drawn on Esri World Imagery. The combined view is the satellite route only: no circle is painted into it and no circle shape is placed on it. Bad Spot 1, Bad Spot 2, and the rest follow down the left side of the sheet, each a larger map with one movable circle. A site that falls in that view is drawn as a transparent cyan outline, one arm per sector, with the Sector name at the end of the arm and no fill. A circle is at most 10 km across. Where the poor road is in separate pockets, several circles sit inside that 10 km neighbourhood.

## Web coverage map (v1.48)

Same version as the Excel report. Download, unzip, and open `Bogura_DriveTest_Web_v1.48.html`.

**v1.48 web:** https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Web_v1.48.zip

The page opens as one window. Windows can be 1, 2, 3 or 4. One, two and three windows sit in a single row, each with its own controls. Four windows are two by two, and only the top row keeps a control bar. For now the comparison is RSRP; a window can still show RSRQ. Pan and zoom stay together. A spot click marks that place in every window and does not open a popup. Turn on Info, then click a spot, and that window alone shows a small Field and Value table. The table leaves out electrical tilt and mechanical tilt. It has Idle and Active, all bands, and L900 / L1800 / L2100 / L2600. SINR and the bad-spot sheets are not on this page. It needs a network connection: the satellite picture is Esri World Imagery, and the zoom buttons come from Leaflet. Zooming moves the satellite picture and the coloured drive together. Click a spot and the popup shows **CellID** and **CellName** for that spot, plus the RSRP and RSRQ of its 50 m bin. A CellName that is not in the system dump is left blank. The Layer control is a dropdown. Tick more than one band, for example L900 and L1800, and the map draws one coverage trace for those bands, keeping the stronger RSRP in each 50 m bin. Pie widget draws every site along the drive. Sector level shows the sector name at the end of each arm. Site level shows the site name, in a smaller size, at the centre. Ruler follows the pointer after each click. A double-click ends that line and leaves the distances on the map, so the next line does not erase it. Click Ruler to clear every line, or double-click one line to remove only that line. Search finds a site or sector name and moves the map to it. A click also shows Azimuth, Height, E_tilt, M_tilt and RS Power(dBm) from the system dump for that cell.

CellID and CellName come from the drive log and `System DUMP.xlsb`:

- Idle: CellID is the Sector Id column. CellName is the dump column EUtranCellId on the row whose eNBId and cellId match the idle eNB-Id and Sector Id. If several dump rows share that pair, the row with the same DL EARFCN is used.
- Active: CellID is Cell Id mod 256. The eNB id is Cell Id divided by 256, rounded down. CellName is the dump EUtranCellId for that eNBId and cellId, again preferring the matching EARFCN.

Colours match the sheets. RSRP is −95 to 0, −110 to −95, −120 to −110, and −150 to −120. RSRQ is −10 to 0, −17 to −10, and −30 to −17. A spot is one 50 m bin, coloured by the median. Rebuild the page data with `python3 build_web_map.py` after the merged logs or the system dump change.

GitHub shows `index.html` as a file, not as a running map. Download the `web` folder and open `index.html`, or serve that folder with any static file server.

## Download Excel report (v1.48)

**v1.48 (current):** https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.48.xlsx

Latest copy: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report.xlsx

GitHub file page: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/blob/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.48.xlsx

Previous: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/merge-bogura-drive-test-logs-7dfb/output/Bogura_DriveTest_Report_v1.47.xlsx

## Excel report

`output/Bogura_DriveTest_Report_v1.48.xlsx` (also saved as `output/Bogura_DriveTest_Report.xlsx`)

### Bad-spot RSRP and RSRQ sit on satellite imagery (v1.48)

Bad Spot RSRP, Bad Spot RSRQ, and the two Active copies use the same Esri satellite basemap as the RSRP sheet. The legend stays inside the combined view, on the left, with that sheet's sample count in parentheses. An RSRP bad spot is only where the 50 m median is below −120 dBm. RSRQ and SINR rules are unchanged. SINR bad-spot pictures stay as published.

Sheet banners do not include the file version. The filename still carries the version when the report is updated.

| Sheet | Contents |
| --- | --- |
| Cover | Idle dataset, KPI tables, and histograms on top. Active Mode dataset, RSRP / RSRQ / SINR histograms, DT route, and drive details below |
| RSRP / RSRQ / SINR | Idle combined and L900 / L1800 / L2100 / L2600 maps on the left; Active Robi maps on the right |
| Bad Spot RSRP / RSRQ / SINR | Idle bad spots, one sheet per KPI |
| Bad Spot RSRP_Active / RSRQ_Active / SINR_Active | Active Robi bad spots, same rules, one sheet per KPI |
| RSRP vs KPIs | Idle: RSRP vs SINR, RSRP vs RSRQ, dual-axis chart, scatter with trend |
| RSRP vs KPIs_Active | The same chart set for the Active Robi UE |
| Sample Log | Evenly spaced subset of the Idle samples |

Every sheet has freeze panes disabled and worksheet gridlines turned off (screen and print). Charts are native Excel objects, not just pictures.
