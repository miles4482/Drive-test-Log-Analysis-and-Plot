# Drive-test Log Analysis and Plot

LTE **IDLE Mode scanner** and **Active Mode UE** logs from **Bogura** (Robi). Idle source is `Borga_Idle_V2.0.7z` and is unchanged. Active source is `Borga_Active_mode_11Nov26.7z`, which replaces `Borga_Active_mode.csv`. The report draws **Robi** only.

## What these files are

| File(s) | Role | After extract |
| --- | --- | --- |
| `Borga_Idle_V2.0.7z` | IDLE Mode scanner export | `extracted/Borga_Idle_V2.0/1.csv` … `15.csv` |
| `Borga_Active_mode_11Nov26.7z` | Active Mode UE export, 15 CSVs (Robi and Grameenphone) | `output/Bogura_active_merged.csv.gz` |
| `System DUMP.xlsb` | Cell-level site dump. Several EUtranCellId values share one Sector name | Sector name on every arm, plus the distance line, on the individual bad-spot maps only |
| `Physical_Site_Database_V2.0.xlsx` | Earlier physical site database | Not used on the current maps |

Idle scanner columns include `Time`, `Longitude`, `Latitude`, `Cell Id`, `DL EARFCN`, `NB RSRP`, `NB RSRQ`, `NB RS SINR` (mapped to RSRP / RSRQ / SINR). Campaign is **Scanner**. Files 1–15 cover 11–21 May 2026.

Active columns are already named `RSRP`, `RSRQ`, `SINR`, plus `Operator`, `Data Technology`, `DL EARFCN`, `CGI` and `Cell Id`. `Borga_Active_mode_11Nov26` has full timestamps from 11 May 2026 13:01:27 through 21 May 2026 14:07:35. `Operator` contains Robi and Grameenphone. `Type` is M-PCell, M-SCell1 or M-SCell2. The report keeps rows whose operator is exactly **Robi** and whose Type is exactly **M-PCell**. Every Grameenphone row is dropped before a map, a histogram, or a bad spot, including its M-PCell rows. Robi secondary cells are dropped the same way. Cell Id is the Cell Id column, which is the E-UTRAN cell identity. An older export that only has `Home Operator` and `CGI` still loads: Cell Id is then the last number of the CGI.

Active auto mode keeps **one band at each latitude and longitude**: the strongest RSRP. RSRQ and SINR come from that row. Idle still collapses rows that share a latitude, longitude and timestamp, then keeps the strongest band in each scanner sweep (at most 4 s and 40 m).

```
Borga_Idle_V2.0.7z     -->  1.csv … 15.csv               -->  output/Bogura_merged.csv.gz
Borga_Active_mode_11Nov26.7z  -->  1.csv … 15.csv  -->  output/Bogura_active_merged.csv.gz
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
2. Merges the 15 idle CSVs into `output/Bogura_merged.csv.gz`
3. Extracts `Borga_Active_mode_11Nov26.7z` into `extracted/Borga_Active_mode_11Nov26/` and merges those 15 CSVs into `output/Bogura_active_merged.csv.gz`
4. Writes KPI plots under `output/plots/` and a text summary in `output/analysis_summary.md`
5. Writes `output/Bogura_DriveTest_Report.xlsx` — freeze panes off and gridlines off on every sheet

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

Active log (`Borga_Active_mode_11Nov26.7z`, 15 CSVs): 2,185,848 rows, 2,185,081 after exact duplicates are dropped. Robi 1,075,048, Grameenphone 1,110,033. The drawn set is Robi and M-PCell: 553,201 rows, of which 520,596 have a Cell Id. At each latitude and longitude one band remains, the strongest RSRP: 288,308 plotted rows. Plotted mix: L1800 190,026, L2100 89,020, L2600 4,898, L900 4,364. An Active L1800 map is blank wherever another layer was stronger at that same coordinate. RSRQ and SINR come from that same row. The drive runs 2026-05-11 13:01:27 through 2026-05-21 14:07:35. Campaign is AUTO DATA. Idle is still `Borga_Idle_V2.0` and its charts are not retuned.

The Idle all-band coverage maps show the **best server per scanner sweep**. Rows that already share a latitude, longitude and timestamp are reduced to the strongest RSRP first (395,121 of 1,178,276). The scanner still writes the bands at their own millisecond timestamps and revisits each band about every 3 s, so the remaining rows are grouped into sweep windows (at most 4 s and 40 m of the same pass) before the strongest band is taken. Idle band maps are built from those surviving rows, then the same sweep.

RSRP / RSRQ / SINR coverage maps use the Page 1 legend:

**RSRP (dBm):** `-95 to 0` dark green, `-110 to -95` green, `-115 to -110` yellow, `-120 to -115` light red (light red from −115, and −115 and −120 are light red), `-150 to -120` red.

**RSRQ (dB):** `-10 to 0` green, `-17 to -10` yellow, `-30 to -17` red.

**SINR (dB):** `15 <= X < Max` blue, `10 <= X < 15` dark green, `5 <= X < 10` cyan, `0 <= X < 5` yellow, `-5 <= X < 0` orange, `Min < X < -5` red. SINR is unchanged. The RSRP and RSRQ coverage maps, their Cover histograms, and the RSRP and RSRQ bad-spot sheets use the four-colour and three-colour legends above. Each bad-spot legend sits inside the map on the left and shows that sheet's own sample count. SINR bad-spot sheets keep the previous SINR colours.

Throughput (PDCP DL) and CQI are on the same legend sheet but are not in these IDLE logs, so those maps are not drawn yet.

**RSRP / RSRQ / SINR sheets** show the **combined all-band IDLE map on the left**, then band-wise IDLE maps for L900, L1800, L2100 and L2600. The right side is the same view for **Active Mode (Robi UE)**. These sheets have no bad-spot circles. Every map sits on satellite imagery. Those maps do not overlay site pies.

**Bad Spot RSRP / RSRQ / SINR** are the Idle sheets. **Bad Spot RSRP_Active / RSRQ_Active / SINR_Active** apply the same rules to the Active Robi track. Each sheet puts the **combined all-band view on top with no circles**, then the bad spots one under another: Bad Spot 1, Bad Spot 2, and so on. Each of those maps is large and sits on the left. The columns on the right are a suggestion table for that spot. A blank stretch with no plot inside the circle is counted with the bad spot. Map numbers are omitted.

Active Mode plots **one band per latitude and longitude** (Robi only): the strongest RSRP, and the Cell Id that carried it. RSRQ and SINR follow that row, so L1800 and L2100 are not drawn on top of each other at the same coordinate. Idle collapses the same timestamp first, then keeps the strongest band per scanner sweep.

- Poor is **RSRP < −120 dBm** (red only; −120 is yellow and is not a bad spot), **RSRQ ≤ −20 dB** (inside the red bin; red above −20 is not poor), **SINR ≤ 0 dB** (orange and red)
- Maps and rules read the same **50 m road bins**: a bin is coloured by the median of the samples measured in it and counts as poor only when that median meets the rule above, and poor bins are drawn on top of good ones — so a circle can only sit on colour the map actually shows
- **Every poor run is covered by a circle.** Runs within 2 km join one circle when the poor road inside it spans at most 7 km, so the drawn circle stays within 10 km and covers that road. A longer poor road is split into several circles. Pockets more than 2 km apart stay separate circles even inside the same 10 km neighbourhood
- **The combined view has no circles**, on the Bad Spot sheet and on the RSRP, RSRQ and SINR sheets. Each of the 16 longest bad spots has its own large map, in that same order, and **one orange Excel shape** on that map (named Circle 1, Circle 2, …). The shape is not painted into the picture, so it can be dragged. The columns to the right of each map are the suggestion table. A driven stretch with no plot inside the circle is counted with that bad spot
- **Sector marks are only on those individual bad-spot maps.** Cells in `System DUMP.xlsb` are rolled up by the Sector column: one sector name covers many EUtranCellId values. If any sector of a site falls in the view, every sector of that site is drawn. The mark is a transparent cyan outline with no fill, and the Sector name is written at the end of each arm. A yellow dashed line starts at the end of the serving sector, not at the middle of the site, and runs to each red sample inside the orange circle. That cell is the Cell Id on the sample, and it has to be an L900, L1800, L2100 or L2600 cell in the dump. Every serving cell is drawn. The length beside the red is the full red stretch, and a short gap of good coverage stays in that length.
- **Discrete area:** leftover poor bins covering at least 1 km² where that area is at least half poor (dotted circle). With every run already covered, this drive has none
- Each sheet draws the 16 longest circles as separate maps and lists up to 300 in the table, longest first, and states the full count
- Published circle counts, with 0 discrete areas: Idle RSRP 26, RSRQ 45, SINR 204; Active Robi M-PCell on `Borga_Active_mode_11Nov26.7z` RSRP 99, RSRQ 41, SINR 186. These rules are locked in `LOGIC_LOCK`. `verify_report` fails if a rebuild moves a threshold, a colour, the circle size, or a count

### Cover shows the Active histograms and the DT route (v1.45)

The Cover sheet keeps the Idle dataset, KPI table, and Idle histograms. Under Active Mode it adds three native histograms (RSRP, RSRQ, SINR) for the plotted Robi rows, using the same legend colours as the Idle charts. Below that, drive details are on the left and the DT route is on the right. The route is only the road on satellite imagery: one colour, no legend, and no mode title. RSRP, RSRQ, SINR, bad-spot, and vs-KPI sheets are unchanged from v1.43, including the 96 movable circles.

### Maps are satellite, and each bad-spot circle is its own shape (v1.43)

Every coverage map and bad-spot map is drawn on Esri World Imagery. The combined view is the satellite route only: no circle is painted into it and no circle shape is placed on it. Bad Spot 1, Bad Spot 2, and the rest follow down the left side of the sheet, each a larger map with one movable circle. A site that falls in that view is drawn as a transparent cyan outline, one arm per sector, with the Sector name at the end of the arm and no fill. A circle is at most 10 km across. Where the poor road is in separate pockets, several circles sit inside that 10 km neighbourhood.

## Web coverage map (v1.64)

Download, unzip, and open `Bogura_DriveTest_Web_v1.64.html`. Site names are white with a dark edge so they stay readable on the satellite picture. Sector names are a little smaller. With KPI set to Blank, turn on Info and click a sector or the site centre to open the cell table. The Excel workbook stays v1.63.

**v1.64 web:** https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/active-mode-11nov26-6bd5/output/Bogura_DriveTest_Web_v1.64.zip

Previous web: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/active-mode-11nov26-6bd5/output/Bogura_DriveTest_Web_v1.63.zip

The page opens as one window. Windows can be 1, 2, 3 or 4. One, two and three windows sit in a single row, each with its own controls. Four windows are two by two. The top row keeps the full control bar. The bottom row keeps Map, KPI, Mode and Band, so those maps can be set for the comparison. Map chooses Satellite, Google map, Google satellite, Terrain, Topographic or Street. The coloured plot marks are smaller than on v1.56. For now the comparison is RSRP; a window can still show RSRQ. Pan and zoom stay together. A spot click marks that place in every window and does not open a popup. Turn on Info, then click a spot, and that window alone shows a small Field and Value table. The table includes E_tilt, M_tilt, BW (dlChannelBandwidth, shown in MHz), Tx (noOfTxAntennas) and Rx (noOfRxAntennas). KPI can be Blank, which leaves the satellite map clear for general use. It has Idle and Active, all bands, and L900 / L1800 / L2100 / L2600. SINR and the bad-spot sheets are not on this page. It needs a network connection: the default picture is Esri World Imagery, the other map choices come from Google, OpenTopoMap and OpenStreetMap, and the zoom buttons come from Leaflet. Zooming moves the satellite picture and the coloured drive together. Click a spot and the popup shows **CellID** and **CellName** for that spot, plus the RSRP and RSRQ of its 50 m bin. A CellName that is not in the system dump is left blank. The Layer control is a dropdown. Tick more than one band, for example L900 and L1800, and the map draws one coverage trace for those bands, keeping the stronger RSRP in each 50 m bin. Pie widget draws every site from `System DUMP.xlsb` whose centre is inside the chart. Sector level shows the sector name as plain light text at the end of each arm, a little smaller than before, with no filled background. Site level shows the site name in white with a dark edge, at the centre, so it stays readable on the satellite picture. With KPI set to Blank, Info plus a click on a sector or on the site centre opens the same cell table used for a plotted spot: CellID, CellName, azimuth, height, E_tilt, M_tilt, BW, Tx, Rx and RS Power. RSRP and RSRQ stay blank because Blank draws no samples. Ruler follows the pointer after each click. A double-click ends that line and leaves the distances on the map, so the next line does not erase it. Click Ruler to clear every line, or double-click one line to remove only that line. Search finds a site or sector name and moves the map to it. Lat-long, beside Search, moves the map to a typed latitude and longitude, and the box shows the position after each move. A click also shows Azimuth, Height, E_tilt, M_tilt, BW, Tx, Rx and RS Power(dBm) from the system dump for that cell.

CellID and CellName come from the drive log and `System DUMP.xlsb`:

- Idle: CellID is the Sector Id column. CellName is the dump column EUtranCellId on the row whose eNBId and cellId match the idle eNB-Id and Sector Id. If several dump rows share that pair, the row with the same DL EARFCN is used.
- Active: CellID is Cell Id mod 256. The eNB id is Cell Id divided by 256, rounded down. CellName is the dump EUtranCellId for that eNBId and cellId, again preferring the matching EARFCN.

Colours match the sheets. RSRP is −95 to 0, −110 to −95, −120 to −110, and −150 to −120. RSRQ is −10 to 0, −17 to −10, and −30 to −17. A spot is one 50 m bin, coloured by the median. Rebuild the page data with `python3 build_web_map.py` after the merged logs or the system dump change.

GitHub shows `index.html` as a file, not as a running map. Download the `web` folder and open `index.html`, or serve that folder with any static file server.

## Download Excel report (v1.63)

**v1.63 (current):** https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/active-mode-11nov26-6bd5/output/Bogura_DriveTest_Report_v1.63.xlsx

Latest copy: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/active-mode-11nov26-6bd5/output/Bogura_DriveTest_Report.xlsx

GitHub file page: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/blob/cursor/active-mode-11nov26-6bd5/output/Bogura_DriveTest_Report_v1.63.xlsx

Previous: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/raw/cursor/active-mode-11nov26-6bd5/output/Bogura_DriveTest_Report_v1.62.xlsx

The web map is v1.63 as well.

## Excel report

`output/Bogura_DriveTest_Report_v1.63.xlsx` (also saved as `output/Bogura_DriveTest_Report.xlsx`)

### Connected sector marks, summary beside the route (v1.63)

Sector marks are the connected Y again. On the Cover sheet, Bad Spot Summary across Mode sits to the right of the DT route, with Idle and Active counts together. The count table is no longer on the bad-spot sheets.

### Every bad spot, one cover table (v1.62)

Idle and Active each draw a map for every bad spot: 26 idle maps and 99 active maps. The Cover sheet has one BAD Spot Summary table for both modes, with the sector and the serving cell on each row. Sector marks are transparent and the arms no longer meet at the centre.

### 3 km fourth-sector reach, cover summary, and EP Details (v1.61)

A 4th L900 sector replaces a new site only when a host is inside 3 km. Past 3 km the final step is a new 3-sector site. The Cover sheet lists the Idle and Active bad-spot counts and the final solution for every spot. The EP Details sheet is the engineering-parameter tracker: one row per sector or cell, with the existing height, azimuth, E_tilt and M_tilt beside the proposed values. Sector marks on the map are lighter, and the sector and site names have no filled background.

### Fewer new sites, and downtilt limited to 2° (v1.60)

A new 3-sector site was kept only when the nearest existing site was farther than 4 km. Inside 4 km, a 4th L900 sector was aimed at the bad spot instead. Downtilt of an existing sector is limited to 2°. When the spot needs more downtilt, the final step is a 4th sector on a site within 1 km. Uptilt is unchanged. The web map adds Satellite, Google map, Google satellite, Terrain, Topographic and Street, and the plot marks are smaller.

### Scale bar, and the solution aimed at the red road (v1.59)

Every coverage map, the DT route, and every bad-spot map has a scale bar. The sector choice is aimed at the red samples. A blank inside the circle is still counted with that spot, and it no longer steers the tilt onto a sector that faces the blank instead of the red road. The findings name the serving cells on those red samples.

### Proposed site on the map, and a solution count (v1.58)

Each RSRP bad-spot map labels the yellow dashed line from a serving sector to the bad spot with that distance in metres. When the final solution is a new 3-sector site, that site is drawn in purple, and the sector toward the bad spot is long enough to reach it. When the final solution is a 4th L900 sector, that sector is drawn in orange from the existing site to the bad spot. Above the identified-spots table, one summary counts tilt and azimuth, 4th sector, and new site.

### RSRP bad spots only, with a final solution on every row (v1.57)

The workbook no longer has Bad Spot RSRQ, Bad Spot SINR, Bad Spot RSRQ_Active, or Bad Spot SINR_Active. Idle and Active each keep the RSRP bad-spot sheet. Every bad spot is listed, not only the 16 maps. The table adds Site, Population, and Final solution on the right.

A blank with no plot inside the circle is still part of that bad spot. The final step follows the distance. Inside about 2.5 km, azimuth and tilt of a facing sector can cover the spot, and the turn is kept only when WorldPop 2020 still serves the customers on the old bearing. A spot in the null between two sectors and within 1.5 km gets a new L900 sector aimed through that null. One or two sites inside 2.5 km use that 4th sector when the spot is in their null and the site is inside 2 km. Several nearby sites are tuned when one of them already faces the spot; otherwise a 4th sector inside 2 km is used. A 4th sector past 2 km is not proposed. If none of those cover the spot, the final step is a new 3-sector L900 site with an azimuth, E_tilt, and M_tilt for each sector.

### Serving-cell lines on the bad-spot maps (v1.53)

Each bad-spot map draws a yellow dashed line from the end of the serving sector to the red samples inside the orange circle. The cell is the Cell Id on L900, L1800, L2100 or L2600 in `System DUMP.xlsb`. The length beside the red is the full red stretch, and a short gap of good coverage stays in that length. Coverage sheets are the v1.52 maps.

### Active Mode uses the 11 Nov 2026 log (v1.52)

The Active maps, histograms, bad spots, DT route and the web map read `Borga_Active_mode_11Nov26.7z`, keeping operator **Robi** and Type **M-PCell**. Idle stays on `Borga_Idle_V2.0` with the same charts and the same circle counts (RSRP 26, RSRQ 45, SINR 204). Active bad-spot circles are RSRP 99, RSRQ 41 and SINR 186.

### RSRP legend has light red from −115 (v1.51)

RSRP colours are −95 to 0 dark green, −110 to −95 green, −115 to −110 yellow, −120 to −115 light red, and −150 to −120 red. −115 and −120 are light red. A bad spot is still only the red bin, below −120 dBm.

### Active Mode uses Robi M-PCell rows (v1.50)

The v1.50 Active maps read `Borga_Active_mode.csv`, keeping Home Operator **Robi** and Type **M-PCell**. That short log is replaced in v1.52. Idle stays on `Borga_Idle_V2.0`. The v1.50 Active bad-spot circles were RSRP 2, RSRQ 0 and SINR 6.

### Bad-spot RSRP and RSRQ sit on satellite imagery (v1.48)

Bad Spot RSRP, Bad Spot RSRQ, and the two Active copies use the same Esri satellite basemap as the RSRP sheet. The legend stays inside the combined view, on the left, with that sheet's sample count in parentheses. An RSRP bad spot is only where the 50 m median is below −120 dBm. Light red, from −115 through −120, is not a bad spot. RSRQ and SINR rules are unchanged. SINR bad-spot pictures stay as published.

Sheet banners do not include the file version. The filename still carries the version when the report is updated.

| Sheet | Contents |
| --- | --- |
| Cover | Idle dataset, KPI tables, and histograms on top. Active Mode dataset, RSRP / RSRQ / SINR histograms, DT route, and drive details below |
| RSRP / RSRQ / SINR | Idle combined and L900 / L1800 / L2100 / L2600 maps on the left; Active Robi maps on the right |
| Bad Spot RSRP | Idle RSRP bad spots, with Site, Population and Final solution for every spot |
| Bad Spot RSRP_Active | Active Robi RSRP bad spots, same columns |
| RSRP vs KPIs | Idle: RSRP vs SINR, RSRP vs RSRQ, dual-axis chart, scatter with trend |
| RSRP vs KPIs_Active | The same chart set for the Active Robi UE |
| Sample Log | Evenly spaced subset of the Idle samples |

Every sheet has freeze panes disabled and worksheet gridlines turned off (screen and print). Charts are native Excel objects, not just pictures.
