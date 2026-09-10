# Drive-test Log Analysis and Plot

LTE drive-test logs from **Bogura**, split across two capture files that must be joined before analysis.

## What these files are

The repository originally contained only compressed archives. They are **not** six independent datasets. They are two Excel workbooks, one of which was split so it could be uploaded.

| File(s) | Role | After extract |
| --- | --- | --- |
| `Bogura_P1.part1.rar` … `Bogura_P1.part5.rar` | One multi-volume RAR (parts 1–5 of the same archive) | `Bogura_P1.xlsx` (~1.015 million rows) |
| `Bogura_P2.rar` | Standalone RAR | `Bogura_P2.xlsx` (~60 thousand rows) |

P1 and P2 use the **same 8 columns**:

`Time`, `RSRP`, `RSRQ`, `SINR`, `Longitude`, `Latitude`, `Cell Id`, `DL EARFCN`

They are two drive sessions (P1 around 11 May 2026, P2 around 21 May 2026) in the Bogura area. The final log is P1 and P2 concatenated, then sorted by time.

```
Bogura_P1.part1 + part2 + part3 + part4 + part5  -->  Bogura_P1.xlsx
Bogura_P2.rar                                    -->  Bogura_P2.xlsx
Bogura_P1.xlsx  +  Bogura_P2.xlsx                 -->  Bogura_merged.csv.gz  (final file)
```

## How to rebuild the final file and plots

Needs `unrar` plus Python packages in `requirements.txt`.

```bash
sudo apt-get install unrar
pip install -r requirements.txt
python merge_and_analyze.py
```

The script:

1. Extracts the five P1 volumes into `extracted/Bogura_P1.xlsx`
2. Extracts P2 into `extracted/Bogura_P2.xlsx`
3. Merges them into `output/Bogura_merged.csv.gz` (gzip CSV; a full `.xlsx` of 1M+ rows is impractical)
4. Writes KPI plots under `output/plots/` and a text summary in `output/analysis_summary.md`
5. Writes `output/Bogura_DriveTest_Report.xlsx` — Excel report with native RSRP, RSRQ and SINR charts, **freeze panes off** and **gridlines off** on every sheet

You can rebuild only the Excel report from the merged CSV:

```bash
python generate_excel_report.py
```

Split RAR volumes are not concatenated with `cat`. `unrar` reads `Bogura_P1.part1.rar` and automatically consumes part2–part5.

## Plots

After a successful run:

- `output/plots/01_route_rsrp.png` — drive route colored by RSRP
- `output/plots/02_kpi_histograms.png` — RSRP / RSRQ / SINR histograms
- `output/plots/03_kpi_vs_time.png` — KPIs over time
- `output/plots/04_rsrp_quality_bins.png` — coverage quality bins
- `output/plots/05_band_counts.png` — LTE band from DL EARFCN
- `output/plots/06_top_cells.png` — most-seen Cell Ids
- `output/plots/07_p1_vs_p2_counts.png` — sample counts by source file

If `output/Bogura_merged.csv.gz` already exists, skip Excel reload with:

```bash
python merge_and_analyze.py --from-merged
```

## Merged-log snapshot

| Item | Value |
| --- | --- |
| P1 samples | 1,015,226 (11–20 May 2026) |
| P2 samples | 60,159 (21 May 2026) |
| Merged | 1,075,385 |
| Mean RSRP / RSRQ / SINR | −99.1 dBm / −10.9 dB / 8.4 dB |
| Dominant bands | B3 1800 (50.5%), B1 2100 (42.3%), B41 2500 (6.5%), B8 900 (0.7%) |
| Unique Cell Ids | 2,521 |

RSRP is mostly fair-to-poor: 48% of samples are below −100 dBm; 7% are excellent (> −80 dBm).

## Excel report

`output/Bogura_DriveTest_Report.xlsx`

| Sheet | Contents |
| --- | --- |
| Cover | Dataset size, KPI scorecard, band mix, RSRP/RSRQ/SINR histograms |
| RSRP / RSRQ / SINR | Histogram, CDF, quality-bin charts, coverage map |
| Time Series | 10-minute mean line charts for all three KPIs |
| Sample Log | Evenly spaced subset of the merged samples |

Every sheet has freeze panes disabled and worksheet gridlines turned off (screen and print). Charts are native Excel objects, not just pictures.
