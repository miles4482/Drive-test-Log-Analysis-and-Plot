# Bogura Drive-Test Analysis Summary

## What the source files are

- `Borga_Idle_V2.0.7z` is the IDLE Mode scanner export (folder `Borga_Idle_V2.0/` with `1.csv` … `15.csv`).
- Columns include Time, Longitude, Latitude, Cell Id, DL EARFCN, and scanner KPIs `NB RSRP` / `NB RSRQ` / `NB RS SINR` (renamed to RSRP / RSRQ / SINR).
- Concatenating the 15 CSVs produces the **final log**: `output/Bogura_merged.csv.gz`.

## Row counts

- CSV files: 15
- Merged scanner detections: 1,178,276 samples

## Time and area

- 1.csv: 2026-05-11 13:01:29.878000 → 2026-05-11 18:50:10.709000; lon 89.10423–89.75081, lat 24.39244–24.87869
- 2.csv: 2026-05-12 08:09:53.466000 → 2026-05-12 15:57:13.407000; lon 89.35092–89.58578, lat 24.78635–24.90211
- 3.csv: 2026-05-12 16:00:18.164000 → 2026-05-13 18:51:03.176000; lon 89.18838–89.69032, lat 24.41795–25.02303
- 4.csv: 2026-05-13 08:10:02.652000 → 2026-05-13 14:04:32.790000; lon 89.19661–89.54474, lat 24.59393–24.90163
- 5.csv: 2026-05-14 08:11:19.677000 → 2026-05-14 21:38:41.140000; lon 88.98762–89.63143, lat 23.90495–25.26850
- 6.csv: 2026-05-14 14:11:54.894000 → 2026-05-14 18:58:41.240000; lon 89.01444–89.37310, lat 24.84484–25.17747
- 7.csv: 2026-05-15 08:40:47.132000 → 2026-05-16 16:15:57.832000; lon 89.24648–89.69026, lat 23.83306–24.80621
- 8.csv: 2026-05-16 08:09:28.107000 → 2026-05-16 13:54:13.864000; lon 89.34830–89.40431, lat 24.80596–24.87997
- 9.csv: 2026-05-16 16:16:50.547000 → 2026-05-18 16:00:40.773000; lon 89.00137–89.70179, lat 24.36253–25.55249
- 10.csv: 2026-05-17 08:36:28.110000 → 2026-05-20 17:09:35.830000; lon 89.36907–89.61400, lat 24.30490–25.56161
- 11.csv: 2026-05-18 09:10:59.206000 → 2026-05-18 14:26:20.538000; lon 88.87570–89.38816, lat 25.12611–25.54954
- 12.csv: 2026-05-19 08:54:01.122000 → 2026-05-19 13:16:31.701000; lon 89.05164–89.43594, lat 24.01020–24.62607
- 13.csv: 2026-05-19 13:17:57.957000 → 2026-05-20 15:22:25.740000; lon 89.20648–89.69800, lat 23.99574–24.33376
- 14.csv: 2026-05-20 09:57:34.560000 → 2026-05-20 13:28:35.500000; lon 89.68255–89.72066, lat 24.22735–24.47755
- 15.csv: 2026-05-21 09:47:15.621000 → 2026-05-21 14:07:39.580000; lon 89.32870–89.64711, lat 23.91517–24.42915

## Radio KPIs (merged)

- RSRP: n=1,177,963  min=-155.2  p5=-132.7  mean=-101.3  median=-100.3  p95=-71.9  max=-44.5
- RSRQ: n=1,177,963  min=-59.7  p5=-33.6  mean=-19.6  median=-18.4  p95=-10.9  max=0.0
- SINR: n=1,163,267  min=-41.3  p5=-19.8  mean=-3.5  median=-5.0  p95=17.2  max=45.0 (blank SINR samples are skipped)

## RSRP ranges (dBm)

- -90 <= X < Max: 328,354 (27.9%)
- -100 <= X < -90: 252,506 (21.4%)
- -110 <= X < -100: 232,485 (19.7%)
- -115 <= X < -110: 88,230 (7.5%)
- -120 <= X < -115: 68,671 (5.8%)
- -Min <= X < -120: 207,717 (17.6%)

## Bands (from DL EARFCN)

- B8 900 MHz (8): 334,480 (28.4%)
- B1 2100 MHz (1): 298,935 (25.4%)
- B3 1800 MHz (3): 277,097 (23.5%)
- B41 2500 MHz (41): 267,764 (22.7%)

- Unique Cell Ids: 5,132

## Notes

- Campaign is **Scanner** (NB RSRP / NB RSRQ / NB RS SINR). Coverage maps use the strongest RSRP at each timestamp so weak neighbors do not paint over the route.
- Files 1.csv–15.csv cover 11–21 May 2026. Operator is Robi. Layers: L900 (B8), L1800 (B3), L2100 (B1), L2600 (B41).

Plots are written to `output/plots/`.
