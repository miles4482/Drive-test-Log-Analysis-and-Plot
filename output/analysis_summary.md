# Bogura Drive-Test Analysis Summary

## What the source files are

- `Bogura_P1.part1.rar` … `Bogura_P1.part5.rar` are one split RAR volume set. Extracting them yields **Bogura_P1.xlsx** (Part 1 of the drive test).
- `Bogura_P2.rar` is already a complete archive. Extracting it yields **Bogura_P2.xlsx** (Part 2).
- Both workbooks use the same 8 columns: Time, RSRP, RSRQ, SINR, Longitude, Latitude, Cell Id, DL EARFCN.
- Concatenating P1 then P2 produces the **final log**: `output/Bogura_merged.csv.gz`.

## Row counts

- P1: 1,015,226 samples
- P2: 60,159 samples
- Merged: 1,075,385 samples

## Time and area

- P1: 2026-05-11 13:01:28.303000 → 2026-05-20 18:16:38.412000; lon 88.87569–89.75081, lat 23.83306–25.56161
- P2: 2026-05-21 09:49:55.536000 → 2026-05-21 14:07:35.787000; lon 89.32870–89.64711, lat 23.91553–24.42916

## Radio KPIs (merged)

- RSRP: n=1,075,385  min=-140.0  p5=-119.0  mean=-99.1  median=-99.5  p95=-78.4  max=-48.9
- RSRQ: n=1,075,385  min=-30.0  p5=-17.3  mean=-10.9  median=-10.2  p95=-6.3  max=10.0
- SINR: n=816,456  min=-20.0  p5=-4.0  mean=8.4  median=8.0  p95=22.2  max=31.1 (blank SINR samples are skipped)

## RSRP quality bins

- Very Poor (<-110): 208,326 (19.4%)
- Poor (-110 to -100): 310,445 (28.9%)
- Fair (-100 to -90): 288,629 (26.8%)
- Good (-90 to -80): 193,489 (18.0%)
- Excellent (>-80): 74,496 (6.9%)

## Bands (from DL EARFCN)

- B3 1800 MHz (3): 542,758 (50.5%)
- B1 2100 MHz (1): 455,361 (42.3%)
- B41 2500 MHz (41): 69,783 (6.5%)
- B8 900 MHz (8): 7,483 (0.7%)

- Unique Cell Ids: 2,521

## Notes

- P1 covers multiple drive days (11–20 May 2026). P2 is a shorter session on 21 May 2026.
- About 24% of samples have no SINR (typical when the logger recorded neighbor cells).
- Band 41 (TDD 2500) appears as EARFCN ~40742 / 40940 / 41138.

Plots are written to `output/plots/`.
