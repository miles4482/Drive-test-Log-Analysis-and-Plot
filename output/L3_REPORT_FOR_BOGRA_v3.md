# L3 REPORT FOR BOGRA — v3

Robi vs Grameenphone. Every decoded Layer-3 parameter is on the Comparison sheet.
Mobility, Power, and Idle mode stay first. Parameters that were grouped as Others now have their own section when the IE has a known home (Connected DRX, MAC, RLC and PDCP, CQI and MIMO, Carrier Aggregation, Security, and the rest).

Meaning columns convert the mapped 3GPP value (RSRP index to dBm, offset and hysteresis to dB, q-RxLevMin to dBm). A note is filled where the two dominant values change UE behaviour.

## Versions

| Version | File | Contents |
| --- | --- | --- |
| v1 | `output/L3_REPORT_FOR_BOGRA_v1.xlsx` | Full list. Sections: Mobility, Power, Idle mode, Others. |
| v2 | `output/L3_REPORT_FOR_BOGRA_v2.xlsx` | Stable settings only, with the idle-search finding. |
| v3 | `output/L3_REPORT_FOR_BOGRA_v3.xlsx` | Full list, named sections, one sheet per section, difference sheets. Current. |

Parameters in v3: 247. Rows that differ: 63.

## Sections

| Section | Parameters | Different | Robi only | GP only | Same |
| --- | ---: | ---: | ---: | ---: | ---: |
| Mobility | 79 | 14 | 5 | 2 | 58 |
| Power | 28 | 0 | 0 | 0 | 28 |
| Idle mode | 73 | 18 | 9 | 1 | 45 |
| Connected DRX | 6 | 2 | 0 | 0 | 4 |
| MAC | 5 | 1 | 0 | 0 | 4 |
| Scheduling request | 2 | 1 | 0 | 0 | 1 |
| Logical channel | 4 | 0 | 0 | 0 | 4 |
| RLC and PDCP | 11 | 0 | 0 | 0 | 11 |
| CQI and MIMO | 9 | 2 | 0 | 0 | 7 |
| Carrier Aggregation | 6 | 3 | 0 | 0 | 3 |
| RACH and PRACH | 3 | 1 | 0 | 0 | 2 |
| SRS | 10 | 3 | 0 | 0 | 7 |
| Security | 5 | 1 | 0 | 0 | 4 |
| RRC release | 3 | 0 | 0 | 0 | 3 |
| Others | 3 | 0 | 0 | 0 | 3 |

Open Comparison for the full list. Open Differences for every row that is not the same dominant value. Each section also has its own sheet and, when anything differs, a `Diff <section>` sheet.
