# L3 REPORT FOR BOGRA — v8

Robi vs Grameenphone parameter comparison, plus on-net VoLTE call setup time for Robi, Grameenphone, and Banglalink.
Call setup time is the caller's SIP INVITE to 180 Ringing. Power stays first on the Comparison sheet.

## Versions

| Version | File | Contents |
| --- | --- | --- |
| v1 | `output/L3_REPORT_FOR_BOGRA_v1.xlsx` | Full list. Sections: Mobility, Power, Idle mode, Others. |
| v2 | `output/L3_REPORT_FOR_BOGRA_v2.xlsx` | Stable settings only, with the idle-search finding. |
| v3 | `output/L3_REPORT_FOR_BOGRA_v3.xlsx` | Full list, named sections, one sheet per section, difference sheets. |
| v4 | `output/L3_REPORT_FOR_BOGRA_v4.xlsx` | Frame structure section for the TDD subframe parameters. |
| v5 | `output/L3_REPORT_FOR_BOGRA_v5.xlsx` | Mapped meaning inside the Robi and GP value cells. |
| v6 | `output/L3_REPORT_FOR_BOGRA_v6.xlsx` | Power first on Comparison and Differences. |
| v7 | `output/L3_REPORT_FOR_BOGRA_v7.xlsx` | VoLTE call setup time and the Robi actions. |
| v8 | `output/L3_REPORT_FOR_BOGRA_v8.xlsx` | Setup time split into segments against Banglalink. Current. |

Parameters in v8: 247. Rows that differ: 63.

## Sections

| Section | Parameters | Different | Robi only | GP only | Same |
| --- | ---: | ---: | ---: | ---: | ---: |
| Power | 28 | 0 | 0 | 0 | 28 |
| Mobility | 76 | 14 | 5 | 2 | 55 |
| Idle mode | 70 | 17 | 9 | 1 | 43 |
| Frame structure | 9 | 2 | 0 | 0 | 7 |
| Connected DRX | 6 | 2 | 0 | 0 | 4 |
| MAC | 5 | 1 | 0 | 0 | 4 |
| Scheduling request | 2 | 1 | 0 | 0 | 1 |
| Logical channel | 4 | 0 | 0 | 0 | 4 |
| RLC and PDCP | 11 | 0 | 0 | 0 | 11 |
| CQI and MIMO | 9 | 2 | 0 | 0 | 7 |
| Carrier Aggregation | 4 | 2 | 0 | 0 | 2 |
| RACH and PRACH | 3 | 1 | 0 | 0 | 2 |
| SRS | 10 | 3 | 0 | 0 | 7 |
| Security | 5 | 1 | 0 | 0 | 4 |
| RRC release | 3 | 0 | 0 | 0 | 3 |
| Others | 2 | 0 | 0 | 0 | 2 |

Open Comparison for the full list. Open VoLTE CST for setup time and VoLTE actions for what to change on Robi. Open Differences for every row that is not the same dominant value.
