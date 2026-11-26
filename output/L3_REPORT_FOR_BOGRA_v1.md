# L3 REPORT FOR BOGRA

**Layer-3 parameter comparison — Robi vs Grameenphone (GP)**

Source: `Layer 3_Active Mode_Data_Robi_v1.7z` (Active Mode decoded LTE-RRC).
Aligned with `cursor/active-mode-11nov26-6bd5` Bogura drive-test work.
Section order: **1. Mobility → 2. Power → 3. Idle mode → 4. Others**.

Excel workbook: [`output/L3_REPORT_FOR_BOGRA.xlsx`](output/L3_REPORT_FOR_BOGRA.xlsx)

## Summary counts

| Section | Parameters | Same | Different | Robi only | GP only |
| --- | ---: | ---: | ---: | ---: | ---: |
| Mobility | 79 | 58 | 14 | 5 | 2 |
| Power | 26 | 26 | 0 | 0 | 0 |
| Idle mode | 73 | 45 | 18 | 9 | 1 |
| Others | 69 | 55 | 14 | 0 | 0 |
| TOTAL | 247 | 184 | 46 | 14 | 3 |

## Highlights (easy compare)

| Parameter | Robi (dominant) | GP (dominant) | Status |
| --- | --- | --- | --- |
| Event A3 offset (a3_Offset) | 2 → 1 dB | 2 → 1 dB | Same (dominant) |
| Hysteresis (hysteresis) | 2 → 1 dB | 2 → 1 dB | Same (dominant) |
| Time To Trigger (timeToTrigger) | ms640 | ms640 | Same (dominant) |
| s-Measure (s_Measure) | 0 → disabled (always measure) | 0 → disabled (always measure) | Same (dominant) |
| q-Hyst (idle) (q_Hyst) | dB4 | dB4 | Same (dominant) |
| s-IntraSearch (s_IntraSearch) | 31 → 62 dB | 29 → 58 dB | Different |
| p0-NominalPUSCH (p0_NominalPUSCH) | -67 dBm | -67 dBm | Same (dominant) |
| alpha (alpha) | al07 | al07 | Same (dominant) |
| PDSCH p-a (p_a) | dB_3 | dB_3 | Same (dominant) |
| defaultPagingCycle (defaultPagingCycle) | rf128 | rf128 | Same (dominant) |

## All differences (Diff only)

| Section | Group | Parameter | Robi | GP |
| --- | --- | --- | --- | --- |
| Mobility | Event A2 | `threshold_RSRP` | 22 → -118 dBm  [476, 34%] | 21 → -119 dBm  [631, 37%] |
| Mobility | Event A4 | `threshold_RSRP` | 16 → -124 dBm  [190, 88%] | 24 → -116 dBm  [67, 74%] |
| Mobility | Event B1 (IRAT) | `b1_Threshold` | — | b1_ThresholdGERAN  [1, 100%] |
| Mobility | Event B2 (IRAT) | `b2_Threshold1` | threshold_RSRP  [13, 100%] | — |
| Mobility | Event B2 (IRAT) | `b2_Threshold2` | b2_Threshold2GERAN  [13, 100%] | — |
| Mobility | Event B2 (IRAT) | `threshold_RSRP` | 22 → -118 dBm  [6, 46%] | — |
| Mobility | HO Control (mobilityControlInfo) | `antennaPortsCount` | an2  [44, 48%] | an4  [82, 52%] |
| Mobility | HO Control (mobilityControlInfo) | `bandIndicator` | — | dcs1800  [49, 100%] |
| Mobility | HO Control (mobilityControlInfo) | `carrierFreq` | 426  [37, 16%] | 250  [44, 27%] |
| Mobility | HO Control (mobilityControlInfo) | `dl_CarrierFreq` | 1399  [232, 74%] | 250  [199, 55%] |
| Mobility | HO Control (mobilityControlInfo) | `maxHARQ_Msg3Tx` | 4  [2, 50%] | 5  [28, 55%] |
| Mobility | HO Control (mobilityControlInfo) | `pusch_HoppingOffset` | 0  [314, 100%] | 22  [284, 78%] |
| Mobility | HO Control (mobilityControlInfo) | `rootSequenceIndex` | 498  [29, 9%] | 18  [35, 10%] |
| Mobility | HO Control (mobilityControlInfo) | `targetPhysCellId` | 61  [33, 11%] | 30  [27, 7%] |
| Mobility | HO Control (mobilityControlInfo) | `ul_CarrierFreq` | 19399  [232, 76%] | 18250  [199, 66%] |
| Mobility | HO Control (mobilityControlInfo) | `zeroCorrelationZoneConfig` | 10  [242, 96%] | 11  [312, 100%] |
| Mobility | Meas Gap | `gp0` | 0  [47, 14%] | 23  [34, 11%] |
| Mobility | Meas Gap | `gp1` | 2  [14, 41%] | — |
| Mobility | Meas Object | `carrierFreq` | 426  [591, 39%] | 1600  [585, 29%] |
| Mobility | Meas Object | `measCycleSCell_r10` | sf320  [708, 79%] | sf640  [71, 100%] |
| Mobility | Measurement / Mobility | `si_RequestForHO_r9` | setup  [8, 100%] | — |
| Idle mode | IRAT GERAN Reselection | `cellReselectionPriority` | — | 5  [24, 49%] |
| Idle mode | IRAT GERAN Reselection | `startingARFCN` | 51  [56, 78%] | 89  [362, 26%] |
| Idle mode | Idle / SIB Common | `allowedMeasBandwidth` | mbw75  [35, 35%] | mbw100  [152, 100%] |
| Idle mode | Idle / SIB Common | `freqBandIndicator` | 3  [373, 68%] | 1  [301, 54%] |
| Idle mode | Idle / SIB Common | `specialSubframePatterns` | ssp7  [19, 61%] | ssp5  [68, 73%] |
| Idle mode | Inter-freq Reselection (SIB5) | `allowedMeasBandwidth_r12` | mbw100  [217, 97%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `cellReselectionPriority` | 6  [1012, 57%] | 7  [723, 57%] |
| Idle mode | Inter-freq Reselection (SIB5) | `cellReselectionPriority_r12` | 6  [212, 100%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `dl_CarrierFreq` | 40940  [373, 18%] | 1600  [327, 26%] |
| Idle mode | Inter-freq Reselection (SIB5) | `dl_CarrierFreq_r12` | 40742  [99, 44%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `presenceAntennaPort1_r12` | FALSE  [224, 100%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `q_RxLevMin_r12` | -64  [224, 100%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `reducedMeasPerformance_r12` | true  [224, 100%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `t_ReselectionEUTRA_r12` | 1  [224, 100%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `threshX_High` | 7 → 14 dB  [1817, 86%] | 6 → 12 dB  [1328, 51%] |
| Idle mode | Inter-freq Reselection (SIB5) | `threshX_High_r12` | 7  [212, 95%] | — |
| Idle mode | Inter-freq Reselection (SIB5) | `threshX_Low` | 7 → 14 dB  [1815, 86%] | 1 → 2 dB  [1328, 51%] |
| Idle mode | Inter-freq Reselection (SIB5) | `threshX_Low_r12` | 7  [212, 95%] | — |
| Idle mode | Intra-freq / Serving Reselection (SIB1/SIB3) | `cellReselectionPriority` | 5  [355, 82%] | 3  [1328, 73%] |
| Idle mode | Intra-freq / Serving Reselection (SIB1/SIB3) | `q_QualMin_r9` | -34  [4, 100%] | -18  [52, 100%] |
| Idle mode | Intra-freq / Serving Reselection (SIB1/SIB3) | `q_RxLevMin` | -64 → -128 dBm  [919, 93%] | 0 → 0 dBm  [1328, 56%] |
| Idle mode | Intra-freq / Serving Reselection (SIB1/SIB3) | `s_IntraSearch` | 31 → 62 dB  [373, 98%] | 29 → 58 dB  [485, 100%] |
| Idle mode | Intra-freq / Serving Reselection (SIB1/SIB3) | `s_NonIntraSearch` | 27 → 54 dB  [349, 92%] | 5 → 10 dB  [419, 86%] |
| Idle mode | Intra-freq / Serving Reselection (SIB1/SIB3) | `threshServingLow` | 6 → 12 dB  [373, 98%] | 3 → 6 dB  [419, 86%] |
| Idle mode | PRACH / RACH Common (SIB2) | `maxHARQ_Msg3Tx` | 5  [373, 98%] | 8  [419, 86%] |
| Idle mode | PRACH / RACH Common (SIB2) | `zeroCorrelationZoneConfig` | 10  [365, 96%] | 11  [485, 100%] |
| Idle mode | SIB2 Common Config | `pusch_HoppingOffset` | 0  [379, 100%] | 22  [382, 79%] |
| Idle mode | UE Timers & Constants (SIB2) | `t300` | ms1000  [379, 100%] | ms200  [351, 72%] |
| Others | Antenna / TM | `antennaPortsCount` | an2  [280, 78%] | an4  [284, 56%] |
| Others | Antenna / TM | `codebookSubsetRestriction` | n2TxAntenna_tm4  [141, 72%] | n4TxAntenna_tm4  [152, 59%] |
| Others | Carrier Aggregation (SCell) | `dl_Bandwidth_r10` | n75  [201, 56%] | n100  [505, 100%] |
| Others | Carrier Aggregation (SCell) | `dl_CarrierFreq_r10` | 426  [201, 56%] | 1600  [234, 46%] |
| Others | Carrier Aggregation (SCell) | `specialSubframePatterns` | ssp7  [106, 90%] | ssp5  [146, 84%] |
| Others | Connected DRX | `drx_InactivityTimer` | psf200  [95, 58%] | psf80  [309, 61%] |
| Others | Connected DRX | `onDurationTimer` | psf200  [95, 58%] | psf2  [296, 58%] |
| Others | MAC Main Config | `maxHARQ_Tx` | n8  [153, 62%] | n5  [177, 69%] |
| Others | RACH Common | `rootSequenceIndex` | 498  [32, 8%] | 18  [42, 9%] |
| Others | SRS | `freqDomainPosition` | 0  [5, 29%] | 12  [38, 31%] |
| Others | SRS | `soundingRS_UL_ConfigDedicated` | release  [134, 89%] | setup  [123, 55%] |
| Others | SRS | `transmissionComb` | 0  [11, 65%] | 1  [65, 53%] |
| Others | Scheduling Request | `sr_ConfigIndex` | 5  [48, 12%] | 12  [80, 17%] |
| Others | Security | `nextHopChainingCount` | 0  [89, 28%] | 2  [77, 21%] |

## How values were taken

- Downlink configuration messages only: RRC Connection Setup / Reconfiguration, SIB1 / SystemInformation, MobilityFromEUTRA, SecurityMode, Release.
- For each parameter, the table shows the **dominant** value and up to 5 most common values with counts and share %.
- Meaning column converts 3GPP mapped fields (RSRP index → dBm = index − 140; a3-Offset / hysteresis → dB = value × 0.5; q-RxLevMin → dBm = value × 2).
- Open the **Event Profiles** sheet for full A1–A6 mobility profiles side by side.
- Parameter taxonomy cross-checked against a full LTE L3 IE list (Mobility / Power / Idle / Others).
