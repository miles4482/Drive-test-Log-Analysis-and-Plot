# Email — L3 Report for Bogura (complete)

**Subject:** Bogura Drive Test — L3 differences (Robi vs GP), including Event A4

---

Dear Sir / Madam,

Please find the complete summary of the **Layer-3 comparison between Robi and Grameenphone** from the Bogura Active Mode drive test.

**Purpose of this work:** find where Robi and GP settings differ. The target is the difference itself, not to make Robi the same as GP.

**Source:** decoded LTE-RRC / NAS / SIP from `Layer 3_Active Mode_Data_Robi_v1.7z`  
**Report:** L3 REPORT FOR BOGRA v8  
**Scope:** 247 network-configured L3 parameters. **63** rows differ on the dominant value (48 both present but different, 14 Robi only, 3 GP only).

---

## 1. Overall summary

1. **Idle mode is the main policy gap.** Robi keeps measuring other LTE layers until serving RSRP is about **−74 dBm** (`s-NonIntraSearch` **54 dB**). GP stops that search near **−118 dBm** (`s-NonIntraSearch` **10 dB**). Robi `t300` is **1000 ms**; GP dominant `t300` is **200 ms**.
2. **Connected A3 handover recipe is the same** on the dominant setting (offset 1 dB, hysteresis 1 dB, time-to-trigger 640 ms). Other mobility items differ, especially **Event A4**.
3. **Event A4 is a real difference, and it is not one value.** In the drive-test log the dominant Robi A4 is **−124 dBm** (190 of 217, 88%). The database value **−114 dBm** is also in the log, but only **10 times (5%)**. GP dominant A4 is **−116 dBm** (67 of 91, 74%).
4. **Frame structure:** both use TDD configuration **sa2**. Special subframe is **ssp7 (Robi)** vs **ssp5 (GP)**.
5. **Power:** no dominant-value difference was found in this dump (28 parameters).
6. **VoLTE call setup** (on-net, INVITE → 180 Ringing): median **Robi 2.87 s**, **GP 2.46 s**, **Banglalink 0.84 s**. Robi’s extra delay is mainly the **caller QCI-1 bearer (~2.4 s vs GP ~0.22 s)**. That bearer path is separate from the A4 measurement setting.

---

## 2. Section-wise summary table

Order: Power → Idle → Mobility → Frame → others.

| Section | Params | Different | Robi only | GP only | What differs |
| --- | ---: | ---: | ---: | ---: | --- |
| Power | 28 | 0 | 0 | 0 | No difference found in this dump |
| **Idle mode** | 70 | **17** | **9** | **1** | **Largest gap** — search, reselection, t300 |
| **Mobility** | 76 | **14** | **5** | **2** | A3 same; **A4 threshold and profile differ** |
| **Frame structure** | 9 | **2** | 0 | 0 | ssp7 (Robi) vs ssp5 (GP) |
| Connected DRX | 6 | 2 | 0 | 0 | Longer DRX timers on Robi |
| MAC | 5 | 1 | 0 | 0 | maxHARQ-Tx |
| Scheduling request | 2 | 1 | 0 | 0 | SR config index |
| Logical channel | 4 | 0 | 0 | 0 | No difference found |
| RLC and PDCP | 11 | 0 | 0 | 0 | No difference found |
| CQI and MIMO | 9 | 2 | 0 | 0 | Antenna ports: Robi 2Tx vs GP 4Tx |
| Carrier Aggregation | 4 | 2 | 0 | 0 | SCell bandwidth and EARFCN |
| RACH and PRACH | 3 | 1 | 0 | 0 | Root sequence (cell plan) |
| SRS | 10 | 3 | 0 | 0 | GP enables dedicated SRS more often |
| Security | 5 | 1 | 0 | 0 | Algorithms same (eea2 / eia2); chaining count differs |
| RRC release | 3 | 0 | 0 | 0 | No difference found |
| Others | 2 | 0 | 0 | 0 | No difference found |
| **Total** | **247** | **48** | **14** | **3** | **63** non-matching rows |

---

## 3. Major differences

### A. Idle mode

| Parameter | Robi (dominant) | GP (dominant) | Effect |
| --- | --- | --- | --- |
| s-NonIntraSearch | **54 dB** | **10 dB** | Robi searches other layers until ~−74 dBm. GP stops near ~−118 dBm. |
| threshServingLow | **12 dB** | **6 dB** | Serving cell is marked low later on Robi. |
| t300 | **1000 ms** | **200 ms** | Robi waits 1 s for RRC setup; GP typically 200 ms. |
| cellReselectionPriority (serving) | **5** | **3** | Different priority of the camped layer. |
| threshX-High / threshX-Low (SIB5) | **14 dB / 14 dB** | **12 dB / 2 dB** | Different bars to move to or from another LTE frequency. |

### B. Mobility — Event A3 vs Event A4

**A3 (main connected handover):** dominant settings match — offset **1 dB**, hysteresis **1 dB**, time-to-trigger **640 ms**, trigger quantity RSRP.

**A4 (absolute neighbour threshold):** this is where the operators differ.

| | Robi | Grameenphone |
| --- | ---: | ---: |
| A4 reportConfigs in the log | **217** | **91** |
| Dominant threshold | **−124 dBm** (190, 88%) | **−116 dBm** (67, 74%) |
| Database value −114 dBm (IE 26) | **10 (5%)** | **8 (9%)** |
| Other values seen | −134, −127, −110, −108 | −136, −118 |

Robi dominant profile (the 88%):

- threshold **−124 dBm**
- hysteresis **1 dB**
- time-to-trigger **640 ms**
- maxReportCells **8**
- reportInterval **ms5120**
- reportAmount **r1**
- triggerQuantity **RSRP**

The −114 dBm cases (the 5%) are a **different profile**: maxReportCells **4**, reportInterval **ms240**, reportAmount **infinity**.

GP dominant profile:

- threshold **−116 dBm**
- hysteresis **1 dB**
- time-to-trigger **640 ms**
- maxReportCells **4**
- reportInterval **ms240**
- reportAmount **infinity**

**Where A4 sits in the message.**  
A4 is only inside downlink **LTE-RRC `DCCH-RRCConnectionReconfiguration`**, under `measConfig` → `reportConfigToAddModList` → `reportConfigEUTRA` → `eventId = eventA4`. It is linked by `reportConfigId` → `measId` → `measObjectId` → EARFCN.  
It is **not** inside the UE `DCCH-MeasurementReport` (that message only carries `measId` and RSRP/RSRQ). It is **not** tied to a QCI. QCI 1 appears only in NAS ESM dedicated-bearer messages.

**What the A4 is used for (from the EARFCN link, not from an explicit “load HO” flag).**  
L3 does not label a config as load handover. The split below is from target frequency vs the serving A1/A2/A3 layer, plus the reporting style.

| Use | Robi | GP |
| --- | --- | --- |
| **Inter-freq coverage HO** (sparse: ms5120 / r1) | **Main case.** −124 dBm toward another layer. Typical: serving B3 EARFCN **1399** → target B1 **426**, and the reverse; also B41 **40940 / 40742 / 41138**. | Smaller set. Example: −114 dBm, B3 **1600** while serving B1 **250**. |
| **Inter-freq continuous** (ms240 / infinity) — HO or load balance | The **−114 dBm** cases, often toward several layers at once (B1 / B3 / B41). | Dominant **−116 dBm** toward B41 **40142 / 40340 / 40538**. |
| **Early / load-like push** (very low threshold + continuous report) | Small set only: **−134 dBm** toward B1 **426**. | Small set: **−136 dBm** on B1 **250**. |
| **Same-layer (intra) A4** | Rare (about 3). Not inter-freq HO. | Not the main pattern. |

So the drive-test dominant Robi A4 (**−124 dBm**, single sparse report) behaves as **inter-frequency coverage handover preparation**. The database **−114 dBm** profile is present, but it is the minority continuous multi-layer profile, closer to a load-balance style measurement. GP’s main A4 (**−116 dBm**, continuous) is aimed at the B41 layer.

### C. Frame structure

| Parameter | Robi | GP |
| --- | --- | --- |
| subframeAssignment | sa2 | sa2 |
| specialSubframePatterns (SIB and SCell) | **ssp7** | **ssp5** |
| UL cyclic prefix | Normal | Normal |

### D. Other differences worth noting

| Area | Robi | GP |
| --- | --- | --- |
| Connected DRX | Longer onDuration / inactivity (e.g. psf200) | Shorter (e.g. psf2 / psf80) |
| Antenna ports | Dominant an2 (2Tx) | Dominant an4 (4Tx) |
| CA SCell bandwidth | Dominant n75 | Dominant n100 |
| Dedicated SRS | More often release | More often setup |
| VoLTE caller QCI-1 | ~2.4 s | ~0.22 s |
| VoLTE CST (INVITE → 180) | 2.87 s | 2.46 s (Banglalink 0.84 s) |

Security algorithms match (eea2 / eia2). Paging cycle matches (rf128). Cell identity fields (PCI, root sequence, many EARFCNs) differ because the two networks are different plans.

---

## 4. Section-wise summary

**Power.** No dominant difference in this Bogura dump. Uplink open-loop power is the same dominant setting (`p0-NominalPUSCH` −67 dBm, alpha 0.7, RS power 18 dBm).

**Idle mode.** This is the primary L3 difference. Robi’s high `s-NonIntraSearch` keeps the UE measuring other layers until coverage is strong. GP exits that search much earlier. Together with `threshServingLow`, the SIB5 move thresholds, and a 1 s vs 200 ms `t300`, idle camping and RRC setup will not behave the same.

**Mobility.** A3 handover settings match. A4 does not. Robi’s live dominant A4 is −124 dBm on an inter-frequency object, reported once every 5.12 s. GP’s live dominant A4 is −116 dBm with continuous reporting toward B41. The −114 dBm value in the Robi database is in the log, but it is only about 5% of A4 configs and uses a different report profile.

**Frame structure.** Both are TDD sa2. The frame difference to plan around is special subframe **ssp7 vs ssp5**.

**Remaining sections.** DRX, MIMO antenna ports, and CA SCell size differ by deployment. RLC, PDCP, logical channel, and RRC release show no dominant difference. RACH root sequence is cell-specific.

---

## 5. Last note

> The comparison is the **dominant configured value in the Bogura drive-test log**, with sample share. It is not a claim that every cell in the OSS database uses that value.  
> **Idle search / reselection and `t300` are the main L3 policy differences between Robi and GP.**  
> **Event A4 is the main connected-measurement difference:** live Robi is mostly **−124 dBm inter-freq coverage HO**, while the database **−114 dBm** is a minority continuous profile. GP is mostly **−116 dBm** toward B41. A4 is not a QCI parameter and is not written inside the MeasurementReport.  
> **VoLTE setup** is a separate item: Robi’s caller voice bearer (QCI 1) takes about 2.4 s, against about 0.2 s on GP.

Full workbook: `output/L3_REPORT_FOR_BOGRA_v8.xlsx` (Comparison, one sheet per section, Differences).

I can walk through Idle or A4 in the next review, or turn this note into one slide.

Best regards,  
**Mohammad Selim**  
Data Scientist

---
