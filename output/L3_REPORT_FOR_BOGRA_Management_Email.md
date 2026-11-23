# Email draft — L3 Report for Bogura (Robi vs Grameenphone)

**Subject:** Bogura Drive Test — L3 Parameter Comparison (Robi vs GP) | Summary for Management

---

Dear Sir / Madam,

Please find below a management summary of the **Layer-3 (L3) parameter comparison** between **Robi** and **Grameenphone (GP)** from the Bogura Active Mode drive-test logs. A short **VoLTE call setup time** note is included at the end of the major-difference section.

**Source:** Decoded LTE-RRC / SIP Active Mode dumps (`Layer 3_Active Mode_Data_Robi_v1.7z`)  
**Report:** L3 REPORT FOR BOGRA — **v7**  
**Scope:** 247 network-configured L3 parameters compared on dominant (most common) setting

---

## 1. Overall Summary Notes

1. **Power control is aligned.** All **28** Power parameters show the **same dominant setting** on Robi and GP (including `p0-NominalPUSCH` −67 dBm, alpha 0.7, `referenceSignalPower` 18 dBm).
2. **Connected-mode A3 handover behaviour is largely the same.** Dominant A3 offset (1 dB), hysteresis (1 dB), and time-to-trigger (640 ms) match on both networks.
3. **The clearest operator difference is in Idle mode inter-frequency search.** Robi keeps searching other LTE layers until serving RSRP is relatively strong (~**−74 dBm**). GP stops that search much earlier (~**−118 dBm**). This is driven mainly by `s-NonIntraSearch` (Robi **54 dB** vs GP **10 dB**).
4. **RRC connection setup timer differs.** Robi `t300` = **1000 ms**; GP dominant `t300` = **200 ms**. Robi waits longer before declaring RRC setup failure.
5. **TDD frame structure is mostly common** (UL/DL configuration **sa2**, normal UL cyclic prefix). The notable difference is **special subframe pattern**: Robi **ssp7** vs GP **ssp5** (SIB and SCell).
6. **Security algorithms match** (ciphering **eea2**, integrity **eia2**). Paging cycle is the same (**rf128**).
7. Out of **247** parameters, **63** rows differ on dominant value (including Robi-only / GP-only). Most differences are cell-/carrier-specific (EARFCN, PCI, root sequence) or secondary timers; a smaller set is strategically important (idle search, t300, special subframe, DRX, antenna ports).
8. **VoLTE (on-net):** Median call setup time (INVITE → 180 Ringing) is **Robi 2.87 s**, **GP 2.46 s**, **Banglalink 0.84 s**. Robi’s main delay is the **caller QCI-1 dedicated bearer** (~2.4 s vs GP ~0.22 s).

---

## 2. Section-wise Summary Table

*(Order: Power → Idle → Mobility → Frame → others)*

| Section | Parameters | Same | Different* | Robi only | GP only | Management read-out |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| **Power** | 28 | 28 | 0 | 0 | 0 | Fully aligned on dominant settings |
| **Idle mode** | 70 | 43 | 17 | 9 | 1 | Largest strategic gap (search / reselection / t300) |
| **Mobility** | 76 | 55 | 14 | 5 | 2 | A3 core same; some event thresholds & HO control differ |
| **Frame structure** | 9 | 7 | 2 | 0 | 0 | sa2 same; special subframe ssp7 (Robi) vs ssp5 (GP) |
| Connected DRX | 6 | 4 | 2 | 0 | 0 | DRX timers longer on Robi |
| MAC | 5 | 4 | 1 | 0 | 0 | Minor (max HARQ Tx) |
| Scheduling request | 2 | 1 | 1 | 0 | 0 | Config index only |
| Logical channel | 4 | 4 | 0 | 0 | 0 | Aligned |
| RLC and PDCP | 11 | 11 | 0 | 0 | 0 | Aligned |
| CQI and MIMO | 9 | 7 | 2 | 0 | 0 | Antenna ports: Robi 2Tx vs GP 4Tx (dominant) |
| Carrier Aggregation | 4 | 2 | 2 | 0 | 0 | SCell bandwidth / frequency differ |
| RACH and PRACH | 3 | 2 | 1 | 0 | 0 | Root sequence (cell-specific) |
| SRS | 10 | 7 | 3 | 0 | 0 | GP more often enables dedicated SRS |
| Security | 5 | 4 | 1 | 0 | 0 | Algorithms same; chaining count differs |
| RRC release | 3 | 3 | 0 | 0 | 0 | Aligned |
| Others | 2 | 2 | 0 | 0 | 0 | Aligned |
| **Total** | **247** | **184** | **48** | **14** | **3** | **63** non-matching rows overall |

\*“Different” = both operators present, but dominant value is not the same.

---

## 3. Major Differences

### A. Idle mode (highest impact)

| Parameter | Robi (dominant) | GP (dominant) | Why it matters |
| --- | --- | --- | --- |
| **s-NonIntraSearch** | 27 → **54 dB** | 5 → **10 dB** | Robi continues inter-freq / inter-RAT measurement search until serving level is much stronger (~−74 dBm). GP stops near ~−118 dBm. Affects idle reselection aggressiveness and measurement load. |
| **threshServingLow** | 6 → **12 dB** | 3 → **6 dB** | Robi marks the serving cell “low” later (relative to q-RxLevMin). |
| **t300** | **ms1000** | **ms200** | Robi allows 1 s for RRC Connection Setup; GP typically 200 ms. Impacts setup behaviour under poor radio / congestion. |
| **cellReselectionPriority** (serving) | **5** | **3** | Different priority ranking of the serving layer. |
| **threshX-High / threshX-Low** (SIB5) | 7 → **14 dB** / 7 → **14 dB** | 6 → **12 dB** / 1 → **2 dB** | Different bars to move to / from neighbour LTE frequencies. |

### B. Mobility (connected)

| Area | Observation |
| --- | --- |
| **A3 (main HO event)** | Dominant settings **aligned** (offset 1 dB, hysteresis 1 dB, TTT 640 ms). |
| **A2 / A4 thresholds** | Dominant RSRP thresholds differ slightly (network fine-tuning). |
| **IRAT B1 / B2** | Profile not identical (Robi shows B2-GERAN; GP shows B1-GERAN in this dump). |
| **Meas gap / SCell meas cycle** | Some gap offset and `measCycleSCell` differences. |
| **HO control fields** | Many differences are **cell identity / carrier** (PCI, EARFCN, root sequence) — expected between operators, not a “policy” mismatch. |

### C. Frame structure

| Parameter | Robi | GP | Note |
| --- | --- | --- | --- |
| **subframeAssignment** | **sa2** | **sa2** | Same TDD UL/DL configuration |
| **specialSubframePatterns** (SIB & SCell) | **ssp7** | **ssp5** | Different special-subframe DwPTS/GP/UpPTS split |
| **ul-CyclicPrefixLength** | Normal | Normal | Same |

### D. Other notable L3 gaps

| Section | Difference |
| --- | --- |
| **Connected DRX** | Robi longer `onDurationTimer` / `drx-InactivityTimer` (e.g. psf200) vs GP shorter (e.g. psf2 / psf80). |
| **CQI / MIMO** | Dominant antenna ports: Robi **an2** vs GP **an4**; codebook restriction follows (2Tx vs 4Tx TM4). |
| **Carrier Aggregation** | SCell DL bandwidth dominant: Robi **n75** vs GP **n100**; different SCell EARFCNs. |
| **SRS** | GP more often **setup** dedicated SRS; Robi more often **release**. |
| **Power** | **No major difference** on dominant settings. |

### E. VoLTE call setup time (on-net, INVITE → 180 Ringing)

| Operator | Median CST | Caller QCI-1 bearer | Called paging (idle UE) |
| --- | ---: | ---: | ---: |
| **Robi** | **2,871 ms** (n=23) | **2,414 ms** | 1,756 ms |
| **Grameenphone** | **2,458 ms** (n=22) | **224 ms** | 1,860 ms |
| **Banglalink** | **839 ms** (n=22) | **94 ms** | (called UE often still connected) |

**Robi action focus:** (1) bring originating QCI-1 bearer under ~200 ms (Rx/Gx / PCRF / PGW path); (2) shorten INVITE → paging for idle called UE. Ringing/answer delay is similar on all three networks and is not the CST gap.

---

## 4. Section-wise Summary

**Power**  
No dominant-value gap. Uplink open-loop power (`p0-NominalPUSCH` −67 dBm, alpha 0.7), PUCCH nominal power, and RS power (18 dBm) are the same on both networks. Power is not a differentiator in this Bogura dump.

**Idle mode**  
This is the main L3 policy difference. Robi’s high `s-NonIntraSearch` keeps the UE measuring other layers until coverage is strong; GP exits that search early. Combined with different `threshServingLow`, inter-freq thresholds, and a much longer Robi `t300`, idle/camping and RRC setup behaviour will not feel the same even when coverage maps look similar.

**Mobility**  
Core A3 handover recipe is shared. Differences are mostly thresholds for secondary events, IRAT event presence, measurement-gap details, and handover-target identifiers (PCI/EARFCN). Connected mobility strategy looks broadly similar; idle strategy does not.

**Frame structure**  
Both use TDD configuration **sa2**. The operational difference to note for RF/capacity planning is **special subframe pattern ssp7 (Robi) vs ssp5 (GP)** on SIB and SCell.

**Connected DRX / MAC / Scheduling**  
Smaller gaps: DRX timers (Robi longer), `maxHARQ-Tx` (Robi n8 vs GP n5), and SR config index. Useful for optimisation deep-dives; secondary to idle/search for an executive view.

**Logical channel / RLC & PDCP / RRC release**  
Aligned on dominant settings.

**CQI & MIMO / CA / RACH / SRS / Security**  
Antenna-port and CA SCell capability differ by operator deployment. Security algorithms (eea2/eia2) match. RACH root sequences differ as expected per cell plan.

---

## 5. Last Note

This comparison is based on **decoded Active Mode L3 messages** from the Bogura drive test and reports each parameter’s **dominant (most frequent) configured value**, with mapped units and sample share in the workbook. Cell-specific identities (PCI, EARFCN, root sequence) will always differ between operators and should not be read as a “wrong configuration.”

The action-relevant takeaway for management is:

> **Power and connected A3 mobility are largely the same; Idle-mode search / reselection policy and RRC setup timer (`t300`) are not. Treat Idle configuration as the primary L3 differentiator between Robi and GP in Bogura.**  
> **Separately, Robi VoLTE setup is slower than GP/BL mainly because the caller voice bearer (QCI 1) takes ~2.4 s — that is a core/IMS bearer-path item, not an RF L3 offset issue.**

Detailed workbook (section sheets, difference sheets, VoLTE CST / actions):  
`output/L3_REPORT_FOR_BOGRA_v7.xlsx`  
PR: https://github.com/miles4482/Drive-test-Log-Analysis-and-Plot/pull/6

I am available to walk through any section or to prepare a one-page slide version for the next review.

Best regards,  
**Mohammad Selim**  
Data Scientist

---
