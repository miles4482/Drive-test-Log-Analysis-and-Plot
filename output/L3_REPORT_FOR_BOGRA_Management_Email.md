# Email draft — L3 Report for Bogura (short)

**Subject:** Bogura DT — L3 Differences: Robi vs GP

---

Dear Sir / Madam,

Please find a short summary of **L3 parameter differences between Robi and GP** from the Bogura Active Mode drive test (report v8).

**Purpose of this work:** identify where Robi and GP settings differ — not to match GP.

**Scope:** 247 L3 parameters compared. **63** rows show a difference (including Robi-only / GP-only).

---

### 1. Overall notes

- Main gap is **Idle mode** (search / reselection / `t300`).
- **Mobility:** A3 core is the same; several event thresholds and HO fields differ.
- **Frame:** TDD sa2 is the same; special subframe is **ssp7 (Robi)** vs **ssp5 (GP)**.
- **Power:** no difference found in this dump (28/28 same dominant value).
- **VoLTE CST** (INVITE → 180): Robi **2.87 s**, GP **2.46 s**, BL **0.84 s**. Robi’s extra delay is mainly the **caller QCI-1 bearer** (~2.4 s).

---

### 2. Section-wise difference table  
*(Power → Idle → Mobility → Frame → others)*

| Section | Params | Diff* | Robi only | GP only | Difference found |
| --- | ---: | ---: | ---: | ---: | --- |
| Power | 28 | 0 | 0 | 0 | None in this dump |
| **Idle mode** | 70 | **17** | **9** | **1** | **Largest gap** |
| **Mobility** | 76 | **14** | **5** | **2** | Thresholds / IRAT / HO fields |
| **Frame structure** | 9 | **2** | 0 | 0 | ssp7 vs ssp5 |
| Connected DRX | 6 | 2 | 0 | 0 | Longer DRX timers on Robi |
| CQI / MIMO | 9 | 2 | 0 | 0 | 2Tx (Robi) vs 4Tx (GP) |
| CA | 4 | 2 | 0 | 0 | SCell BW / EARFCN |
| SRS | 10 | 3 | 0 | 0 | Dedicated SRS usage |
| Others (MAC, SR, Security, …) | 35 | 4 | 0 | 0 | Minor / cell-specific |
| **Total** | **247** | **48** | **14** | **3** | **63** non-matching rows |

\*Both operators present; dominant value not the same.

---

### 3. Major differences (Robi vs GP)

| Area | Robi | GP |
| --- | --- | --- |
| **s-NonIntraSearch** | **54 dB** (~search until −74 dBm) | **10 dB** (~stop near −118 dBm) |
| **t300** | **1000 ms** | **200 ms** |
| **threshServingLow** | 12 dB | 6 dB |
| **Special subframe** | ssp7 | ssp5 |
| **VoLTE caller QCI-1** | ~2.4 s | ~0.22 s |

---

### 4. Section-wise summary

- **Power:** No Robi–GP difference found here.
- **Idle:** Primary difference — inter-freq search, reselection thresholds, and RRC setup timer.
- **Mobility:** Connected A3 similar; other event / IRAT / HO settings differ.
- **Frame:** ssp7 vs ssp5 is the clear frame difference.

---

### 5. Last note

> Focus is **where Robi differs from GP**. Idle L3 policy and VoLTE caller bearer timing are the main findings from Bogura. Full detail: `L3_REPORT_FOR_BOGRA_v8.xlsx`.

Best regards,  
**Mohammad Selim**  
Data Scientist

---
