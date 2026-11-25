Subject: Dhaka VoLTE CST — Robi vs Banglalink (focus areas)

Dear Team,

Please find below a short summary of the Dhaka Layer-3 VoLTE drive-test comparison between **Robi** and **Banglalink (BL)** only.

**Bottom line:** Under the same dial campaign and matched RF (18 idle pairs with weaker-phone RSRP within 8 dB), Robi’s call setup is slower than BL in **12 of 18** cases. The typical gap is about **+0.6 s**, and the worst case reaches about **+4.7 s**. Because coverage was aligned, this is **not** an RF story — it points to **IMS / core timing** (late QCI 1 + slow 183 on every pair, plus a paging/INVITE outlier on one bad session).

**Where the gap comes from**

1. **Persistent gap on every idle pair — 183 return + late QCI 1 (main focus)**  
   Robi’s 183 Session Progress takes about **+0.65 s** longer to reach UE-A than BL on every pair. Robi also starts the QCI 1 voice bearer late (mean ~**2.85 s**, after B’s 183), while BL starts it early (mean ~**0.19 s**). This looks like a systematic IMS / PCRF timing issue, not RF.

2. **Outlier gap — INVITE / paging path**  
   On the worst session (Robi pair `17179869212`, CST **6.37 s** vs BL **2.15 s**), almost all delay is **before paging**: page arrives at **5.10 s** on Robi vs **1.52 s** on BL. After the page, radio is similar (~0.13–0.15 s). So the outlier is late paging trigger in IMS/core/MME, not slow RRC after the page.

3. **Not the main driver — PRACK / UPDATE**  
   The step from 183 to 180 Ringing (PRACK/UPDATE) is **not** where Robi loses time; Robi is often slightly faster than BL here. Please do not prioritize SDP renegotiation for this CST gap.

**Recommended focus**

- Fix **late QCI 1 + slow 183 return** (present on all idle pairs).  
- Investigate **INVITE/paging delay** on outlier session `17179869212`.  
- De-prioritize PRACK/UPDATE for CST improvement.

Detailed ladders, pair table, and RCA notes are in the Dhaka call-flow chart (Category tab).

Regards,
