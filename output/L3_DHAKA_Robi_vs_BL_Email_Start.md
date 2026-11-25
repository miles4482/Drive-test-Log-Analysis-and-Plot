Subject: Dhaka VoLTE CST — Robi vs Banglalink (focus areas)

Dear Team,

Please find below a short summary of the Dhaka Layer-3 VoLTE drive-test comparison between **Robi** and **Banglalink (BL)** only.

**Bottom line:** Under the same dial campaign and matched RF (18 idle pairs with weaker-phone RSRP within 8 dB), Robi’s call setup is slower than BL in **12 of 18** cases. The typical gap is about **+0.6 s**, and the worst case reaches about **+4.7 s**. Because coverage was aligned, this is **not** an RF story — it points to **IMS / core timing** (late QCI 1 + slow 183 on every pair, plus a paging/INVITE outlier on one bad session).

**Bottom line — Late QCI 1 (every call)**  
Robi starts the voice bearer only **after** B’s 183 (mean ~**2.85 s**); BL starts it right after INVITE (mean ~**0.19 s**). The same pattern shows up as a slow **183 return** (~**+0.65 s** vs BL on every idle pair). This is the **systematic** CST loss — fix IMS/PCRF bearer timing first; it will help all calls, not one session.

**Bottom line — Paging / INVITE (outlier)**  
On the worst call (`17179869212`), CST is **6.37 s** vs BL **2.15 s**. Almost all of that is **waiting for the page** (Robi page at **5.10 s**, BL at **1.52 s**). After paging, RRC→INVITE is the same (~0.13 s). So radio is fine; the page is triggered late in **IMS/core/MME**. Other Robi calls page in ~0.8–1.7 s — treat this as a **session-level core fault**, not a coverage issue.

**Not the main driver — PRACK / UPDATE**  
183 → 180 is not where Robi loses time (often slightly faster than BL). Do not prioritize SDP renegotiation for this CST gap.

**Recommended focus**
- Fix **late QCI 1 + slow 183 return** (all idle pairs).  
- Investigate **late paging** on outlier `17179869212`.  
- De-prioritize PRACK/UPDATE.

Detailed ladders, pair table, and RCA notes are in the Dhaka call-flow chart (Category tab).

Regards,
