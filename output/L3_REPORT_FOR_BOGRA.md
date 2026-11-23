# L3 report for Bogura

Robi and Grameenphone, from the Active Mode Layer-3 decode (`Layer 3_Active Mode_Data_Robi_v1.7z`).

A value is treated as the operator setting only when it is the same on at least 70% of the cells that broadcast it, and both operators have at least 40 samples. Frequencies, PCI, and per-cell indexes are left out of that table. They describe the drive, not a policy.

Messages in the dump: Robi 20,251, Grameenphone 23,465.

## What is different

The idle UE is where the two networks do not behave the same.

Robi keeps searching other LTE layers until the serving cell is strong (s-NonIntraSearch 54 dB, about -74 dBm given q-RxLevMin -128 dBm). Grameenphone stops that search once the serving cell is merely above the minimum (10 dB, about -118 dBm). Robi also calls the serving cell low sooner (threshServingLow 12 dB vs 6 dB) and waits 1 s for RRC setup. Most GP cells wait 200 ms.

Connected handover is much closer. A3 offset, hysteresis, time-to-trigger, s-Measure, and the L3 RSRP/RSRQ filter are the same. Uplink power (p0-NominalPUSCH -67 dBm, alpha 0.7, p0-NominalPUCCH -115 dBm) is the same. Paging (rf128) is the same. Security is eea2 / eia2 on both.

| What changes | Robi | GP | How stable |
| --- | --- | --- | --- |
| s_NonIntraSearch | 27 → 54 dB | 5 → 10 dB | Robi 92% of 379, GP 86% of 485 |
| threshServingLow | 6 → 12 dB | 3 → 6 dB | Robi 98% of 379, GP 86% of 485 |
| s_IntraSearch | 31 → 62 dB | 29 → 58 dB | Robi 98% of 379, GP 100% of 485 |
| cellReselectionPriority | 5 | 3 | Robi 82% of 433, GP 73% of 1813 |
| t300 | ms1000 | ms200 | Robi 100% of 379, GP 72% of 485 |
| threshold_RSRP | 16 → -124 dBm | 24 → -116 dBm | Robi 88% of 217, GP 74% of 91 |
| measCycleSCell_r10 | sf320 | sf640 | Robi 79% of 893, GP 100% of 71 |
| maxHARQ_Msg3Tx | 5 | 8 | Robi 98% of 379, GP 86% of 485 |
| pusch_HoppingOffset | 0 | 22 | Robi 100% of 379, GP 79% of 485 |
| zeroCorrelationZoneConfig | 10 | 11 | Robi 96% of 379, GP 100% of 485 |
| specialSubframePatterns | ssp7 | ssp5 | Robi 90% of 118, GP 84% of 173 |

### Why those rows matter

- **s_NonIntraSearch.** This is the idle threshold for looking at other LTE layers. With q-RxLevMin at -128 dBm, Robi keeps measuring other layers until serving RSRP is about -74 dBm. GP stops that search once serving RSRP is only about -118 dBm.
- **threshServingLow.** Idle rule for leaving a higher-priority layer. Robi treats the serving cell as low from about -116 dBm. GP waits until about -122 dBm.
- **s_IntraSearch.** Both still measure the same frequency across almost the whole drive. The 4 dB gap (about -66 dBm vs -70 dBm) is small next to the inter-layer gap.
- **cellReselectionPriority.** Absolute priority of the serving frequency the UE was camped on. Robi's typical camped layer is priority 5. GP's is priority 3. This follows the band plan, not a single timer.
- **t300.** How long the UE waits for RRC setup before it gives up. Robi allows 1 s. On most GP cells the wait is 200 ms, so a weak setup fails and retries sooner.
- **threshold_RSRP.** Event A4: a neighbour must beat this level to be reported. Robi accepts a neighbour around -124 dBm. GP asks for about -116 dBm, so the report starts later.
- **measCycleSCell_r10.** How often a deactivated secondary cell is measured. Robi uses 320 ms on most carriers. The GP samples that carry this field use 640 ms.
- **maxHARQ_Msg3Tx.** How many times Msg3 is retransmitted during access. GP allows more attempts (8 vs 5).
- **pusch_HoppingOffset.** PUSCH frequency-hopping offset in the common config. Robi uses 0. GP uses 22 on most cells.
- **zeroCorrelationZoneConfig.** PRACH zero-correlation zone. One step apart (10 vs 11). A cell-radius planning choice, not a mobility policy.
- **specialSubframePatterns.** TDD special-subframe pattern on the cells that broadcast it. Robi ssp7, GP ssp5.

## What is the same

83 settings are the same dominant value on both networks, each on at least 90% of samples.

| Setting | Value | Robi samples | GP samples |
| --- | --- | ---: | ---: |
| triggerQuantity | rsrp | 4084 | 5295 |
| presenceAntennaPort1 | FALSE | 1507 | 2048 |
| q_RxLevMin | -64 → -128 dBm | 2054 | 1264 |
| t_ReselectionEUTRA | 1 | 2054 | 1264 |
| cqi_FormatIndicatorPeriodic_r10 | widebandCQI_r10 | 842 | 1175 |
| timeAlignmentTimerDedicated | infinity | 669 | 1009 |
| ue_TransmitAntennaSelection | release | 633 | 1004 |
| cellBarred | notBarred | 550 | 560 |
| intraFreqReselection | allowed | 550 | 560 |
| si_WindowLength | ms40 | 550 | 560 |
| dsr_TransMax | n64 | 401 | 461 |
| messagePowerOffsetGroupB | dB10 | 383 | 536 |
| powerRampingStep | dB2 | 383 | 536 |
| preambleInitialReceivedTargetPower | dBm_104 | 383 | 536 |
| additionalSpectrumEmission | 1 | 379 | 485 |
| alpha | al07 | 379 | 486 |
| defaultPagingCycle | rf128 | 379 | 485 |
| deltaF_PUCCH_Format1 | deltaF0 | 379 | 486 |
| deltaF_PUCCH_Format1b | deltaF3 | 379 | 486 |
| deltaF_PUCCH_Format2 | deltaF1 | 379 | 486 |
| deltaF_PUCCH_Format2a | deltaF2 | 379 | 486 |
| deltaF_PUCCH_Format2b | deltaF2 | 379 | 486 |
| deltaPUCCH_Shift | ds1 | 379 | 485 |
| deltaPreambleMsg3 | 4 → 8 dB | 379 | 486 |
| enable64QAM | TRUE | 379 | 485 |
| groupHoppingEnabled | FALSE | 379 | 485 |
| highSpeedFlag | FALSE | 379 | 485 |
| hoppingMode | interSubFrame | 379 | 485 |
| mac_ContentionResolutionTimer | sf64 | 379 | 485 |
| messageSizeGroupA | b56 | 379 | 485 |
| modificationPeriodCoeff | n2 | 379 | 485 |
| n310 | n10 | 379 | 485 |
| n311 | n1 | 379 | 485 |
| nB | oneT | 379 | 485 |
| n_SB | 4 | 379 | 485 |
| numberOfRA_Preambles | n52 | 379 | 485 |
| p0_NominalPUCCH | -115 dBm | 379 | 486 |
| p0_NominalPUSCH | -67 dBm | 379 | 486 |
| preambleTransMax | n10 | 379 | 485 |
| ra_ResponseWindowSize | sf10 | 379 | 485 |
| sequenceHoppingEnabled | FALSE | 379 | 485 |
| sf_High | oDot75 | 379 | 485 |
| sf_Medium | lDot0 | 379 | 485 |
| sizeOfRA_PreamblesGroupA | n28 | 379 | 485 |
| t301 | ms200 | 379 | 485 |
| t310 | ms1000 | 379 | 485 |
| t311 | ms10000 | 379 | 485 |
| timeAlignmentTimerCommon | infinity | 379 | 485 |
| ul_CyclicPrefixLength | len1 | 379 | 485 |
| phich_Duration | normal | 360 | 505 |
| phich_Resource | one | 360 | 505 |
| filterCoefficientRSRP | fc6 | 350 | 415 |
| filterCoefficientRSRQ | fc6 | 350 | 415 |
| prohibitPHR_Timer | sf100 | 386 | 336 |
| t304 | ms500 | 314 | 362 |
| cipheringAlgorithm | eea2 | 301 | 304 |
| integrityProtAlgorithm | eia2 | 301 | 304 |
| s_Measure | 0 → disabled (always measure) | 301 | 303 |
| reportAddNeighMeas_r10 | setup | 255 | 911 |
| maxRetxThreshold | t32 | 257 | 254 |
| retxBSR_Timer | sf320 | 245 | 258 |
| ttiBundling | FALSE | 245 | 258 |
| altCQI_Table_r12 | allSubframes | 189 | 271 |
| drx_RetransmissionTimer | psf8 | 163 | 508 |
| ackNackRepetition | release | 147 | 182 |
| periodicBSR_Timer | sf10 | 147 | 182 |
| subframeAssignment | sa2 | 118 | 173 |
| statusReportRequired | TRUE | 124 | 107 |
| cqi_FormatIndicatorPeriodic | widebandCQI | 215 | 98 |
| sn_FieldLength | size10 | 190 | 90 |
| filterCoefficient | fc6 | 88 | 168 |
| maxLayersMIMO_r10 | fourLayers | 88 | 372 |
| accumulationEnabled | TRUE | 75 | 166 |
| deltaMCS_Enabled | en0 | 75 | 166 |
| p0_UE_PUCCH | 0 dBm | 75 | 166 |
| p0_UE_PUSCH | 0 dBm | 75 | 166 |
| betaOffset_CQI_Index | 12 | 105 | 68 |
| betaOffset_RI_Index | 5 | 105 | 68 |
| tpc_PDCCH_ConfigPUCCH | release | 71 | 68 |
| tpc_PDCCH_ConfigPUSCH | release | 71 | 68 |
| t_ReselectionGERAN | 1 | 54 | 332 |
| releaseCause | other | 52 | 46 |
| pdcp_SN_Size | len12bits | 46 | 45 |

## Not one network-wide value

These are configured, but the dominant value is under 70% on at least one operator. Do not quote a single number for them.

| Setting | Robi mode | GP mode |
| --- | --- | --- |
| maxReportCells | 1 (56%) | 1 (55%) |
| reportInterval | ms480 (45%) | ms480 (35%) |
| reportAmount | r1 (61%) | r1 (58%) |
| reportQuantity | both (63%) | both (58%) |
| threshX_High | 7 → 14 dB (86%) | 6 → 12 dB (51%) |
| threshX_Low | 7 → 14 dB (86%) | 1 → 2 dB (51%) |
| offsetFreq | dB0 (99%) | dB0 (63%) |
| allowedMeasBandwidth | mbw100 (47%) | mbw100 (84%) |
| threshold_RSRP | 22 → -118 dBm (34%) | 21 → -119 dBm (37%) |
| cellReselectionPriority | 6 (57%) | 7 (57%) |
| si_Periodicity | rf16 (46%) | rf16 (37%) |
| nomPDSCH_RS_EPRE_Offset | 0 → 0 dB (48%) | 0 → 0 dB (35%) |
| simultaneousAckNackAndCQI | FALSE (64%) | FALSE (60%) |
| q_RxLevMin | -64 → -128 dBm (93%) | 0 → 0 dBm (56%) |
| threshold_RSRP | 24 → -116 dBm (32%) | 24 → -116 dBm (24%) |
| a3_Offset | 2 → 1 dB (46%) | 2 → 1 dB (35%) |
| reportOnLeave | FALSE (59%) | FALSE (62%) |
| prioritisedBitRate | kBps8 (64%) | kBps8 (61%) |
| bucketSizeDuration | ms300 (51%) | ms300 (52%) |
| p_a | dB_3 (88%) | dB_3 (63%) |
| measGapConfig | setup (59%) | setup (64%) |
| nomPDSCH_RS_EPRE_Offset_r10 | 0 (47%) | 0 (44%) |
| prach_FreqOffset | 7 (92%) | 7 (51%) |
| dl_Bandwidth_r10 | n75 (56%) | n100 (100%) |
| antennaPortsCount | an2 (78%) | an4 (56%) |

Workbook: `output/L3_REPORT_FOR_BOGRA.xlsx`.
