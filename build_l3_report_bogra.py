#!/usr/bin/env python3
"""Build L3 REPORT FOR BOGRA — Robi vs Grameenphone parameter comparison.

Source: Layer 3_Active Mode_Data_Robi_v1.7z (decoded LTE-RRC Active Mode dumps).
Compares every network-configured L3 leaf parameter seen for Robi and GP.
Section order: Mobility → Power → Idle mode → Others.
"""

from __future__ import annotations

import argparse
import glob
import re
import statistics
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DEFAULT_L3_DIR = ROOT / "extracted" / "l3" / "data"
DEFAULT_OUT = ROOT / "output" / "L3_REPORT_FOR_BOGRA.xlsx"

OPS = ("Robi", "Grameenphone")
OP_LABEL = {"Robi": "Robi", "Grameenphone": "GP (Grameenphone)"}

HDR = re.compile(
    r"^(\d{2}:\d{2}:\d{2}\.\d+)\s+([UD])\s+(\S+)\s+(.+?)\s+"
    r"(LTE|UMTS|GSM|VoIP)\s+(\S+)\s+(\S+)\s+(\S+)\s*$"
)

# Leaf IE names that are network settings (not UE identity / resource indexes).
PARAM_WHITELIST = {
    # Mobility / measurement
    "a3_Offset",
    "a6_Offset_r10",
    "a6_ReportOnLeave_r10",
    "hysteresis",
    "timeToTrigger",
    "triggerQuantity",
    "reportQuantity",
    "maxReportCells",
    "reportInterval",
    "reportAmount",
    "threshold_RSRP",
    "threshold_RSRQ",
    "s_Measure",
    "filterCoefficientRSRP",
    "filterCoefficientRSRQ",
    "allowedMeasBandwidth",
    "allowedMeasBandwidth_r12",
    "presenceAntennaPort1",
    "presenceAntennaPort1_r12",
    "offsetFreq",
    "carrierFreq",
    "dl_CarrierFreq",
    "dl_CarrierFreq_r10",
    "dl_CarrierFreq_r12",
    "ul_CarrierFreq",
    "measCycleSCell_r10",
    "reportOnLeave",
    "reportAddNeighMeas_r10",
    "useAutonomousGaps_r14",
    "purpose",
    "includeNeighMeas",
    "si_RequestForHO_r9",
    "t304",
    "t304_r9",
    "targetPhysCellId",
    "carrierBandwidth_dl",
    "carrierBandwidth_ul",
    "cellIndividualOffset",
    "gapOffset",
    "gp0",
    "gp1",
    "measGapConfig",
    "handoverType",
    "keyChangeIndicator",
    "nextHopChainingCount",
    "bandIndicator",
    "triggerType",
    "b2_Threshold1",
    "b2_Threshold2",
    "b1_Threshold",
    "threshold_UTRA_RSCP",
    "threshold_GERAN",
    # Power
    "p0_NominalPUSCH",
    "p0_NominalPUCCH",
    "alpha",
    "p0_UE_PUSCH",
    "p0_UE_PUCCH",
    "pSRS_Offset",
    "deltaMCS_Enabled",
    "accumulationEnabled",
    "filterCoefficient",
    "deltaPreambleMsg3",
    "p_a",
    "p_b",
    "referenceSignalPower",
    "p_Max",
    "preambleInitialReceivedTargetPower",
    "powerRampingStep",
    "messagePowerOffsetGroupB",
    "deltaF_PUCCH_Format1",
    "deltaF_PUCCH_Format1b",
    "deltaF_PUCCH_Format2",
    "deltaF_PUCCH_Format2a",
    "deltaF_PUCCH_Format2b",
    "periodicPHR_Timer",
    "prohibitPHR_Timer",
    "dl_PathlossChange",
    "nomPDSCH_RS_EPRE_Offset",
    "nomPDSCH_RS_EPRE_Offset_r10",
    # Idle / SIB / reselection / paging
    "q_RxLevMin",
    "q_RxLevMin_r12",
    "q_QualMin_r9",
    "q_Hyst",
    "s_IntraSearch",
    "s_NonIntraSearch",
    "threshServingLow",
    "cellReselectionPriority",
    "cellReselectionPriority_r12",
    "t_ReselectionEUTRA",
    "t_ReselectionEUTRA_r12",
    "t_ReselectionUTRA",
    "t_ReselectionGERAN",
    "sf_Medium",
    "sf_High",
    "defaultPagingCycle",
    "nB",
    "modificationPeriodCoeff",
    "intraFreqReselection",
    "cellBarred",
    "si_WindowLength",
    "si_Periodicity",
    "freqBandIndicator",
    "ul_Bandwidth",
    "dl_Bandwidth_r10",
    "additionalSpectrumEmission",
    "timeAlignmentTimerCommon",
    "t300",
    "t301",
    "t310",
    "t311",
    "n310",
    "n311",
    "threshX_High",
    "threshX_Low",
    "threshX_High_r12",
    "threshX_Low_r12",
    "q_OffsetFreq",
    "t_Evaluation",
    "t_HystNormal",
    "n_CellChangeMedium",
    "n_CellChangeHigh",
    "q_HystSF_Medium",
    "q_HystSF_High",
    "s_IntraSearchP_r9",
    "s_IntraSearchQ_r9",
    "s_NonIntraSearchP_r9",
    "s_NonIntraSearchQ_r9",
    "threshServingLowQ_r9",
    "startingARFCN",
    "t320",
    "reducedMeasPerformance_r12",
    # Others / MAC / RLC / DRX / CA / security-adjacent
    "maxHARQ_Tx",
    "periodicBSR_Timer",
    "retxBSR_Timer",
    "ttiBundling",
    "timeAlignmentTimerDedicated",
    "onDurationTimer",
    "drx_InactivityTimer",
    "drx_RetransmissionTimer",
    "longDRX_CycleStartOffset",
    "longDRX_Cycle",
    "shortDRX_Cycle",
    "drxShortCycleTimer",
    "cqi_Mask_r9",
    "sr_ConfigIndex",
    "dsr_TransMax",
    "transmissionMode",
    "transmissionMode_r10",
    "codebookSubsetRestriction",
    "ue_TransmitAntennaSelection",
    "ackNackRepetition",
    "betaOffset_ACK_Index",
    "betaOffset_RI_Index",
    "betaOffset_CQI_Index",
    "cqi_ReportModeAperiodic",
    "cqi_ReportModeAperiodic_r10",
    "simultaneousAckNackAndCQI",
    "cqi_FormatIndicatorPeriodic",
    "cqi_FormatIndicatorPeriodic_r10",
    "numberOfRA_Preambles",
    "sizeOfRA_PreamblesGroupA",
    "messageSizeGroupA",
    "preambleTransMax",
    "ra_ResponseWindowSize",
    "mac_ContentionResolutionTimer",
    "maxHARQ_Msg3Tx",
    "highSpeedFlag",
    "zeroCorrelationZoneConfig",
    "enable64QAM",
    "enable64QAM_v1270",
    "hoppingMode",
    "n_SB",
    "pusch_HoppingOffset",
    "groupHoppingEnabled",
    "sequenceHoppingEnabled",
    "deltaPUCCH_Shift",
    "ul_CyclicPrefixLength",
    "t_PollRetransmit",
    "pollPDU",
    "pollByte",
    "maxRetxThreshold",
    "t_Reordering",
    "t_StatusProhibit",
    "priority",
    "prioritisedBitRate",
    "bucketSizeDuration",
    "logicalChannelGroup",
    "cipheringAlgorithm",
    "integrityProtAlgorithm",
    "releaseCause",
    "redirectedCarrierInfo",
    "extendedWaitTime_r10",
    "antennaPortsCount",
    "phich_Duration",
    "subframeAssignment",
    "specialSubframePatterns",
    "tdd_SpecialSubframe_r11",
    "maxLayersMIMO_r10",
    "altCQI_Table_r12",
    "discardTimer",
    "pdcp_SN_Size",
    "statusReportRequired",
    "srs_BandwidthConfig",
    "srs_SubframeConfig",
    "ackNackSRS_SimultaneousTransmission",
    "srs_MaxUpPts",
    "tdd_AckNackFeedbackMode",
    "phich_Resource",
    "rootSequenceIndex",
    "prach_ConfigIndex",
    "prach_FreqOffset",
    "sn_FieldLength",
    "headerCompression",
    "tpc_PDCCH_ConfigPUCCH",
    "tpc_PDCCH_ConfigPUSCH",
    "soundingRS_UL_ConfigDedicated",
    "soundingRS_UL_ConfigCommon",
    "srs_Bandwidth",
    "srs_HoppingBandwidth",
    "freqDomainPosition",
    "duration",
    "transmissionComb",
    "semiPersistSchedIntervalDL",
    "semiPersistSchedIntervalUL",
    "numberOfConfSPS_Processes",
    "implicitReleaseAfter",
    "sCellDeactivationTimer_r10",
    "sr_ProhibitTimer_r9",
}

# Skip UE-specific / identity / opaque resource indexes.
PARAM_BLACKLIST_PREFIX = (
    "rrc_TransactionIdentifier",
    "measId",
    "measObjectId",
    "reportConfigId",
    "srb_Identity",
    "drb_Identity",
    "eps_BearerIdentity",
    "logicalChannelIdentity",
    "physCellId",
    "cellIdentity",
    "CI",
    "Macro",
    "Sector",
    "MCC_MNC_Digit",
    "trackingAreaCode",
    "newUE_Identity",
    "cqi_PUCCH_ResourceIndex",
    "cqi_pmi_ConfigIndex",
    "ri_ConfigIndex",
    "sr_PUCCH_ResourceIndex",
    "prach_FreqOffset",
    "rootSequenceIndex",
    "prach_ConfigIndex",
    "n1PUCCH_AN",
    "nRB_CQI",
    "nCS_AN",
    "groupAssignmentPUSCH",
    "cyclicShift",
    "systemInfoValueTag",
    "spare",
)

# v1 used four buckets. From v3, rows that used to sit in Others keep a real section
# name when the IE belongs to one. Mobility, Power, and Idle mode stay first.
SECTION_ORDER = (
    "Mobility",
    "Power",
    "Idle mode",
    "Frame structure",
    "Connected DRX",
    "MAC",
    "Scheduling request",
    "Logical channel",
    "RLC and PDCP",
    "CQI and MIMO",
    "Carrier Aggregation",
    "RACH and PRACH",
    "SRS",
    "SPS",
    "Security",
    "RRC release",
    "Others",
)

# Comparison and Differences list Power before Mobility. Section tabs keep SECTION_ORDER.
COMPARE_SECTION_ORDER = ("Power",) + tuple(s for s in SECTION_ORDER if s != "Power")

_OTHERS_SECTION = {
    "Connected DRX": "Connected DRX",
    "MAC Main Config": "MAC",
    "Scheduling Request": "Scheduling request",
    "Logical Channel": "Logical channel",
    "RLC": "RLC and PDCP",
    "PDCP / RLC UM": "RLC and PDCP",
    "Antenna / TM": "CQI and MIMO",
    "CQI / PUSCH Dedicated": "CQI and MIMO",
    "PHICH / Antenna Common": "CQI and MIMO",
    "Carrier Aggregation (SCell)": "Carrier Aggregation",
    "RACH Common": "RACH and PRACH",
    "SRS": "SRS",
    "SPS": "SPS",
    "Security": "Security",
    "RRC Release": "RRC release",
    "TPC (group)": "Power",
}


# TDD-Config and the fields that describe the radio frame. SRS subframe config stays under SRS.
_FRAME_GROUP = {
    "subframeAssignment": "TDD UL/DL configuration",
    "specialSubframePatterns": "Special subframe",
    "tdd_SpecialSubframe_r11": "Special subframe",
    "ul_CyclicPrefixLength": "Cyclic prefix",
    "tdd_AckNackFeedbackMode": "TDD ACK/NACK",
}


def publish_section(section: str, group: str, param: str = "") -> str:
    """Keep Mobility / Power / Idle. Name the old Others groups."""
    if param in _FRAME_GROUP:
        return "Frame structure"
    if group == "TPC (group)":
        return "Power"
    if section == "Others":
        return _OTHERS_SECTION.get(group, "Others")
    return section

SECTION_OF = {}


def _classify(param: str, path: str) -> tuple[str, str]:
    """Return (section, group) for a parameter given its path context."""
    p = param
    pl = path.lower()

    # Power by parameter name first (SIB2 power must not fall into Idle).
    POWER_PARAMS = {
        "p0_NominalPUSCH",
        "p0_NominalPUCCH",
        "alpha",
        "p0_UE_PUSCH",
        "p0_UE_PUCCH",
        "pSRS_Offset",
        "deltaMCS_Enabled",
        "accumulationEnabled",
        "deltaPreambleMsg3",
        "p_a",
        "p_b",
        "referenceSignalPower",
        "p_Max",
        "preambleInitialReceivedTargetPower",
        "powerRampingStep",
        "messagePowerOffsetGroupB",
        "periodicPHR_Timer",
        "prohibitPHR_Timer",
        "dl_PathlossChange",
        "nomPDSCH_RS_EPRE_Offset",
        "nomPDSCH_RS_EPRE_Offset_r10",
        "deltaF_PUCCH_Format1",
        "deltaF_PUCCH_Format1b",
        "deltaF_PUCCH_Format2",
        "deltaF_PUCCH_Format2a",
        "deltaF_PUCCH_Format2b",
        "filterCoefficient",
    }
    if p in POWER_PARAMS or any(
        k in pl
        for k in (
            "uplinkpowercontrol",
            "pdsch_configdedicated",
            "pdsch_configcommon",
            "phr_config",
            "powerramping",
            "messagepoweroffset",
        )
    ):
        if "phr" in pl or p in {"periodicPHR_Timer", "prohibitPHR_Timer", "dl_PathlossChange"}:
            return "Power", "PHR"
        if "pdsch" in pl or p in {"p_a", "p_b", "referenceSignalPower", "nomPDSCH_RS_EPRE_Offset", "nomPDSCH_RS_EPRE_Offset_r10"}:
            return "Power", "PDSCH / RS Power"
        if "rach" in pl or "powerramping" in pl or p in {
            "preambleInitialReceivedTargetPower",
            "powerRampingStep",
            "messagePowerOffsetGroupB",
            "deltaPreambleMsg3",
        }:
            return "Power", "RACH Power"
        return "Power", "UL Power Control"

    # Idle: reselection / paging / timers / SIB selection — but not SIB2 power (handled above)
    IDLE_PARAMS = {
        "q_RxLevMin",
        "q_RxLevMin_r12",
        "q_QualMin_r9",
        "q_Hyst",
        "s_IntraSearch",
        "s_NonIntraSearch",
        "threshServingLow",
        "cellReselectionPriority",
        "cellReselectionPriority_r12",
        "t_ReselectionEUTRA",
        "t_ReselectionEUTRA_r12",
        "t_ReselectionUTRA",
        "t_ReselectionGERAN",
        "defaultPagingCycle",
        "nB",
        "modificationPeriodCoeff",
        "intraFreqReselection",
        "cellBarred",
        "si_WindowLength",
        "si_Periodicity",
        "t300",
        "t301",
        "t310",
        "t311",
        "n310",
        "n311",
        "threshX_High",
        "threshX_Low",
        "threshX_High_r12",
        "threshX_Low_r12",
        "q_OffsetFreq",
        "sf_Medium",
        "sf_High",
        "t_Evaluation",
        "t_HystNormal",
        "n_CellChangeMedium",
        "n_CellChangeHigh",
        "timeAlignmentTimerCommon",
        "ul_Bandwidth",
        "additionalSpectrumEmission",
        "freqBandIndicator",
        "t320",
        "reducedMeasPerformance_r12",
        "startingARFCN",
    }
    idle_path = any(
        k in pl
        for k in (
            "sib1",
            "sib3",
            "sib4",
            "sib5",
            "sib6",
            "sib7",
            "systeminformationblocktype1",
            "cellreselection",
            "cellselection",
            "pcch_config",
            "bcch_config",
            "ue_timersandconstants",
            "freqinfo",
            "schedulinginfolist",
            "intrafreqcellreselection",
            "interfreqcarrierfreq",
            "carrierfreqsinfo",
        )
    )
    # SIB2 non-power commons also idle
    if p in IDLE_PARAMS or idle_path or (
        "sib2" in pl
        and p
        in {
            "numberOfRA_Preambles",
            "sizeOfRA_PreamblesGroupA",
            "messageSizeGroupA",
            "preambleTransMax",
            "ra_ResponseWindowSize",
            "mac_ContentionResolutionTimer",
            "maxHARQ_Msg3Tx",
            "highSpeedFlag",
            "zeroCorrelationZoneConfig",
            "enable64QAM",
            "enable64QAM_v1270",
            "hoppingMode",
            "n_SB",
            "pusch_HoppingOffset",
            "groupHoppingEnabled",
            "sequenceHoppingEnabled",
            "deltaPUCCH_Shift",
            "ul_CyclicPrefixLength",
            "specialSubframePatterns",
            "subframeAssignment",
            "allowedMeasBandwidth",
        }
    ):
        if p in {"t300", "t301", "t310", "t311", "n310", "n311"}:
            return "Idle mode", "UE Timers & Constants (SIB2)"
        if p in {"defaultPagingCycle", "nB", "modificationPeriodCoeff", "si_WindowLength", "si_Periodicity"}:
            return "Idle mode", "Paging / SI Scheduling"
        if "geran" in pl or p in {"t_ReselectionGERAN", "startingARFCN"}:
            return "Idle mode", "IRAT GERAN Reselection"
        if "utra" in pl or p == "t_ReselectionUTRA":
            return "Idle mode", "IRAT UTRA Reselection"
        if "interfreq" in pl or p in {
            "threshX_High",
            "threshX_Low",
            "threshX_High_r12",
            "threshX_Low_r12",
            "q_OffsetFreq",
            "q_RxLevMin_r12",
            "t_ReselectionEUTRA_r12",
            "cellReselectionPriority_r12",
            "reducedMeasPerformance_r12",
        }:
            return "Idle mode", "Inter-freq Reselection (SIB5)"
        if p == "t320":
            return "Idle mode", "Idle Redirect / Priority (t320)"
        if p in {
            "q_Hyst",
            "s_IntraSearch",
            "s_NonIntraSearch",
            "threshServingLow",
            "cellReselectionPriority",
            "t_ReselectionEUTRA",
            "q_RxLevMin",
            "q_QualMin_r9",
            "sf_Medium",
            "sf_High",
            "intraFreqReselection",
            "cellBarred",
        }:
            return "Idle mode", "Intra-freq / Serving Reselection (SIB1/SIB3)"
        if p in {
            "rootSequenceIndex",
            "prach_ConfigIndex",
            "prach_FreqOffset",
            "highSpeedFlag",
            "zeroCorrelationZoneConfig",
            "numberOfRA_Preambles",
            "sizeOfRA_PreamblesGroupA",
            "messageSizeGroupA",
            "preambleTransMax",
            "ra_ResponseWindowSize",
            "mac_ContentionResolutionTimer",
            "maxHARQ_Msg3Tx",
        } or "prach_config" in pl or "rach_config" in pl:
            return "Idle mode", "PRACH / RACH Common (SIB2)"
        if "sib2" in pl or p in {
            "pusch_HoppingOffset",
            "ul_Bandwidth",
            "timeAlignmentTimerCommon",
            "soundingRS_UL_ConfigCommon",
            "srs_BandwidthConfig",
            "srs_SubframeConfig",
            "ackNackSRS_SimultaneousTransmission",
            "srs_MaxUpPts",
            "phich_Resource",
            "phich_Duration",
        }:
            return "Idle mode", "SIB2 Common Config"
        return "Idle mode", "Idle / SIB Common"

    # Mobility
    if any(
        k in pl
        for k in (
            "measconfig",
            "measobject",
            "reportconfig",
            "quantityconfig",
            "mobilitycontrol",
            "eventa",
            "measgap",
            "mobilityfromeutra",
        )
    ) or p in {
        "a3_Offset",
        "a6_Offset_r10",
        "hysteresis",
        "timeToTrigger",
        "triggerQuantity",
        "reportQuantity",
        "maxReportCells",
        "reportInterval",
        "reportAmount",
        "threshold_RSRP",
        "threshold_RSRQ",
        "s_Measure",
        "filterCoefficientRSRP",
        "filterCoefficientRSRQ",
        "allowedMeasBandwidth",
        "offsetFreq",
        "carrierFreq",
        "t304",
        "reportOnLeave",
        "purpose",
        "measCycleSCell_r10",
    }:
        if "eventa3" in pl or p == "a3_Offset":
            return "Mobility", "Event A3 (HO)"
        if "eventa1" in pl:
            return "Mobility", "Event A1"
        if "eventa2" in pl:
            return "Mobility", "Event A2"
        if "eventa4" in pl:
            return "Mobility", "Event A4"
        if "eventa5" in pl:
            return "Mobility", "Event A5"
        if "eventa6" in pl or p in {"a6_Offset_r10", "a6_ReportOnLeave_r10"}:
            return "Mobility", "Event A6 (SCell)"
        if "eventb1" in pl or p in {"b1_Threshold", "threshold_UTRA_RSCP", "threshold_GERAN"}:
            return "Mobility", "Event B1 (IRAT)"
        if "eventb2" in pl or p in {"b2_Threshold1", "b2_Threshold2"}:
            return "Mobility", "Event B2 (IRAT)"
        if p == "cellIndividualOffset":
            return "Mobility", "Cell Individual Offset"
        if p in {"gapOffset", "gp0", "gp1", "measGapConfig"} or "measgap" in pl:
            return "Mobility", "Meas Gap"
        if p in {"handoverType", "keyChangeIndicator", "nextHopChainingCount", "t304"} or "mobilitycontrol" in pl:
            return "Mobility", "HO Control (mobilityControlInfo)"
        if "measobject" in pl or p in {
            "carrierFreq",
            "dl_CarrierFreq",
            "dl_CarrierFreq_r10",
            "dl_CarrierFreq_r12",
            "ul_CarrierFreq",
            "allowedMeasBandwidth",
            "allowedMeasBandwidth_r12",
            "offsetFreq",
            "presenceAntennaPort1",
            "presenceAntennaPort1_r12",
            "measCycleSCell_r10",
            "bandIndicator",
        }:
            return "Mobility", "Meas Object"
        if "quantityconfig" in pl or p in {"filterCoefficientRSRP", "filterCoefficientRSRQ"}:
            return "Mobility", "Quantity Config"
        if p in {"reportAddNeighMeas_r10", "triggerType"}:
            return "Mobility", "Report Config"
        if p == "s_Measure":
            return "Mobility", "s-Measure"
        return "Mobility", "Measurement / Mobility"

    # Others
    if "drx" in pl or p in {"onDurationTimer", "drx_InactivityTimer", "drx_RetransmissionTimer", "longDRX_Cycle", "longDRX_CycleStartOffset", "shortDRX_Cycle", "drxShortCycleTimer"}:
        return "Others", "Connected DRX"
    if "mac_mainconfig" in pl or p in {"maxHARQ_Tx", "periodicBSR_Timer", "retxBSR_Timer", "ttiBundling", "timeAlignmentTimerDedicated"}:
        return "Others", "MAC Main Config"
    if "schedulingrequest" in pl or p in {"sr_ConfigIndex", "dsr_TransMax"}:
        return "Others", "Scheduling Request"
    if "antennainfo" in pl or p in {"transmissionMode"}:
        return "Others", "Antenna / TM"
    if "cqi_report" in pl or "pusch_configdedicated" in pl or p.startswith("betaOffset"):
        return "Others", "CQI / PUSCH Dedicated"
    if "rlc" in pl or p in {"t_PollRetransmit", "pollPDU", "pollByte", "maxRetxThreshold", "t_Reordering", "t_StatusProhibit"}:
        return "Others", "RLC"
    if "logicalchannel" in pl or p in {"priority", "prioritisedBitRate", "bucketSizeDuration", "logicalChannelGroup"}:
        return "Others", "Logical Channel"
    if "rach_config" in pl or p in {"numberOfRA_Preambles", "sizeOfRA_PreamblesGroupA", "messageSizeGroupA", "preambleTransMax", "ra_ResponseWindowSize", "mac_ContentionResolutionTimer", "maxHARQ_Msg3Tx", "highSpeedFlag", "zeroCorrelationZoneConfig"}:
        return "Others", "RACH Common"
    if "security" in pl or p in {"cipheringAlgorithm", "integrityProtAlgorithm"}:
        return "Others", "Security"
    if "release" in pl or p in {"releaseCause", "redirectedCarrierInfo", "extendedWaitTime_r10", "t320"}:
        return "Others", "RRC Release"
    if "scell" in pl or p == "sCellDeactivationTimer_r10":
        return "Others", "Carrier Aggregation (SCell)"
    if p in {"tpc_PDCCH_ConfigPUCCH", "tpc_PDCCH_ConfigPUSCH"} or "tpc_pdcch" in pl:
        return "Others", "TPC (group)"
    if "soundingrs" in pl or p in {
        "soundingRS_UL_ConfigDedicated",
        "srs_Bandwidth",
        "srs_HoppingBandwidth",
        "freqDomainPosition",
        "duration",
        "transmissionComb",
    }:
        return "Others", "SRS"
    if "sps" in pl or p in {
        "semiPersistSchedIntervalDL",
        "semiPersistSchedIntervalUL",
        "numberOfConfSPS_Processes",
        "implicitReleaseAfter",
    }:
        return "Others", "SPS"
    if p in {"sn_FieldLength", "headerCompression", "discardTimer", "pdcp_SN_Size", "statusReportRequired"}:
        return "Others", "PDCP / RLC UM"
    if p == "sr_ProhibitTimer_r9":
        return "Others", "Scheduling Request"
    if p in {"phich_Resource", "phich_Duration", "dl_Bandwidth_r10", "antennaPortsCount"}:
        return "Others", "PHICH / Antenna Common"
    return "Others", "Other RRC"


def interpret(param: str, raw: str) -> str:
    """Human-readable meaning for common LTE mapped values."""
    m = re.match(r"^([-\d]+)", str(raw).strip())
    num = int(m.group(1)) if m else None
    enum = str(raw).split("\t")[0].strip() if "\t" in str(raw) else str(raw).strip()
    # strip trailing (n)
    enum = re.sub(r"\s*\(\d+\)\s*$", "", enum).strip()

    if param in {"threshold_RSRP"} and num is not None:
        return f"{enum} → {num - 140} dBm"
    if param in {"threshold_RSRQ"} and num is not None:
        return f"{enum} → {(num - 40) / 2:.1f} dB"
    if param == "a3_Offset" and num is not None:
        return f"{enum} → {num * 0.5:g} dB"
    if param == "a6_Offset_r10" and num is not None:
        return f"{enum} → {num * 0.5:g} dB"
    if param == "hysteresis" and num is not None:
        return f"{enum} → {num * 0.5:g} dB"
    if param == "q_RxLevMin" and num is not None:
        return f"{enum} → {num * 2} dBm"
    if param in {"s_IntraSearch", "s_NonIntraSearch", "threshServingLow", "threshX_High", "threshX_Low"} and num is not None:
        return f"{enum} → {num * 2} dB"
    if param == "s_Measure" and num is not None:
        if num == 0:
            return f"{enum} → disabled (always measure)"
        return f"{enum} → {num - 140} dBm"
    if param in {"p0_NominalPUSCH", "p0_NominalPUCCH", "p0_UE_PUSCH", "p0_UE_PUCCH"} and num is not None:
        return f"{enum} dBm"
    if param == "referenceSignalPower" and num is not None:
        return f"{enum} dBm"
    if param == "deltaPreambleMsg3" and num is not None:
        return f"{enum} → {num * 2} dB"
    if param == "nomPDSCH_RS_EPRE_Offset" and num is not None:
        return f"{enum} → {num} dB"
    return enum


def iter_messages(path: Path):
    cur = None
    buf: list[str] = []
    with path.open("r", errors="replace") as f:
        for line in f:
            m = HDR.match(line.rstrip("\n"))
            if m:
                if cur is not None:
                    yield cur, "".join(buf)
                t, d, layer, name, tech, side, sid, op = m.groups()
                cur = {
                    "time": t,
                    "dir": d,
                    "layer": layer,
                    "name": name.strip(),
                    "tech": tech,
                    "side": side,
                    "sid": sid,
                    "op": op,
                    "file": path.name,
                }
                buf = [line]
            else:
                if cur is not None:
                    buf.append(line)
        if cur is not None:
            yield cur, "".join(buf)


def extract_leaves(text: str) -> list[tuple[str, str, str]]:
    """Return list of (path, param, value) from a decoded message body."""
    lines = text.splitlines()
    stack: list[tuple[int, str]] = []
    out: list[tuple[str, str, str]] = []
    leaf_re = re.compile(r"^(\s*)([A-Za-z][A-Za-z0-9_]*)\s*:\s*(.+?)\s*$")
    node_re = re.compile(r"^(\s*)([A-Za-z][A-Za-z0-9_]*)\s*$")
    idx_re = re.compile(r"^(\s*)\[(\d+)\]\s+([A-Za-z][A-Za-z0-9_]*)\s*$")
    elem_re = re.compile(r"^(\s*)element(?:\s*:\s*(.+))?\s*$")

    for line in lines:
        if "\tEncoded" in line or line.strip() == "Encoded":
            break
        if line.strip() in {"Decoded", ""}:
            continue
        # indexed list item
        m = idx_re.match(line)
        if m:
            indent, idx, name = m.groups()
            depth = len(indent)
            while stack and stack[-1][0] >= depth:
                stack.pop()
            stack.append((depth, f"{name}[{idx}]"))
            continue
        m = elem_re.match(line)
        if m:
            indent = m.group(1)
            depth = len(indent)
            while stack and stack[-1][0] >= depth:
                stack.pop()
            label = "element"
            if m.group(2):
                label = f"element={m.group(2).strip()}"
            stack.append((depth, label))
            continue
        m = leaf_re.match(line)
        if m:
            indent, key, val = m.groups()
            # skip nested enum wrappers that are only selectors with children following
            depth = len(indent)
            while stack and stack[-1][0] >= depth:
                stack.pop()
            # Detect if value looks like a choice marker with child block: "c1 (0)" etc. still a leaf if no deeper same-family; keep as leaf.
            path = "/".join(s[1] for s in stack)
            out.append((path, key, val.strip()))
            # Also push as context for nested threshold fields
            stack.append((depth, key))
            continue
        m = node_re.match(line)
        if m:
            indent, key = m.groups()
            depth = len(indent)
            while stack and stack[-1][0] >= depth:
                stack.pop()
            stack.append((depth, key))
            continue
    return out


def keep_param(param: str) -> bool:
    if param in PARAM_WHITELIST:
        return True
    if param.startswith(PARAM_BLACKLIST_PREFIX):
        return False
    # keep event threshold wrapper names indirectly via threshold_RSRP
    if param.endswith("_Threshold") or param.endswith("_Threshold1") or param.endswith("_Threshold2"):
        return False
    return False


CONFIG_MESSAGES = {
    "DCCH-RRCConnectionReconfiguration",
    "CCCH-RRCConnectionSetup",
    "BCCH:SCH-SystemInformation",
    "BCCH:SCH-SystemInformationBlockType1",
    "DCCH-MobilityFromEUTRACommand",
    "DCCH-RRCConnectionRelease",
    "DCCH-SecurityModeCommand",
}


def _clean_val(v: str) -> str:
    return re.sub(r"\s*\(\d+\)\s*$", "", str(v).split("\t")[0]).strip()


def extract_event_profiles(text: str) -> list[tuple[str, dict[str, str]]]:
    """Pull compact event A1–A6 profiles from one reconfiguration body."""
    profiles: list[tuple[str, dict[str, str]]] = []
    # Split roughly by eventId lines
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        m = re.search(r"eventId\s*:\s*(eventA[0-9]_?(?:r10)?)", lines[i])
        if not m:
            i += 1
            continue
        event = m.group(1)
        block = "\n".join(lines[i : i + 25])
        fields: dict[str, str] = {"event": event}
        for key in (
            "a3_Offset",
            "a6_Offset_r10",
            "hysteresis",
            "timeToTrigger",
            "threshold_RSRP",
            "threshold_RSRQ",
            "triggerQuantity",
            "reportQuantity",
            "maxReportCells",
            "reportInterval",
            "reportAmount",
            "reportOnLeave",
            "a6_ReportOnLeave_r10",
        ):
            km = re.search(rf"{key}\s*:\s*([^\n]+)", block)
            if km:
                fields[key] = _clean_val(km.group(1))
        # a5 has two thresholds — capture both in order
        ths = re.findall(r"threshold_RSRP\s*:\s*([^\n]+)", block)
        if event.startswith("eventA5") and len(ths) >= 2:
            fields["a5_Threshold1_RSRP"] = _clean_val(ths[0])
            fields["a5_Threshold2_RSRP"] = _clean_val(ths[1])
            fields.pop("threshold_RSRP", None)
        elif "threshold_RSRP" in fields and event.startswith("eventA"):
            pass
        profiles.append((event, fields))
        i += 1
    return profiles


def scan(l3_dir: Path):
    """Aggregate parameter value counters per operator."""
    # key: (section, group, param) -> op -> Counter(raw_value)
    values: dict[tuple[str, str, str], dict[str, Counter]] = defaultdict(
        lambda: {op: Counter() for op in OPS}
    )
    msg_counts = {op: Counter() for op in OPS}
    event_profiles: dict[str, dict[str, Counter]] = defaultdict(
        lambda: {op: Counter() for op in OPS}
    )
    files = sorted(l3_dir.glob("*.txt"))
    if not files:
        raise SystemExit(f"No L3 .txt files under {l3_dir}")

    for fp in files:
        for meta, text in iter_messages(fp):
            if meta["op"] not in OPS:
                continue
            msg_counts[meta["op"]][meta["name"]] += 1
            if meta["dir"] != "D":
                continue
            if meta["name"] not in CONFIG_MESSAGES:
                continue
            for path, param, raw in extract_leaves(text):
                if not keep_param(param):
                    continue
                section, group = _classify(param, path + "/" + meta["name"])
                values[(section, group, param)][meta["op"]][raw] += 1
            if meta["name"] == "DCCH-RRCConnectionReconfiguration":
                for event, fields in extract_event_profiles(text):
                    # Stable signature string for counting
                    ordered = []
                    for k in (
                        "a3_Offset",
                        "a6_Offset_r10",
                        "threshold_RSRP",
                        "a5_Threshold1_RSRP",
                        "a5_Threshold2_RSRP",
                        "hysteresis",
                        "timeToTrigger",
                        "triggerQuantity",
                        "maxReportCells",
                        "reportInterval",
                        "reportAmount",
                    ):
                        if k in fields:
                            ordered.append(f"{k}={fields[k]}")
                    sig = "; ".join(ordered) if ordered else "(empty)"
                    event_profiles[event][meta["op"]][sig] += 1

    return values, msg_counts, event_profiles


def top_values(counter: Counter, n: int = 5) -> str:
    if not counter:
        return "— (not seen)"
    parts = []
    total = sum(counter.values())
    for val, cnt in counter.most_common(n):
        pct = 100.0 * cnt / total
        parts.append(f"{val}  [{cnt}, {pct:.0f}%]")
    return "\n".join(parts)


def top_interpreted(param: str, counter: Counter, n: int = 5) -> str:
    if not counter:
        return "—"
    parts = []
    total = sum(counter.values())
    for val, cnt in counter.most_common(n):
        pct = 100.0 * cnt / total
        parts.append(f"{interpret(param, val)}  [{cnt}, {pct:.0f}%]")
    return "\n".join(parts)


def top_merged(param: str, counter: Counter, n: int = 5) -> str:
    """Raw IE and mapped unit in one line, with the sample share."""
    if not counter:
        return "— (not seen)"
    return top_interpreted(param, counter, n)


def dominant(counter: Counter) -> str:
    if not counter:
        return ""
    return counter.most_common(1)[0][0]


def compare_status(robi: Counter, gp: Counter) -> str:
    if not robi and not gp:
        return "Not seen"
    if robi and not gp:
        return "Robi only"
    if gp and not robi:
        return "GP only"
    r = dominant(robi)
    g = dominant(gp)
    # compare interpreted enums without counts
    ri = interpret("x", r).split("→")[0].strip() if False else interpret(
        # use param-agnostic strip of (n)
        "hysteresis" if False else "dummy",
        r,
    )
    # Better: compare cleaned raw
    def clean(v: str) -> str:
        return re.sub(r"\s*\(\d+\)\s*$", "", v.split("\t")[0]).strip()

    if clean(r) == clean(g):
        # check if distributions are essentially same
        return "Same (dominant)"
    return "Different"


def compare_status_param(param: str, robi: Counter, gp: Counter) -> str:
    if not robi and not gp:
        return "Not seen"
    if robi and not gp:
        return "Robi only"
    if gp and not robi:
        return "GP only"

    def clean(v: str) -> str:
        return re.sub(r"\s*\(\d+\)\s*$", "", str(v).split("\t")[0]).strip()

    r = clean(dominant(robi))
    g = clean(dominant(gp))
    if r == g:
        return "Same (dominant)"
    return "Different"


def style_header(ws, row: int, cols: int):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin


def apply_body(ws, start_row: int, end_row: int, cols: int):
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    fills = {
        "Different": PatternFill("solid", fgColor="FCE4D6"),
        "Same (dominant)": PatternFill("solid", fgColor="E2EFDA"),
        "Robi only": PatternFill("solid", fgColor="DDEBF7"),
        "GP only": PatternFill("solid", fgColor="FFF2CC"),
    }
    section_fills = {
        "Mobility": PatternFill("solid", fgColor="D6EAF8"),
        "Power": PatternFill("solid", fgColor="F5B7B1"),
        "Idle mode": PatternFill("solid", fgColor="D5F5E3"),
        "Others": PatternFill("solid", fgColor="E8DAEF"),
    }
    for r in range(start_row, end_row + 1):
        status = ws.cell(r, 6).value
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.border = thin
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.font = Font(name="Calibri", size=10)
        if status in fills:
            ws.cell(r, 6).fill = fills[status]
            ws.cell(r, 6).font = Font(name="Calibri", size=10, bold=True)
        sec = ws.cell(r, 1).value
        if sec in section_fills:
            ws.cell(r, 1).fill = section_fills[sec]


def build_workbook(values, msg_counts, event_profiles, out_path: Path):
    wb = Workbook()

    # --- Cover ---
    cover = wb.active
    cover.title = "Cover"
    cover["A1"] = "L3 REPORT FOR BOGRA"
    cover["A1"].font = Font(name="Calibri", size=22, bold=True, color="1F4E78")
    cover["A2"] = "Layer-3 parameter comparison — Robi vs Grameenphone (GP)"
    cover["A2"].font = Font(name="Calibri", size=14, color="5B5B5B")
    cover["A4"] = "Source"
    cover["B4"] = "Layer 3_Active Mode_Data_Robi_v1.7z (Active Mode decoded LTE-RRC)"
    cover["A5"] = "Operators compared"
    cover["B5"] = "Robi  vs  Grameenphone (GP)   — first pair (Mema)"
    cover["A6"] = "Aligned with"
    cover["B6"] = "cursor/active-mode-11nov26-6bd5  ·  Bogura drive-test pipeline (Robi focus + multi-op L3 compare)"
    cover["A7"] = "Section order"
    cover["B7"] = "1) Mobility   2) Power   3) Idle mode   4) Others"
    cover["A9"] = "How to read the Comparison sheet"
    cover["A9"].font = Font(bold=True, size=12, color="1F4E78")
    cover["A10"] = (
        "Each row is one L3 parameter. Robi Value and GP Value show the dominant setting "
        "seen in downlink configuration messages (RRC Connection Setup / Reconfiguration, SIB1/SI, "
        "MobilityFromEUTRA, SecurityMode, Release), with up to 5 most common values and their share. "
        "Meaning converts mapped 3GPP fields (e.g. RSRP index → dBm, a3-Offset → dB). "
        "Status: Same (dominant) / Different / Robi only / GP only."
    )
    cover["A10"].alignment = Alignment(wrap_text=True)
    cover.merge_cells("A10:F12")

    cover["A14"] = "Message volume in L3 dump"
    cover["A14"].font = Font(bold=True, size=12, color="1F4E78")
    cover["A15"] = "Message"
    cover["B15"] = "Robi"
    cover["C15"] = "GP"
    style_header(cover, 15, 3)
    all_msgs = sorted(set(msg_counts["Robi"]) | set(msg_counts["Grameenphone"]))
    r = 16
    for name in all_msgs:
        cover.cell(r, 1, name)
        cover.cell(r, 2, msg_counts["Robi"].get(name, 0))
        cover.cell(r, 3, msg_counts["Grameenphone"].get(name, 0))
        r += 1
    cover.cell(r, 1, "TOTAL messages")
    cover.cell(r, 2, sum(msg_counts["Robi"].values()))
    cover.cell(r, 3, sum(msg_counts["Grameenphone"].values()))
    for c, w in enumerate([42, 55, 18, 18, 18, 18], 1):
        cover.column_dimensions[get_column_letter(c)].width = w

    # --- Comparison ---
    ws = wb.create_sheet("Comparison", 1)
    headers = [
        "Section",
        "Group",
        "Parameter",
        "Robi Value (dominant + top)",
        "GP Value (dominant + top)",
        "Status",
        "Robi Meaning",
        "GP Meaning",
        "Robi hits",
        "GP hits",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    style_header(ws, 1, len(headers))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # Sort rows by section order, group, param
    rows = []
    for (section, group, param), by_op in values.items():
        robi = by_op["Robi"]
        gp = by_op["Grameenphone"]
        if not robi and not gp:
            continue
        rows.append((section, group, param, robi, gp))

    def sort_key(item):
        section, group, param, *_ = item
        try:
            si = SECTION_ORDER.index(section)
        except ValueError:
            si = 99
        return (si, group, param)

    rows.sort(key=sort_key)

    r = 2
    for section, group, param, robi, gp in rows:
        ws.cell(r, 1, section)
        ws.cell(r, 2, group)
        ws.cell(r, 3, param)
        ws.cell(r, 4, top_values(robi))
        ws.cell(r, 5, top_values(gp))
        ws.cell(r, 6, compare_status_param(param, robi, gp))
        ws.cell(r, 7, top_interpreted(param, robi))
        ws.cell(r, 8, top_interpreted(param, gp))
        ws.cell(r, 9, sum(robi.values()))
        ws.cell(r, 10, sum(gp.values()))
        r += 1
    apply_body(ws, 2, r - 1, len(headers))
    widths = [12, 34, 28, 42, 42, 16, 36, 36, 10, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30

    # --- Per-section sheets ---
    for section in SECTION_ORDER:
        sw = wb.create_sheet(section[:31])
        for c, h in enumerate(headers, 1):
            sw.cell(1, c, h)
        style_header(sw, 1, len(headers))
        sw.freeze_panes = "A2"
        rr = 2
        for section2, group, param, robi, gp in rows:
            if section2 != section:
                continue
            sw.cell(rr, 1, section2)
            sw.cell(rr, 2, group)
            sw.cell(rr, 3, param)
            sw.cell(rr, 4, top_values(robi))
            sw.cell(rr, 5, top_values(gp))
            sw.cell(rr, 6, compare_status_param(param, robi, gp))
            sw.cell(rr, 7, top_interpreted(param, robi))
            sw.cell(rr, 8, top_interpreted(param, gp))
            sw.cell(rr, 9, sum(robi.values()))
            sw.cell(rr, 10, sum(gp.values()))
            rr += 1
        if rr > 2:
            apply_body(sw, 2, rr - 1, len(headers))
        for i, w in enumerate(widths, 1):
            sw.column_dimensions[get_column_letter(i)].width = w

    # --- Diff only ---
    diff = wb.create_sheet("Diff only")
    for c, h in enumerate(headers, 1):
        diff.cell(1, c, h)
    style_header(diff, 1, len(headers))
    diff.freeze_panes = "A2"
    rr = 2
    for section, group, param, robi, gp in rows:
        st = compare_status_param(param, robi, gp)
        if st not in {"Different", "Robi only", "GP only"}:
            continue
        diff.cell(rr, 1, section)
        diff.cell(rr, 2, group)
        diff.cell(rr, 3, param)
        diff.cell(rr, 4, top_values(robi))
        diff.cell(rr, 5, top_values(gp))
        diff.cell(rr, 6, st)
        diff.cell(rr, 7, top_interpreted(param, robi))
        diff.cell(rr, 8, top_interpreted(param, gp))
        diff.cell(rr, 9, sum(robi.values()))
        diff.cell(rr, 10, sum(gp.values()))
        rr += 1
    if rr > 2:
        apply_body(diff, 2, rr - 1, len(headers))
    for i, w in enumerate(widths, 1):
        diff.column_dimensions[get_column_letter(i)].width = w

    # --- Summary counts ---
    summary = wb.create_sheet("Summary", 2)
    summary["A1"] = "L3 REPORT FOR BOGRA — Summary"
    summary["A1"].font = Font(size=16, bold=True, color="1F4E78")
    summary["A3"] = "Section"
    summary["B3"] = "Parameters"
    summary["C3"] = "Same"
    summary["D3"] = "Different"
    summary["E3"] = "Robi only"
    summary["F3"] = "GP only"
    style_header(summary, 3, 6)
    from collections import Counter as C

    sec_stat = {s: C() for s in SECTION_ORDER}
    for section, group, param, robi, gp in rows:
        st = compare_status_param(param, robi, gp)
        sec_stat[section][st] += 1
        sec_stat[section]["Parameters"] += 1
    rr = 4
    for section in SECTION_ORDER:
        summary.cell(rr, 1, section)
        summary.cell(rr, 2, sec_stat[section]["Parameters"])
        summary.cell(rr, 3, sec_stat[section]["Same (dominant)"])
        summary.cell(rr, 4, sec_stat[section]["Different"])
        summary.cell(rr, 5, sec_stat[section]["Robi only"])
        summary.cell(rr, 6, sec_stat[section]["GP only"])
        rr += 1
    summary.cell(rr, 1, "TOTAL")
    for c in range(2, 7):
        summary.cell(rr, c, sum(summary.cell(i, c).value or 0 for i in range(4, rr)))
    for i, w in enumerate([14, 14, 12, 12, 12, 12], 1):
        summary.column_dimensions[get_column_letter(i)].width = w

    summary["A10"] = "Highlights (Mobility — first look)"
    summary["A10"].font = Font(bold=True, size=12, color="1F4E78")
    # Pull a few key mobility diffs for the cover note
    highlights = []
    key_params = [
        ("a3_Offset", "Event A3 offset"),
        ("hysteresis", "Hysteresis"),
        ("timeToTrigger", "Time To Trigger"),
        ("s_Measure", "s-Measure"),
        ("q_Hyst", "q-Hyst (idle)"),
        ("s_IntraSearch", "s-IntraSearch"),
        ("p0_NominalPUSCH", "p0-NominalPUSCH"),
        ("alpha", "alpha"),
        ("p_a", "PDSCH p-a"),
        ("defaultPagingCycle", "defaultPagingCycle"),
    ]
    # index rows by param (may appear in multiple groups — take first mobility/power/idle)
    by_param = defaultdict(list)
    for section, group, param, robi, gp in rows:
        by_param[param].append((section, group, robi, gp))
    hr = 11
    summary.cell(hr, 1, "Parameter")
    summary.cell(hr, 2, "Robi (dominant)")
    summary.cell(hr, 3, "GP (dominant)")
    summary.cell(hr, 4, "Status")
    style_header(summary, hr, 4)
    hr = 12
    for param, label in key_params:
        if param not in by_param:
            continue
        section, group, robi, gp = by_param[param][0]
        rd = interpret(param, dominant(robi)) if robi else "—"
        gd = interpret(param, dominant(gp)) if gp else "—"
        summary.cell(hr, 1, f"{label} ({param})")
        summary.cell(hr, 2, rd)
        summary.cell(hr, 3, gd)
        summary.cell(hr, 4, compare_status_param(param, robi, gp))
        hr += 1
    apply_body(summary, 12, hr - 1, 4)
    summary.column_dimensions["A"].width = 36
    summary.column_dimensions["B"].width = 40
    summary.column_dimensions["C"].width = 40
    summary.column_dimensions["D"].width = 16

    # --- Event Profiles (Mobility at a glance) ---
    ev = wb.create_sheet("Event Profiles", 3)
    ev["A1"] = "Mobility Event Profiles — Robi vs GP"
    ev["A1"].font = Font(size=14, bold=True, color="1F4E78")
    ev["A2"] = (
        "Each row is one distinct event configuration seen in RRCConnectionReconfiguration. "
        "RSRP thresholds are shown as mapped index (dBm = index − 140). "
        "a3-Offset / hysteresis: dB = value × 0.5."
    )
    ev["A2"].alignment = Alignment(wrap_text=True)
    ev.merge_cells("A2:F3")
    ev_headers = ["Event", "Operator", "Rank", "Count", "Share %", "Profile (key settings)"]
    for c, h in enumerate(ev_headers, 1):
        ev.cell(5, c, h)
    style_header(ev, 5, len(ev_headers))
    er = 6
    for event in sorted(event_profiles.keys()):
        for op in OPS:
            counter = event_profiles[event][op]
            if not counter:
                continue
            total = sum(counter.values())
            for rank, (sig, cnt) in enumerate(counter.most_common(8), 1):
                ev.cell(er, 1, event)
                ev.cell(er, 2, OP_LABEL[op])
                ev.cell(er, 3, rank)
                ev.cell(er, 4, cnt)
                ev.cell(er, 5, round(100.0 * cnt / total, 1))
                # Add dBm/dB hints inline for readability
                pretty = sig
                pretty = re.sub(
                    r"threshold_RSRP=([-]?\d+)",
                    lambda m: f"threshold_RSRP={m.group(1)} ({int(m.group(1))-140} dBm)",
                    pretty,
                )
                pretty = re.sub(
                    r"a5_Threshold([12])_RSRP=([-]?\d+)",
                    lambda m: f"a5_Th{m.group(1)}={m.group(2)} ({int(m.group(2))-140} dBm)",
                    pretty,
                )
                pretty = re.sub(
                    r"a3_Offset=([-]?\d+)",
                    lambda m: f"a3_Offset={m.group(1)} ({int(m.group(1))*0.5:g} dB)",
                    pretty,
                )
                pretty = re.sub(
                    r"hysteresis=([-]?\d+)",
                    lambda m: f"hysteresis={m.group(1)} ({int(m.group(1))*0.5:g} dB)",
                    pretty,
                )
                ev.cell(er, 6, pretty)
                er += 1
    if er > 6:
        apply_body(ev, 6, er - 1, len(ev_headers))
    for i, w in enumerate([14, 22, 8, 10, 10, 100], 1):
        ev.column_dimensions[get_column_letter(i)].width = w

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path, len(rows), sec_stat


# Identity, geography, and per-cell indexes. They describe where the UE was,
# not a setting an operator chose once for the network.
_NOT_A_POLICY = {
    "carrierFreq",
    "dl_CarrierFreq",
    "ul_CarrierFreq",
    "dl_CarrierFreq_r10",
    "dl_CarrierFreq_r12",
    "targetPhysCellId",
    "startingARFCN",
    "freqBandIndicator",
    "bandIndicator",
    "rootSequenceIndex",
    "gp0",
    "gp1",
    "gapOffset",
    "nextHopChainingCount",
    "keyChangeIndicator",
    "handoverType",
    "codebookSubsetRestriction",
    "sr_ConfigIndex",
    "priority",
    "logicalChannelGroup",
    "cellIndividualOffset",
    "neighCellConfig",
    "systemInfoValueTag",
}

_WHY = {
    "s_NonIntraSearch": (
        "This is the idle threshold for looking at other LTE layers. "
        "With q-RxLevMin at -128 dBm, Robi keeps measuring other layers until serving RSRP is about -74 dBm. "
        "GP stops that search once serving RSRP is only about -118 dBm."
    ),
    "threshServingLow": (
        "Idle rule for leaving a higher-priority layer. "
        "Robi treats the serving cell as low from about -116 dBm. GP waits until about -122 dBm."
    ),
    "s_IntraSearch": (
        "Both still measure the same frequency across almost the whole drive. "
        "The 4 dB gap (about -66 dBm vs -70 dBm) is small next to the inter-layer gap."
    ),
    "cellReselectionPriority": (
        "Absolute priority of the serving frequency the UE was camped on. "
        "Robi's typical camped layer is priority 5. GP's is priority 3. This follows the band plan, not a single timer."
    ),
    "t300": (
        "How long the UE waits for RRC setup before it gives up. "
        "Robi allows 1 s. On most GP cells the wait is 200 ms, so a weak setup fails and retries sooner."
    ),
    "threshold_RSRP": (
        "Event A4: a neighbour must beat this level to be reported. "
        "Robi accepts a neighbour around -124 dBm. GP asks for about -116 dBm, so the report starts later."
    ),
    "measCycleSCell_r10": (
        "How often a deactivated secondary cell is measured. "
        "Robi uses 320 ms on most carriers. The GP samples that carry this field use 640 ms."
    ),
    "maxHARQ_Msg3Tx": (
        "How many times Msg3 is retransmitted during access. GP allows more attempts (8 vs 5)."
    ),
    "pusch_HoppingOffset": (
        "PUSCH frequency-hopping offset in the common config. Robi uses 0. GP uses 22 on most cells."
    ),
    "zeroCorrelationZoneConfig": (
        "PRACH zero-correlation zone. One step apart (10 vs 11). A cell-radius planning choice, not a mobility policy."
    ),
    "specialSubframePatterns": (
        "TDD special-subframe pattern on the cells that broadcast it. Robi ssp7, GP ssp5."
    ),
}


def _clean_raw(v: str) -> str:
    return re.sub(r"\s*\(\d+\)\s*$", "", str(v).split("\t")[0]).strip()


def _dom(counter: Counter):
    if not counter:
        return "", 0.0, 0
    total = sum(counter.values())
    val, cnt = counter.most_common(1)[0]
    return val, cnt / total, total


def _policy_rows(values):
    """One row per distinct (parameter, Robi value, GP value), keeping the largest sample."""
    best = {}
    for (section, group, param), by_op in values.items():
        if param in _NOT_A_POLICY:
            continue
        robi, gp = by_op["Robi"], by_op["Grameenphone"]
        rv, rs, nr = _dom(robi)
        gv, gs, ng = _dom(gp)
        if nr < 40 or ng < 40:
            continue
        status = compare_status_param(param, robi, gp)
        key = (param, _clean_raw(rv), _clean_raw(gv))
        score = min(nr, ng)
        prev = best.get(key)
        if prev is None or score > prev["score"]:
            best[key] = {
                "section": section,
                "group": group,
                "param": param,
                "nr": nr,
                "ng": ng,
                "rs": rs,
                "gs": gs,
                "robi": interpret(param, rv),
                "gp": interpret(param, gv),
                "status": status,
                "score": score,
                "why": _WHY.get(param, ""),
            }
    return list(best.values())


def _write_md(path: Path, differ, shared, mixed, msg_counts):
    lines = [
        "# L3 report for Bogura",
        "",
        "Robi and Grameenphone, from the Active Mode Layer-3 decode "
        "(`Layer 3_Active Mode_Data_Robi_v1.7z`).",
        "",
        "A value is treated as the operator setting only when it is the same on at least 70% of the cells that broadcast it, and both operators have at least 40 samples. Frequencies, PCI, and per-cell indexes are left out of that table. They describe the drive, not a policy.",
        "",
        f"Messages in the dump: Robi {sum(msg_counts['Robi'].values()):,}, Grameenphone {sum(msg_counts['Grameenphone'].values()):,}.",
        "",
        "## What is different",
        "",
        "The idle UE is where the two networks do not behave the same.",
        "",
        "Robi keeps searching other LTE layers until the serving cell is strong (s-NonIntraSearch 54 dB, about -74 dBm given q-RxLevMin -128 dBm). Grameenphone stops that search once the serving cell is merely above the minimum (10 dB, about -118 dBm). Robi also calls the serving cell low sooner (threshServingLow 12 dB vs 6 dB) and waits 1 s for RRC setup. Most GP cells wait 200 ms.",
        "",
        "Connected handover is much closer. A3 offset, hysteresis, time-to-trigger, s-Measure, and the L3 RSRP/RSRQ filter are the same. Uplink power (p0-NominalPUSCH -67 dBm, alpha 0.7, p0-NominalPUCCH -115 dBm) is the same. Paging (rf128) is the same. Security is eea2 / eia2 on both.",
        "",
        "| What changes | Robi | GP | How stable |",
        "| --- | --- | --- | --- |",
    ]
    for row in differ:
        lines.append(
            f"| {row['param']} | {row['robi']} | {row['gp']} | "
            f"Robi {row['rs']:.0%} of {row['nr']}, GP {row['gs']:.0%} of {row['ng']} |"
        )
    lines += [
        "",
        "### Why those rows matter",
        "",
    ]
    for row in differ:
        if row["why"]:
            lines.append(f"- **{row['param']}.** {row['why']}")
    lines += [
        "",
        "## What is the same",
        "",
        f"{len(shared)} settings are the same dominant value on both networks, each on at least 90% of samples.",
        "",
        "| Setting | Value | Robi samples | GP samples |",
        "| --- | --- | ---: | ---: |",
    ]
    for row in shared:
        lines.append(f"| {row['param']} | {row['robi']} | {row['nr']} | {row['ng']} |")
    lines += [
        "",
        "## Not one network-wide value",
        "",
        "These are configured, but the dominant value is under 70% on at least one operator. Do not quote a single number for them.",
        "",
        "| Setting | Robi mode | GP mode |",
        "| --- | --- | --- |",
    ]
    for row in mixed[:25]:
        lines.append(
            f"| {row['param']} | {row['robi']} ({row['rs']:.0%}) | {row['gp']} ({row['gs']:.0%}) |"
        )
    lines += [
        "",
        "Workbook: `output/L3_REPORT_FOR_BOGRA.xlsx`.",
        "",
    ]
    path.write_text("\n".join(lines))


def build_analyst_workbook(values, msg_counts, event_profiles, out_path: Path):
    rows = _policy_rows(values)
    differ = [
        r for r in rows
        if r["status"] == "Different" and r["rs"] >= 0.70 and r["gs"] >= 0.70
    ]
    # Lead with the idle behaviour gap, then setup, then the smaller radio deltas.
    lead = [
        "s_NonIntraSearch",
        "threshServingLow",
        "s_IntraSearch",
        "cellReselectionPriority",
        "t300",
        "threshold_RSRP",
        "measCycleSCell_r10",
        "maxHARQ_Msg3Tx",
        "pusch_HoppingOffset",
        "zeroCorrelationZoneConfig",
        "specialSubframePatterns",
    ]
    differ.sort(key=lambda r: lead.index(r["param"]) if r["param"] in lead else 50)

    shared = [
        r for r in rows
        if r["status"].startswith("Same") and r["rs"] >= 0.90 and r["gs"] >= 0.90
    ]
    shared.sort(key=lambda r: (-min(r["nr"], r["ng"]), r["param"]))

    mixed = [
        r for r in rows
        if r["rs"] < 0.70 or r["gs"] < 0.70
    ]
    mixed.sort(key=lambda r: -min(r["nr"], r["ng"]))

    wb = Workbook()
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    navy = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    body = Font(name="Calibri", size=11)
    title_font = Font(name="Calibri", size=20, bold=True, color="1F4E78")

    cover = wb.active
    cover.title = "Findings"
    cover.sheet_view.showGridLines = False
    cover["A1"] = "L3 report for Bogura"
    cover["A1"].font = title_font
    cover["A2"] = "Robi and Grameenphone — what each network actually sets"
    cover["A2"].font = Font(name="Calibri", size=13, color="595959")
    cover.merge_cells("A4:B4")
    cover["A4"] = (
        "Source: Layer 3_Active Mode_Data_Robi_v1.7z, downlink RRC only "
        "(setup, reconfiguration, system information, handover, release). "
        f"Robi {sum(msg_counts['Robi'].values()):,} messages, "
        f"Grameenphone {sum(msg_counts['Grameenphone'].values()):,}."
    )
    cover["A4"].alignment = Alignment(wrap_text=True, vertical="top")
    cover.row_dimensions[4].height = 48

    cover["A6"] = "The difference that matters"
    cover["A6"].font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    cover.merge_cells("A7:B10")
    cover["A7"] = (
        "Idle mode, not handover and not power.\n\n"
        "Robi keeps the UE measuring other LTE layers until the serving cell is strong "
        "(s-NonIntraSearch 54 dB → about -74 dBm, with q-RxLevMin -128 dBm). "
        "Grameenphone stops that search once serving is only slightly above the camping minimum "
        "(10 dB → about -118 dBm). Robi also marks the serving cell as low earlier "
        "(threshServingLow 12 dB vs 6 dB) and allows 1 s for RRC setup. Most GP cells allow 200 ms.\n\n"
        "Connected A3 handover, uplink power, paging, and AS security are the same setting on both."
    )
    cover["A7"].alignment = Alignment(wrap_text=True, vertical="top")
    cover.row_dimensions[7].height = 36
    cover.row_dimensions[8].height = 36
    cover.row_dimensions[9].height = 36
    cover.row_dimensions[10].height = 36

    cover["A12"] = "How a number got into this report"
    cover["A12"].font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    cover.merge_cells("A13:B15")
    cover["A13"] = (
        "Both operators need at least 40 samples. The value shown is the one seen most often. "
        "It is called the setting only when that value covers at least 70% of samples. "
        "EARFCN, PCI, gap offset, and root sequence are omitted here: they change with the cell, "
        "they are not a network-wide choice. The Appendix sheet still lists every decoded leaf."
    )
    cover["A13"].alignment = Alignment(wrap_text=True, vertical="top")
    cover.column_dimensions["A"].width = 36
    cover.column_dimensions["B"].width = 88

    def paint_header(ws, headers):
        for c, h in enumerate(headers, 1):
            cell = ws.cell(1, c, h)
            cell.fill = navy
            cell.font = head_font
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
        ws.row_dimensions[1].height = 22
        ws.sheet_view.showGridLines = False

    def paint_row(ws, r, cols, fill=None):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.font = body
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if fill is not None:
                cell.fill = fill

    diff_ws = wb.create_sheet("Different settings")
    diff_headers = [
        "What it changes",
        "Parameter",
        "Robi",
        "Grameenphone",
        "Robi share",
        "GP share",
        "Why it matters",
    ]
    paint_header(diff_ws, diff_headers)
    amber = PatternFill("solid", fgColor="FDE9D9")
    for i, row in enumerate(differ, 2):
        diff_ws.cell(i, 1, row["group"])
        diff_ws.cell(i, 2, row["param"])
        diff_ws.cell(i, 3, row["robi"])
        diff_ws.cell(i, 4, row["gp"])
        diff_ws.cell(i, 5, f"{row['rs']:.0%} of {row['nr']}")
        diff_ws.cell(i, 6, f"{row['gs']:.0%} of {row['ng']}")
        diff_ws.cell(i, 7, row["why"])
        paint_row(diff_ws, i, len(diff_headers), amber if i <= 6 else None)
        diff_ws.row_dimensions[i].height = 48
    for col, width in zip("ABCDEFG", (42, 28, 22, 22, 16, 16, 78)):
        diff_ws.column_dimensions[col].width = width

    same_ws = wb.create_sheet("Shared settings")
    same_headers = ["Parameter", "Setting on both", "Robi samples", "GP samples", "Seen in"]
    paint_header(same_ws, same_headers)
    for i, row in enumerate(shared, 2):
        same_ws.cell(i, 1, row["param"])
        same_ws.cell(i, 2, row["robi"])
        same_ws.cell(i, 3, row["nr"])
        same_ws.cell(i, 4, row["ng"])
        same_ws.cell(i, 5, row["group"])
        paint_row(same_ws, i, len(same_headers))
    for col, width in zip("ABCDE", (34, 36, 16, 16, 42)):
        same_ws.column_dimensions[col].width = width

    mix_ws = wb.create_sheet("Varies by cell")
    mix_headers = ["Parameter", "Robi most common", "Robi share", "GP most common", "GP share", "Seen in"]
    paint_header(mix_ws, mix_headers)
    for i, row in enumerate(mixed, 2):
        mix_ws.cell(i, 1, row["param"])
        mix_ws.cell(i, 2, row["robi"])
        mix_ws.cell(i, 3, f"{row['rs']:.0%} of {row['nr']}")
        mix_ws.cell(i, 4, row["gp"])
        mix_ws.cell(i, 5, f"{row['gs']:.0%} of {row['ng']}")
        mix_ws.cell(i, 6, row["group"])
        paint_row(mix_ws, i, len(mix_headers))
    for col, width in zip("ABCDEF", (32, 28, 16, 28, 16, 40)):
        mix_ws.column_dimensions[col].width = width

    # Compact event picture: top signature per event per operator.
    ev = wb.create_sheet("Handover events")
    ev_headers = ["Event", "Operator", "Share of that event", "Count", "Profile"]
    paint_header(ev, ev_headers)
    er = 2
    for event in ("eventA3", "eventA2", "eventA1", "eventA4", "eventA5", "eventA6_r10"):
        if event not in event_profiles:
            continue
        for op in OPS:
            counter = event_profiles[event][op]
            if not counter:
                continue
            total = sum(counter.values())
            sig, cnt = counter.most_common(1)[0]
            pretty = sig
            pretty = re.sub(
                r"threshold_RSRP=(-?\d+)",
                lambda m: f"RSRP {int(m.group(1)) - 140} dBm",
                pretty,
            )
            pretty = re.sub(
                r"a3_Offset=(-?\d+)",
                lambda m: f"A3 offset {int(m.group(1)) * 0.5:g} dB",
                pretty,
            )
            pretty = re.sub(
                r"hysteresis=(-?\d+)",
                lambda m: f"hysteresis {int(m.group(1)) * 0.5:g} dB",
                pretty,
            )
            ev.cell(er, 1, event)
            ev.cell(er, 2, "GP" if op == "Grameenphone" else op)
            ev.cell(er, 3, f"{cnt / total:.0%}")
            ev.cell(er, 4, cnt)
            ev.cell(er, 5, pretty)
            paint_row(ev, er, len(ev_headers))
            er += 1
    for col, width in zip("ABCDE", (16, 14, 22, 12, 88)):
        ev.column_dimensions[col].width = width

    # Appendix: every kept leaf, for anyone who wants the raw inventory.
    appendix = wb.create_sheet("Appendix")
    app_headers = ["Section", "Group", "Parameter", "Robi", "Grameenphone", "Status", "Robi n", "GP n"]
    paint_header(appendix, app_headers)
    raw_rows = []
    for (section, group, param), by_op in values.items():
        robi, gp = by_op["Robi"], by_op["Grameenphone"]
        if not robi and not gp:
            continue
        raw_rows.append((section, group, param, robi, gp))
    raw_rows.sort(key=lambda item: (item[0], item[1], item[2]))
    for i, (section, group, param, robi, gp) in enumerate(raw_rows, 2):
        appendix.cell(i, 1, section)
        appendix.cell(i, 2, group)
        appendix.cell(i, 3, param)
        appendix.cell(i, 4, top_interpreted(param, robi, 3))
        appendix.cell(i, 5, top_interpreted(param, gp, 3))
        appendix.cell(i, 6, compare_status_param(param, robi, gp))
        appendix.cell(i, 7, sum(robi.values()))
        appendix.cell(i, 8, sum(gp.values()))
        paint_row(appendix, i, len(app_headers))
    for col, width in zip("ABCDEFGH", (14, 36, 28, 36, 36, 18, 12, 12)):
        appendix.column_dimensions[col].width = width

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    md_path = out_path.with_suffix(".md")
    _write_md(md_path, differ, shared, mixed, msg_counts)
    return out_path, len(differ), len(shared), len(mixed)


def _comparison_rows(values):
    rows = []
    for (section, group, param), by_op in values.items():
        robi, gp = by_op["Robi"], by_op["Grameenphone"]
        if not robi and not gp:
            continue
        published = publish_section(section, group, param)
        if param in _FRAME_GROUP:
            if section == "Mobility" or "mobility" in group.lower():
                where = "handover"
            elif "SCell" in group or "Carrier" in group:
                where = "SCell"
            elif section == "Idle mode" or "sib" in group.lower():
                where = "SIB"
            else:
                where = "dedicated"
            group = f"{_FRAME_GROUP[param]} — {where}"
        status = compare_status_param(param, robi, gp)
        note = _WHY.get(param, "") if status == "Different" else ""
        rows.append(
            {
                "section": published,
                "group": group,
                "param": param,
                "robi": top_merged(param, robi),
                "gp": top_merged(param, gp),
                "status": status,
                "nr": sum(robi.values()),
                "ng": sum(gp.values()),
                "note": note,
            }
        )

    def sort_key(row):
        try:
            si = COMPARE_SECTION_ORDER.index(row["section"])
        except ValueError:
            si = 99
        return (si, row["group"], row["param"])

    rows.sort(key=sort_key)
    return rows


_COMPARE_HEADERS = [
    "Section",
    "Group",
    "Parameter",
    "Robi value",
    "GP value",
    "Status",
    "Robi samples",
    "GP samples",
    "Note",
]


def _paint_header(ws, headers):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    for c, header in enumerate(headers, 1):
        cell = ws.cell(1, c, header)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    ws.row_dimensions[1].height = 30
    ws.sheet_view.showGridLines = False


def _paint_compare_body(ws, start_row, end_row):
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    status_fill = {
        "Different": PatternFill("solid", fgColor="FCE4D6"),
        "Same (dominant)": PatternFill("solid", fgColor="E2EFDA"),
        "Robi only": PatternFill("solid", fgColor="DDEBF7"),
        "GP only": PatternFill("solid", fgColor="FFF2CC"),
    }
    for r in range(start_row, end_row + 1):
        status = ws.cell(r, 6).value
        for c in range(1, len(_COMPARE_HEADERS) + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
        if status in status_fill:
            ws.cell(r, 6).fill = status_fill[status]
            ws.cell(r, 6).font = Font(name="Calibri", size=10, bold=True)


def _write_compare_sheet(ws, rows):
    _paint_header(ws, _COMPARE_HEADERS)
    for i, row in enumerate(rows, 2):
        ws.cell(i, 1, row["section"])
        ws.cell(i, 2, row["group"])
        ws.cell(i, 3, row["param"])
        ws.cell(i, 4, row["robi"])
        ws.cell(i, 5, row["gp"])
        ws.cell(i, 6, row["status"])
        ws.cell(i, 7, row["nr"])
        ws.cell(i, 8, row["ng"])
        ws.cell(i, 9, row["note"])
    if rows:
        _paint_compare_body(ws, 2, 1 + len(rows))
    widths = [22, 36, 28, 42, 42, 16, 14, 14, 55]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width


def _ms_of_day(stamp: str) -> float:
    hour, minute, second = stamp.split(":")
    return (int(hour) * 3600 + int(minute) * 60 + float(second)) * 1000


def _percentile(values: list[float], pct: float) -> float:
    ordered = sorted(values)
    index = min(len(ordered) - 1, max(0, int(round((pct / 100) * (len(ordered) - 1)))))
    return ordered[index]


def volte_calls(l3_dir: Path) -> list[dict]:
    """On-net mobile-originated VoLTE calls. CST is INVITE to 180 Ringing."""
    sip = defaultdict(list)
    radio = defaultdict(list)
    for path in sorted(l3_dir.glob("*.txt")):
        for line in path.open(errors="replace"):
            match = HDR.match(line)
            if not match:
                continue
            stamp, direction, layer, name, _tech, side, sid, operator = match.groups()
            when = _ms_of_day(stamp)
            if layer == "SIP":
                sip[(path.name, operator, sid)].append((when, direction, side, name.strip()))
            elif layer in {"LTE-RRC", "LTE-NAS"}:
                radio[(path.name, operator, side)].append((when, name.strip()))

    def first(events, pred):
        for event in events:
            if pred(event):
                return event
        return None

    def first_radio(file_name, operator, side, start, text, low, high):
        hits = [
            when - start
            for when, message in radio[(file_name, operator, side)]
            if text in message and low <= when - start <= high
        ]
        return min(hits) if hits else None

    calls = []
    for (file_name, operator, _sid), events in sip.items():
        events = sorted(events)
        invite = first(events, lambda event: event[1] == "U" and event[2] == "A" and "INVITE: Request" in event[3])
        if invite is None:
            continue
        start = invite[0]

        def after(pred):
            event = first(events, pred)
            return None if event is None else event[0] - start

        ringing = after(lambda event: event[1] == "D" and event[2] == "A" and "INVITE: Ringing" in event[3])
        progress = after(lambda event: event[1] == "D" and event[2] == "A" and "Session Progress" in event[3])
        trying = after(lambda event: event[1] == "D" and event[2] == "A" and "INVITE: Trying" in event[3])
        answered = after(lambda event: event[1] == "D" and event[2] == "A" and "INVITE: OK" in event[3])
        if ringing is not None:
            outcome = "Alerted"
        elif answered is not None:
            outcome = "Answered without 180"
        else:
            outcome = "No alert"
        paging = None
        rrc = first_radio(file_name, operator, "B", start, "RRCConnectionSetup", 0, 8000)
        if rrc is not None:
            pages = [
                when - start
                for when, message in radio[(file_name, operator, "B")]
                if "Paging" in message and rrc - 3000 <= when - start <= rrc
            ]
            paging = max(pages) if pages else None
        calls.append(
            {
                "operator": operator,
                "file": file_name,
                "outcome": outcome,
                "cst": ringing,
                "progress": progress,
                "trying": trying,
                "answered": answered,
                "mo_qci": first_radio(file_name, operator, "A", start, "dedicated EPS bearer context request", -200, 8000),
                "mt_qci": first_radio(file_name, operator, "B", start, "dedicated EPS bearer context request", -200, 8000),
                "mt_page": paging,
                "mt_rrc": rrc,
            }
        )
    calls.sort(key=lambda row: ({"Robi": 0, "Grameenphone": 1, "Banglalink": 2}.get(row["operator"], 9), row["file"], row["cst"] or 99999))
    return calls


def _cst_stat(calls: list[dict], operator: str, field: str) -> dict | None:
    values = [row[field] for row in calls if row["operator"] == operator and row[field] is not None and row["outcome"] == "Alerted"]
    if not values:
        return None
    return {
        "n": len(values),
        "median": statistics.median(values),
        "mean": sum(values) / len(values),
        "p90": _percentile(values, 90),
    }


def _fmt_ms(stat: dict | None) -> str:
    if not stat:
        return "—"
    return f"{stat['median']:.0f} ms   (mean {stat['mean']:.0f}, p90 {stat['p90']:.0f}, n={stat['n']})"


def _alerted_median(calls: list[dict], operator: str, getter) -> tuple[float | None, int]:
    values = []
    for row in calls:
        if row["operator"] != operator or row["outcome"] != "Alerted":
            continue
        value = getter(row)
        if value is not None:
            values.append(value)
    if not values:
        return None, 0
    return statistics.median(values), len(values)


def _add_cst_segment_sheet(wb, calls: list[dict]) -> None:
    """One row per setup step. Times are medians, in seconds."""
    operators = ("Robi", "Grameenphone", "Banglalink")

    def delta(later, earlier):
        def getter(row):
            end = row.get(later)
            start = 0 if earlier is None else row.get(earlier)
            if end is None or start is None:
                return None
            return (end - start) / 1000
        return getter

    steps = [
        ("1. INVITE → 100 Trying", "Phone to P-CSCF", delta("trying", None)),
        ("2. 100 Trying → caller voice bearer", "P-CSCF to PCRF to gateway to MME", delta("mo_qci", "trying")),
        ("3. Caller voice bearer → 183", "Session progress after the bearer", delta("progress", "mo_qci")),
        ("4. 183 → 180 Ringing", "Called phone starts ringing", delta("cst", "progress")),
        ("Call setup time", "INVITE → 180 Ringing", delta("cst", None)),
        ("Called voice bearer", "INVITE → called UE QCI 1", delta("mt_qci", None)),
        ("INVITE → paging of called UE", "IMS to the page", delta("mt_page", None)),
        ("Paging → called RRC", "Air interface", delta("mt_rrc", "mt_page")),
        ("Called RRC → called voice bearer", "Bearer after the phone connects", delta("mt_qci", "mt_rrc")),
    ]
    sheet = wb.create_sheet("CST segments", 3)
    headers = ["Segment", "Domain", "Robi", "Grameenphone", "Banglalink", "Robi extra vs BL"]
    _paint_header(sheet, headers)
    for index, (label, domain, getter) in enumerate(steps, 2):
        sheet.cell(index, 1, label)
        sheet.cell(index, 2, domain)
        medians = {}
        counts = {}
        for column, operator in enumerate(operators, 3):
            median, count = _alerted_median(calls, operator, getter)
            medians[operator] = median
            counts[operator] = count
            if median is None or count < 5:
                text = "already connected" if "paging" in label.lower() or "RRC" in label or "after the phone" in domain else "—"
                if count >= 5:
                    text = f"{median:.2f} s"
                elif count == 0:
                    text = "—"
                else:
                    text = "already connected"
            else:
                text = f"{median:.2f} s"
            sheet.cell(index, column, text)
        robi = medians["Robi"]
        bl = medians["Banglalink"]
        if robi is not None and bl is not None and counts["Robi"] >= 5 and counts["Banglalink"] >= 5:
            gap = robi - bl
            sheet.cell(index, 6, f"{gap:+.2f} s")
        else:
            sheet.cell(index, 6, "—")
    note = 2 + len(steps) + 1
    sheet.cell(
        note,
        1,
        "Medians for calls that rang. The first four segments are the caller path and add up to the call setup time. "
        "Banglalink's called phone is still connected, so the paging rows are not a Banglalink step. "
        "Segment 2 is the Robi gap: the originating voice bearer. "
        "The called voice bearer is the second gap. Fixing only segment 2 leaves setup waiting on the called phone, near Grameenphone rather than Banglalink.",
    )
    sheet.merge_cells(start_row=note, start_column=1, end_row=note + 1, end_column=6)
    sheet.cell(note, 1).alignment = Alignment(wrap_text=True, vertical="top")
    sheet.row_dimensions[note].height = 48
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    for row in range(2, 2 + len(steps)):
        for column in range(1, 7):
            cell = sheet.cell(row, column)
            cell.font = Font(name="Calibri", size=10, bold=(row == 6))
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = thin
        if row in (3, 7):
            sheet.cell(row, 6).fill = PatternFill("solid", fgColor="FCE4D6")
    for column, width in enumerate((42, 42, 22, 22, 24, 20), 1):
        sheet.column_dimensions[get_column_letter(column)].width = width


def _add_volte_sheets(wb, calls: list[dict]) -> None:
    _add_cst_segment_sheet(wb, calls)
    operators = ("Robi", "Grameenphone", "Banglalink")
    labels = {"Robi": "Robi", "Grameenphone": "Grameenphone", "Banglalink": "Banglalink"}
    metrics = [
        ("cst", "Call setup time: INVITE → 180 Ringing"),
        ("progress", "INVITE → 183 Session Progress"),
        ("trying", "INVITE → 100 Trying"),
        ("answered", "INVITE → 200 OK"),
        ("mo_qci", "Caller dedicated bearer (QCI 1)"),
        ("mt_qci", "Called-UE dedicated bearer (QCI 1)"),
        ("mt_page", "INVITE → paging of the called UE"),
        ("mt_rrc", "INVITE → called UE RRC setup"),
    ]
    summary = wb.create_sheet("VoLTE CST", 4)
    headers = ["Step", "Robi", "Grameenphone", "Banglalink"]
    _paint_header(summary, headers)
    for index, (field, label) in enumerate(metrics, 2):
        summary.cell(index, 1, label)
        for column, operator in enumerate(operators, 2):
            summary.cell(index, column, _fmt_ms(_cst_stat(calls, operator, field)))
    note_row = 2 + len(metrics) + 1
    summary.cell(
        note_row,
        1,
        "On-net calls only: each UE calls its own network. "
        "CST is the time from the caller's SIP INVITE to 180 Ringing. "
        "Times are the median, with the mean, 90th percentile, and number of alerted calls. "
        "A call that never rang is left out of the median.",
    )
    summary.merge_cells(start_row=note_row, start_column=1, end_row=note_row + 1, end_column=4)
    summary.cell(note_row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    summary.row_dimensions[note_row].height = 36
    for column, width in enumerate((52, 42, 42, 42), 1):
        summary.column_dimensions[get_column_letter(column)].width = width
    if len(metrics):
        thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )
        for row in range(2, 2 + len(metrics)):
            for column in range(1, 5):
                cell = summary.cell(row, column)
                cell.font = Font(name="Calibri", size=10, bold=(column == 1 or row == 2))
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                cell.border = thin
            if row == 2:
                for column in range(1, 5):
                    summary.cell(row, column).fill = PatternFill("solid", fgColor="FCE4D6")

    attempts = Counter(row["operator"] for row in calls)
    alerted = Counter(row["operator"] for row in calls if row["outcome"] == "Alerted")
    detail = wb.create_sheet("VoLTE calls", 5)
    detail_headers = [
        "Operator",
        "File",
        "Outcome",
        "CST ms",
        "INVITE to 183 ms",
        "INVITE to 100 ms",
        "INVITE to 200 ms",
        "Caller QCI 1 ms",
        "Called QCI 1 ms",
        "Called paging ms",
        "Called RRC ms",
    ]
    _paint_header(detail, detail_headers)
    for index, row in enumerate(calls, 2):
        values = [
            labels[row["operator"]],
            row["file"],
            row["outcome"],
            None if row["cst"] is None else round(row["cst"]),
            None if row["progress"] is None else round(row["progress"]),
            None if row["trying"] is None else round(row["trying"]),
            None if row["answered"] is None else round(row["answered"]),
            None if row["mo_qci"] is None else round(row["mo_qci"]),
            None if row["mt_qci"] is None else round(row["mt_qci"]),
            None if row["mt_page"] is None else round(row["mt_page"]),
            None if row["mt_rrc"] is None else round(row["mt_rrc"]),
        ]
        for column, value in enumerate(values, 1):
            detail.cell(index, column, value)
    if calls:
        thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )
        for row in range(2, 2 + len(calls)):
            for column in range(1, len(detail_headers) + 1):
                cell = detail.cell(row, column)
                cell.font = Font(name="Calibri", size=10)
                cell.alignment = Alignment(vertical="center")
                cell.border = thin
    for column, width in enumerate((16, 12, 24, 12, 18, 18, 18, 18, 18, 18, 16), 1):
        detail.column_dimensions[get_column_letter(column)].width = width

    actions = wb.create_sheet("VoLTE actions", 6)
    _paint_header(actions, ["Priority", "What the log shows", "Action for Robi"])
    action_rows = [
        (
            "1. Caller voice bearer",
            f"Robi median call setup is {_fmt_ms(_cst_stat(calls, 'Robi', 'cst'))}. "
            f"Grameenphone is {_fmt_ms(_cst_stat(calls, 'Grameenphone', 'cst'))}. "
            f"Banglalink is {_fmt_ms(_cst_stat(calls, 'Banglalink', 'cst'))}. "
            "The caller is already RRC-connected before the INVITE. "
            f"Robi's QCI 1 bearer on that UE then takes {_fmt_ms(_cst_stat(calls, 'Robi', 'mo_qci'))}. "
            f"Grameenphone does the same step in {_fmt_ms(_cst_stat(calls, 'Grameenphone', 'mo_qci'))} "
            f"and Banglalink in {_fmt_ms(_cst_stat(calls, 'Banglalink', 'mo_qci'))}. "
            "183 follows the Robi caller bearer by about 160 ms, so setup waits on this bearer.",
            "Bring the originating dedicated bearer under 200 ms. "
            "The wait is after the UE is connected, so look at the originating Rx/Gx path: "
            "P-CSCF to PCRF, PCRF to the packet gateway, and the create-bearer toward the calling UE.",
        ),
        (
            "2. Paging of the called UE",
            "The called Robi UE has been RRC-released about 19 s earlier, so it is idle. "
            f"Paging is seen {_fmt_ms(_cst_stat(calls, 'Robi', 'mt_page'))} after INVITE. "
            "From that page, RRC setup is about 80 ms and QCI 1 about 360 ms later. "
            "Grameenphone has the same shape. "
            f"Banglalink's called UE is still RRC-connected, and its called bearer is {_fmt_ms(_cst_stat(calls, 'Banglalink', 'mt_qci'))}.",
            "Shorten the time from INVITE to the paging message for an idle called UE. "
            "The air interface after the page is already fast. "
            "Fixing only the caller bearer leaves this terminating wait, and setup would then look like Grameenphone rather than Banglalink.",
        ),
        (
            "Leave these",
            "183 to 180 Ringing is about 300 ms on all three networks. "
            "180 Ringing to 200 OK is about 1.8 s on all three, which is the answer delay. "
            f"Robi alerted {alerted['Robi']} of {attempts['Robi']} calls. "
            f"Grameenphone alerted {alerted['Grameenphone']} of {attempts['Grameenphone']}. "
            f"Banglalink alerted {alerted['Banglalink']} of {attempts['Banglalink']}, and the remaining Banglalink call was answered without a 180.",
            "Do not tune ringing or the answer delay for this CST gap. Setup success is already high.",
        ),
    ]
    for index, (priority, evidence, action) in enumerate(action_rows, 2):
        actions.cell(index, 1, priority)
        actions.cell(index, 2, evidence)
        actions.cell(index, 3, action)
        actions.row_dimensions[index].height = 90
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    for row in range(2, 5):
        for column in range(1, 4):
            cell = actions.cell(row, column)
            cell.font = Font(name="Calibri", size=10, bold=(column == 1))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
        actions.cell(row, 1).fill = PatternFill("solid", fgColor="FCE4D6" if row < 4 else "E2EFDA")
    for column, width in enumerate((28, 88, 78), 1):
        actions.column_dimensions[get_column_letter(column)].width = width


def _sheet_name(prefix: str, section: str) -> str:
    name = f"{prefix}{section}" if prefix else section
    return name[:31]


def build_v3_workbook(values, msg_counts, out_path: Path, calls: list[dict] | None = None):
    """v3: full parameter list (v1) plus engineering meaning (v2), named sections."""
    rows = _comparison_rows(values)
    wb = Workbook()

    cover = wb.active
    cover.title = "Versions"
    cover.sheet_view.showGridLines = False
    cover["A1"] = "L3 REPORT FOR BOGRA"
    cover["A1"].font = Font(name="Calibri", size=20, bold=True, color="1F4E78")
    cover["A2"] = "Version 8 — VoLTE setup split into segments against Banglalink"
    cover["A2"].font = Font(name="Calibri", size=13, color="595959")
    cover["A4"] = "Version"
    cover["B4"] = "What it contains"
    cover["C4"] = "File"
    for col in range(1, 4):
        cover.cell(4, col).font = Font(name="Calibri", bold=True, color="FFFFFF")
        cover.cell(4, col).fill = PatternFill("solid", fgColor="1F4E78")
    versions = [
        ("v1", "Every decoded parameter. Sections: Mobility, Power, Idle mode, Others.", "L3_REPORT_FOR_BOGRA_v1.xlsx"),
        ("v2", "Only settings that are stable on both networks, with the idle-search finding.", "L3_REPORT_FOR_BOGRA_v2.xlsx"),
        ("v3", "Full list, named sections, one sheet per section, and a difference sheet per section.", "L3_REPORT_FOR_BOGRA_v3.xlsx"),
        ("v4", "v3, with subframeAssignment, specialSubframePatterns, cyclic prefix, and TDD ACK/NACK moved to Frame structure.", "L3_REPORT_FOR_BOGRA_v4.xlsx"),
        ("v5", "v4, with the mapped meaning written in the Robi value and GP value cells. The separate meaning columns are gone.", "L3_REPORT_FOR_BOGRA_v5.xlsx"),
        ("v6", "v5, with Power as the first section on Comparison and Differences.", "L3_REPORT_FOR_BOGRA_v6.xlsx"),
        ("v7", "v6, plus on-net VoLTE call setup time for Robi, Grameenphone, and Banglalink, and the two bearer actions for Robi.", "L3_REPORT_FOR_BOGRA_v7.xlsx"),
        ("v8 (this file)", "v7, with call setup split into segments. The slow Robi step is the caller voice bearer.", "L3_REPORT_FOR_BOGRA_v8.xlsx"),
    ]
    for i, (ver, what, fname) in enumerate(versions, 5):
        cover.cell(i, 1, ver)
        cover.cell(i, 2, what)
        cover.cell(i, 3, fname)
        for col in range(1, 4):
            cover.cell(i, col).alignment = Alignment(wrap_text=True, vertical="top")
            cover.cell(i, col).font = Font(name="Calibri", size=11)
        cover.row_dimensions[i].height = 36
    cover["A13"] = (
        f"Source: Layer 3_Active Mode_Data_Robi_v1.7z. "
        f"Robi {sum(msg_counts['Robi'].values()):,} messages, "
        f"Grameenphone {sum(msg_counts['Grameenphone'].values()):,}. "
        "Each Robi value and GP value line is the setting and its mapped unit when one exists "
        "(RSRP index to dBm, a3-Offset and hysteresis to dB, q-RxLevMin to dBm), then the sample count and share. "
        "Frame structure holds subframeAssignment, specialSubframePatterns, tdd-SpecialSubframe-r11, "
        "ul-CyclicPrefixLength, and tdd-AckNackFeedbackMode, including the copies broadcast in SIB, "
        "sent at handover, and configured on an SCell."
    )
    cover["A13"].alignment = Alignment(wrap_text=True, vertical="top")
    cover.merge_cells("A13:C15")
    cover.column_dimensions["A"].width = 18
    cover.column_dimensions["B"].width = 88
    cover.column_dimensions["C"].width = 36

    comparison = wb.create_sheet("Comparison", 1)
    _write_compare_sheet(comparison, rows)

    diff_rows = [r for r in rows if r["status"] in {"Different", "Robi only", "GP only"}]
    differences = wb.create_sheet("Differences", 2)
    _write_compare_sheet(differences, diff_rows)
    if calls:
        _add_volte_sheets(wb, calls)

    used = {"Versions", "Comparison", "Differences"}
    for section in SECTION_ORDER:
        section_rows = [r for r in rows if r["section"] == section]
        if not section_rows:
            continue
        name = _sheet_name("", section)
        if name in used:
            name = name[:28] + " 2"
        used.add(name)
        _write_compare_sheet(wb.create_sheet(name), section_rows)

        section_diff = [r for r in section_rows if r["status"] in {"Different", "Robi only", "GP only"}]
        if not section_diff:
            continue
        dname = _sheet_name("Diff ", section)
        if dname in used:
            dname = ("D " + section)[:31]
        used.add(dname)
        _write_compare_sheet(wb.create_sheet(dname), section_diff)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    lines = [
        "# L3 REPORT FOR BOGRA — v8",
        "",
        "Robi vs Grameenphone parameter comparison, plus on-net VoLTE call setup time for Robi, Grameenphone, and Banglalink.",
        "Call setup time is the caller's SIP INVITE to 180 Ringing. Power stays first on the Comparison sheet.",
        "",
        "## Versions",
        "",
        "| Version | File | Contents |",
        "| --- | --- | --- |",
        "| v1 | `output/L3_REPORT_FOR_BOGRA_v1.xlsx` | Full list. Sections: Mobility, Power, Idle mode, Others. |",
        "| v2 | `output/L3_REPORT_FOR_BOGRA_v2.xlsx` | Stable settings only, with the idle-search finding. |",
        "| v3 | `output/L3_REPORT_FOR_BOGRA_v3.xlsx` | Full list, named sections, one sheet per section, difference sheets. |",
        "| v4 | `output/L3_REPORT_FOR_BOGRA_v4.xlsx` | Frame structure section for the TDD subframe parameters. |",
        "| v5 | `output/L3_REPORT_FOR_BOGRA_v5.xlsx` | Mapped meaning inside the Robi and GP value cells. |",
        "| v6 | `output/L3_REPORT_FOR_BOGRA_v6.xlsx` | Power first on Comparison and Differences. |",
        "| v7 | `output/L3_REPORT_FOR_BOGRA_v7.xlsx` | VoLTE call setup time and the Robi actions. |",
        "| v8 | `output/L3_REPORT_FOR_BOGRA_v8.xlsx` | Setup time split into segments against Banglalink. Current. |",
        "",
        f"Parameters in v8: {len(rows)}. Rows that differ: {len(diff_rows)}.",
        "",
        "## Sections",
        "",
        "| Section | Parameters | Different | Robi only | GP only | Same |",
        "| --- | ---: | ---: | ---: | ---: | ---: |",
    ]
    for section in COMPARE_SECTION_ORDER:
        part = [r for r in rows if r["section"] == section]
        if not part:
            continue
        lines.append(
            "| {sec} | {n} | {d} | {ro} | {go} | {s} |".format(
                sec=section,
                n=len(part),
                d=sum(r["status"] == "Different" for r in part),
                ro=sum(r["status"] == "Robi only" for r in part),
                go=sum(r["status"] == "GP only" for r in part),
                s=sum(r["status"] == "Same (dominant)" for r in part),
            )
        )
    lines += [
        "",
        "Open Comparison for the full list. Open VoLTE CST for setup time and VoLTE actions for what to change on Robi. Open Differences for every row that is not the same dominant value.",
        "",
    ]
    md_path = out_path.with_suffix(".md")
    md_path.write_text("\n".join(lines))
    return out_path, len(rows), len(diff_rows), wb.sheetnames


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--l3-dir", type=Path, default=DEFAULT_L3_DIR)
    ap.add_argument("--out", type=Path, default=ROOT / "output" / "L3_REPORT_FOR_BOGRA_v8.xlsx")
    args = ap.parse_args()
    print(f"Scanning {args.l3_dir} ...")
    values, msg_counts, _events = scan(args.l3_dir)
    calls = volte_calls(args.l3_dir)
    out, n_rows, n_diff, sheets = build_v3_workbook(values, msg_counts, args.out, calls)
    latest = ROOT / "output" / "L3_REPORT_FOR_BOGRA.xlsx"
    latest.write_bytes(out.read_bytes())
    latest_md = ROOT / "output" / "L3_REPORT_FOR_BOGRA.md"
    latest_md.write_text(out.with_suffix(".md").read_text())
    print(f"Wrote {out}")
    print(f"Copied current report to {latest}")
    print(f"  parameters: {n_rows}")
    print(f"  differences: {n_diff}")
    print(f"  sheets: {', '.join(sheets)}")


if __name__ == "__main__":
    main()
