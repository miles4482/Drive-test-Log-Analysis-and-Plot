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

SECTION_ORDER = ("Mobility", "Power", "Idle mode", "Others")

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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--l3-dir", type=Path, default=DEFAULT_L3_DIR)
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = ap.parse_args()
    print(f"Scanning {args.l3_dir} ...")
    values, msg_counts, event_profiles = scan(args.l3_dir)
    out, n, sec_stat = build_workbook(values, msg_counts, event_profiles, args.out)
    print(f"Wrote {out}  ({n} parameter rows)")
    for s in SECTION_ORDER:
        print(f"  {s}: {dict(sec_stat[s])}")


if __name__ == "__main__":
    main()
