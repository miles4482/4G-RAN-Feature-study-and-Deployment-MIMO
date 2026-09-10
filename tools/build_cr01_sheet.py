#!/usr/bin/env python3
"""Build MIMO_Deployment_v4.0.xlsx from v3.0.

- Last sheet: 16. CR01 MIMO Exec Pack  (10-site CME change request)
- Every sheet: append section "Performance and Monitoring Counter"

CR01 source (this repo): CR01_MIMO Performance improvemnet_Change Request.xlsx
taken from 5G-RAN-BASIC-TO-ADVANCE_OPTIMIZATION NSA report v2.1 (DHK 9 Sep 2026).
Huawei CME planned-value export: Proposed columns list only the bits being turned ON
(other live bits stay). 10 macros × 3 sectors = 30 cells. 32T only.
"""
import os, shutil, zipfile, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Font
from mimo_excel_style import *
from mimo_counters import add_perf_monitor, append_counters_to_all_sheets, SECTION_TITLE

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "MIMO_Deployment_v3.0.xlsx")
OUT = os.path.join(ROOT, "MIMO_Deployment_v4.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v4.0.zip")
CR01_NAME = "CR01_MIMO Performance improvemnet_Change Request.xlsx"

SHEET_NAME = "16. CR01 MIMO Exec Pack"  # 24 chars
INCON = "14. MIMO Incon Report"
SUGG = "15. MIMO Suggestions from Doc"
COLS = 11
WIDTHS = [6, 12, 18, 10, 62, 24, 34, 40, 36, 44, 32]

BLUE_HDR = "5B9BD5"
YELLOW_HDR = "FFC000"
LINK_BLUE = "0563C1"

SEV_FILL = {
    "Enable": "F8CBAD", "Landed": "C6EFCE", "Missing": "F8CBAD",
    "Partial": "FFF2CC", "Hold": "FCE4D6", "Skip": "E2EFDA",
    "Seq": "1F4E79", "Info": "E2D5F1", "Check": "DDEBF7",
}

GNBS = [
    ("DHGUL66", {101: 5, 102: 8, 103: 8}),
    ("DHGULV8", {101: 2, 102: 2, 103: 5}),
    ("DHBDD13", {101: 9, 102: 8, 103: 8}),
    ("DHGULO9", {101: 9, 102: 8, 103: 8}),
    ("DHGUL40", {101: 6, 102: 5, 103: 8}),
    ("DHBDD34", {101: 4, 102: 4, 103: 6}),
    ("DHGULQ1", {101: 8, 102: 7, 103: 7}),
    ("DHGUL05", {101: 7, 102: 4, 103: 4}),
    ("DHBDD69", {101: 8, 102: 4, 103: 6}),
    ("DHGULU2", {101: 7, 102: 9, 103: 7}),
]
CELLS = (101, 102, 103)


def href_sheet(cell, sheet, text=None):
    cell.value = text or sheet
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A1",
                               display=str(text or sheet))
    cell.font = Font(name="Calibri", size=9, color=LINK_BLUE, underline="single", bold=True)
    cell.alignment = align("left", "center", True)


def put11(ws, r, values, fills=None, bolds=None, height=None, center=None, font_color=None):
    for i, v in enumerate(values, 1):
        fh = fills[i - 1] if fills and i - 1 < len(fills) else None
        b = bolds[i - 1] if bolds and i - 1 < len(bolds) else False
        h = "center" if center and i in center else "left"
        col = font_color if font_color else "000000"
        put(ws, r, i, v, size=9, bold=b, color=col, fill_hex=fh, h=h, v="top", border=True)
    if height is None:
        longest = max((len(str(v)) if v is not None else 0) for v in values)
        lines = max((str(v).count("\n") + 1) for v in values)
        height = min(88, max(22, 16 + longest // 90 * 10, 16 * lines))
    ws.row_dimensions[r].height = height
    return r + 1


def mml_header(ws, r):
    titles = ["SN", "RAT", "Doc Name", "Action", "MML Command (Proposed)",
              "MO Name", "Parameter ID", "Live Value (CR01 / dump)",
              "proposed Parameter Value", "Purpose/Short Notes", "License"]
    for i, t in enumerate(titles, 1):
        fh = BLUE_HDR if i <= 4 else YELLOW_HDR
        put(ws, r, i, t, size=9, bold=True, color="000000", fill_hex=fh,
            h="center", v="center", border=True)
    ws.row_dimensions[r].height = 28
    return r + 1


def mml_row(ws, r, rec):
    action = rec[3]
    fh = SEV_FILL.get(action, WHITE)
    if action == "Seq":
        return put11(ws, r, rec, fills=[NAVY] * 11, bolds=[True] * 11,
                     font_color=WHITE, height=24)
    fills = [PALE_BLUE] * 4 + [fh] + [PALE_GOLD] * 6
    fills[3] = fh
    return put11(ws, r, rec, fills=fills,
                 bolds=[False, False, False, True, True] + [False] * 6,
                 center={1, 2, 4})


# One Proposed column from CR01 = one MML (bits listed are OR-ed onto the live bitfield).
def template_mml(pid="{NrDuCellId}"):
    return [
        (1, "TDD", "CR01 seq", "Seq", "", "", "", "", "",
         "Safer send order inside the one CME file. CR01 still lands as one planned-value export. "
         "Do not add SRS_IC this night (tight MUX is ON).", ""),
        (2, "TDD", "MIMO TDD W1", "Enable",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1;",
         "NRDUCellBeamAlgo", "WeightAlgoSwitch",
         "SRS_WEIGHT_ESTIMATE_SW-0 (30/30)",
         "SRS_WEIGHT_ESTIMATE_SW-1",
         "Sheet14 MIMO-010. Largest basic-MIMO leftover. Perceive mode already DISTRIBUTED.",
         "FBFD-010003 (no extra license)"),
        (3, "TDD", "MIMO TDD W1", "Enable",
         f"MOD NRDUCELLULRANK: NrDuCellId={pid}, UlRankAlgoSw=UL_RANK_FAST_DECREASE_SW-1;",
         "NRDUCellUlRank", "UlRankAlgoSw",
         "UL_RANK_FAST_DECREASE_SW-0 (30/30)",
         "UL_RANK_FAST_DECREASE_SW-1",
         "Sheet14 MIMO-021. Fast UL rank drop on poor SRS.",
         "No extra license"),
        (4, "TDD", "MIMO TDD W1", "Enable",
         f"MOD NRDUCELLPUSCH: NrDuCellId={pid}, PuschPerformanceSwitch=PUSCH_CE_SINR_LEVEL_ENH_SW-1;",
         "NRDUCellPusch", "PuschPerformanceSwitch",
         "PUSCH_CE_SINR_LEVEL_ENH_SW-0 (30/30)",
         "PUSCH_CE_SINR_LEVEL_ENH_SW-1",
         "Sheet14 MIMO-021. PUSCH CE SINR-level enhance.",
         "No extra license"),
        (5, "TDD", "Beam Mgmt W1", "Enable",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, BeamOptAlgoSwitch=BEAM_TRACKING_SW-1&INTELLIGENT_BEAM_SELECTION_SW-1;",
         "NRDUCellBeamAlgo", "BeamOptAlgoSwitch",
         "BEAM_TRACKING-0 + INTELLIGENT_BEAM_SELECTION-0 (30/30)",
         "BEAM_TRACKING_SW-1&INTELLIGENT_BEAM_SELECTION_SW-1",
         "Sheet14 MIMO-011. Connected-mode BF tracking. CR01 unlabeled Proposed col O.",
         "FBFD-010015"),
        (6, "TDD", "Beam Mgmt W1", "Enable",
         f"MOD NRDUCELLALGOSWITCH: NrDuCellId={pid}, BeamOptSwitch=SSB_BEAM_ADAPT_SW-1&SSB_BEAM_VERTICAL_COV_IMP_SW-1;",
         "NRDUCellAlgoSwitch", "BeamOptSwitch",
         "SSB_BEAM_ADAPT-0 + VERTICAL_COV_IMP-0; DL_INITIAL_BEAM_SELECT already 1",
         "SSB_BEAM_ADAPT_SW-1&SSB_BEAM_VERTICAL_COV_IMP_SW-1",
         "Sheet14 MIMO-012. Keep other BeamOpt bits (initial-select) — CME Proposed lists only new 1s.",
         "FBFD-010015 / FOFD-010100"),
        (7, "TDD", "AHR Turbo leftover", "Enable",
         f"MOD NRDUCELLULPCCONFIG: NrDuCellId={pid}, UlPwrCtrlAlgoSwitch=SRS_JOINT_PC_SW-1;",
         "NRDUCellUlPcConfig", "UlPwrCtrlAlgoSwitch",
         "SRS_JOINT_PC off (Turbo master already ON)",
         "SRS_JOINT_PC_SW-1",
         "Sheet14 MIMO-018 PARTIAL — CR01 does NOT send SRS_IC_SW (correct vs tight MUX).",
         "FOFD-061201 master already ON"),
        (8, "TDD", "iBeam 1.0 W3", "Seq", "", "", "", "", "",
         "Main DL-interference pack. Master then children. License NR0S00BEAM00.", ""),
        (9, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, HighPrecisionBeamSwitch=ON;",
         "NRDUCellFeatureSw", "HighPrecisionBeamSwitch",
         "OFF (30/30). Phase2 stays OFF.",
         "ON",
         "Sheet14 MIMO-014. CR01 unlabeled Proposed col U. LST LICENSE NR0S00BEAM00 first.",
         "FOFD-081201 / NR0S00BEAM00 per cell"),
        (10, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLSRSMEAS: NrDuCellId={pid}, SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-1;",
         "NRDUCellSrsMeas", "SrsMeasOptSwitch",
         "SRS_BLIND_IS_MEAS_SW-0 (30/30)",
         "SRS_BLIND_IS_MEAS_SW-1",
         "Sheet14 MIMO-015. Blind IS for SRS measurement.",
         "FOFD-081201"),
        (11, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLSRS: NrDuCellId={pid}, SrsAlgoSwitch=SRS_SINR_MEAS_OPT_SW-1&SRS_TIGHT_MULTIPLEXING_SW-1;",
         "NRDUCellSrs", "SrsAlgoSwitch",
         "both 0 (30/30). USER_CHARACTER_SRS_ADAPT already 1",
         "SRS_SINR_MEAS_OPT_SW-1&SRS_TIGHT_MULTIPLEXING_SW-1",
         "Sheet14 MIMO-021 + MIMO-015. Do not stack SRS_IC tonight.",
         "FOFD-081201 + basic MIMO"),
        (12, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1;",
         "NRDUCellBeamAlgo", "ChannelOptAlgoSwitch",
         "BEAM_SELECT_OPT_SW-0 (30/30)",
         "BEAM_SELECT_OPT_SW-1",
         "Sheet14 MIMO-013. CR01 unlabeled Proposed col N.",
         "FOFD-081201"),
        (13, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-1&DL_MU_ANTI_INTRF_SCH_SW-1;",
         "NRDUCellDlMimo", "DLMuMimoSchSupplementSw",
         "both off; SMALL_PACKET_MERGE / SE_SU_MU_ADAPT already 1",
         "DL_MU_PRECISE_SCH_SW-1&DL_MU_ANTI_INTRF_SCH_SW-1",
         "Sheet14 MIMO-015. Precise + anti-interference MU scheduling.",
         "FOFD-081201"),
        (14, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLSCH: NrDuCellId={pid}, DlSchAlgoSwitch=DL_BWP_HYBRID_INTRF_RANDOM_SW-1&TAIL_PKT_MCS_OPT_SW-1&RES_BASED_DL_ADAPT_SCH_SW-1&DL_RLC_STAT_RPT_MERGE_SCH_SW-1;",
         "NRDUCellDlSch", "DlSchAlgoSwitch",
         "all four 0 (30/30). HEAVY_LOAD_SCH_PRI_OPT already 1",
         "DL_BWP_HYBRID_INTRF_RANDOM-1&TAIL_PKT_MCS_OPT-1&RES_BASED_DL_ADAPT_SCH-1&DL_RLC_STAT_RPT_MERGE_SCH-1",
         "Sheet14 MIMO-015/020 + extra DL_RLC_STAT_RPT_MERGE_SCH (not on sheet 14 Enable list).",
         "FOFD-081201"),
        (15, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLPDCCH: NrDuCellId={pid}, PdcchAlgoSwitch=PDCCH_AGG_LVL_COMPR_SW-1;",
         "NRDUCellPdcch", "PdcchAlgoSwitch",
         "PDCCH_AGG_LVL_COMPR_SW-0 (30/30)",
         "PDCCH_AGG_LVL_COMPR_SW-1",
         "Sheet14 MIMO-015. CCE compress. Threshold already 60.",
         "FOFD-081201"),
        (16, "TDD", "UL Boosting W4", "Seq", "", "", "", "", "",
         "UL_LOW_NOISE master + pairing children. LST LICENSE NR0S00UAHR00.", ""),
        (17, "TDD", "UL Boosting 1.0", "Enable",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, MimoFeatureSwitch=UL_LOW_NOISE_SW-1;",
         "NRDUCellFeatureSw", "MimoFeatureSwitch",
         "UL_LOW_NOISE_SW-0 (30/30)",
         "UL_LOW_NOISE_SW-1",
         "Sheet14 MIMO-019 master. Phase2 stays OFF. Aerial boost stays OFF.",
         "FOFD-091201 / NR0S00UAHR00 per cell"),
        (18, "TDD", "UL Boosting 1.0", "Enable",
         f"MOD NRDUCELLULMIMO: NrDuCellId={pid}, UlMuMimoAlgoSwitch=UL_MU_GRP_PAIR_SW-1&DIFF_WAVEFORM_PAIR_SW-1&UL_CORR_ACCELERATION_SW-1;",
         "NRDUCellUlMimo", "UlMuMimoAlgoSwitch",
         "group/flex pair off; MU_BEAM_PAIR_OPT already 1",
         "UL_MU_GRP_PAIR_SW-1&DIFF_WAVEFORM_PAIR_SW-1&UL_CORR_ACCELERATION_SW-1",
         "Sheet14 MIMO-019 children. Keep existing MU_BEAM_PAIR_OPT.",
         "FOFD-091201"),
        (19, "TDD", "CR01 do not send", "Seq", "", "", "", "", "",
         "Not in this CR. Keep sheet-14 Hold / later wave.", ""),
        (20, "TDD", "Hold", "Hold",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, WeightAlgoSwitch=PMI_WEIGHT_OPT_SW-1;",
         "NRDUCellBeamAlgo", "WeightAlgoSwitch",
         "PMI_WEIGHT_OPT-0 (30/30)",
         "do not send",
         "Sheet14 Hold — do not mix PMI + SRS weight sources the same night as CR01.",
         "FBFD-010003"),
        (21, "TDD", "Hold", "Hold",
         f"MOD NRDUCELLSRS: NrDuCellId={pid}, SrsDetectionAlgoSwitch=SRS_IC_SW-1;",
         "NRDUCellSrs", "SrsDetectionAlgoSwitch",
         "SRS_IC-0; BEAM_SPECIFIC_SRS_DENOISE already 1",
         "do not send",
         "Sheet14 MIMO-018 leftover. CR01 correctly omits it (tight MUX is ON).",
         "FOFD-061201"),
        (22, "TDD", "Hold", "Hold",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;",
         "NRDUCellDlMimo", "HighLayerMuMimoSw",
         "all multilayer bits 0 (30/30)",
         "do not send",
         "Sheet14 MIMO-016. Next CR after iBeam 1.0 KPI-green.",
         "layer licenses NR0S0DLEPU00"),
        (23, "TDD", "Hold", "Hold",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1;",
         "NRDUCellFeatureSw", "AhrSwitch",
         "AHR_CAPC_UPGRADE_PHASE2-0; Phase1+Turbo already 1",
         "do not send",
         "Sheet14 MIMO-017. After Turbo SRS-IC trial (not this CR).",
         "FOFD-061202 / NR0S00ACT200"),
        (24, "TDD", "Hold", "Hold",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, HighPrecisionBeamPhase2Sw=ON;",
         "NRDUCellFeatureSw", "HighPrecisionBeamPhase2Sw",
         "OFF (30/30)",
         "do not send",
         "iBeam 2.0/3.0 after 1.0 KPI-green. CR01 col S stays OFF.",
         "FOFD-091200"),
        (25, "TDD", "Skip", "Skip",
         f"MOD NRDUCELLPDSCH: NrDuCellId={pid}, MaxMimoLayerNum=LAYER_8;",
         "NRDUCellPdsch", "MaxMimoLayerNum",
         "LAYER_16 on 32T — never send LAYER_8",
         "keep LAYER_16",
         "Not in CR01. Do not add from FPD sample.",
         "FOFD-010020"),
    ]


DELTAS = [
    ("CR01-01", "Landed", "Step1 weights", "NRDUCellBeamAlgo",
     "WeightAlgoSwitch / SRS_WEIGHT_ESTIMATE_SW",
     "0 → 1 (30/30)", "MIMO-010 Major", "Yes — first BF leftover"),
    ("CR01-02", "Landed", "Step1 UL rank", "NRDUCellUlRank",
     "UlRankAlgoSw / UL_RANK_FAST_DECREASE_SW",
     "0 → 1 (30/30)", "MIMO-021 Minor", "Yes"),
    ("CR01-03", "Landed", "Step1 PUSCH CE", "NRDUCellPusch",
     "PuschPerformanceSwitch / PUSCH_CE_SINR_LEVEL_ENH_SW",
     "0 → 1 (30/30)", "MIMO-021 Minor", "Yes"),
    ("CR01-04", "Landed", "Step5 tracking", "NRDUCellBeamAlgo",
     "BeamOptAlgoSwitch / BEAM_TRACKING + INTELLIGENT_BEAM_SELECTION",
     "0 → 1 (30/30)", "MIMO-011 Major", "Yes (unlabeled Proposed col O)"),
    ("CR01-05", "Landed", "Step5 SSB adapt", "NRDUCellAlgoSwitch",
     "BeamOptSwitch / SSB_BEAM_ADAPT + VERTICAL_COV_IMP",
     "0 → 1 (30/30). DL_INITIAL_BEAM_SELECT already 1", "MIMO-012 Major", "Yes"),
    ("CR01-06", "Landed", "Step7 iBeam master", "NRDUCellFeatureSw",
     "HighPrecisionBeamSwitch",
     "OFF → ON (30/30). Phase2 stays OFF", "MIMO-014 Major", "Yes (unlabeled Proposed col U)"),
    ("CR01-07", "Landed", "Step7 SRS meas", "NRDUCellSrsMeas",
     "SrsMeasOptSwitch / SRS_BLIND_IS_MEAS_SW",
     "0 → 1 (30/30)", "MIMO-015 Major", "Yes"),
    ("CR01-08", "Landed", "Step7 SRS algo", "NRDUCellSrs",
     "SrsAlgoSwitch / SRS_SINR_MEAS_OPT + SRS_TIGHT_MULTIPLEXING",
     "both 0 → 1 (30/30)", "MIMO-015 + MIMO-021", "Yes. Omits SRS_IC (correct)."),
    ("CR01-09", "Landed", "Step7 beam select", "NRDUCellBeamAlgo",
     "ChannelOptAlgoSwitch / BEAM_SELECT_OPT_SW",
     "0 → 1 (30/30)", "MIMO-013 Major", "Yes (unlabeled Proposed col N)"),
    ("CR01-10", "Landed", "Step7 DL MU sch", "NRDUCellDlMimo",
     "DLMuMimoSchSupplementSw / PRECISE_SCH + ANTI_INTRF_SCH",
     "0 → 1 (30/30)", "MIMO-015 Major", "Yes"),
    ("CR01-11", "Landed", "Step7 DL sch", "NRDUCellDlSch",
     "DlSchAlgoSwitch / BWP hybrid IR + TAIL_PKT_MCS_OPT + RES_BASED_ADAPT + RLC merge",
     "four bits 0 → 1 (30/30)", "MIMO-015/020 + extra",
     "DL_RLC_STAT_RPT_MERGE_SCH is new vs sheet 14 Enable list"),
    ("CR01-12", "Landed", "Step7 PDCCH", "NRDUCellPdcch",
     "PdcchAlgoSwitch / PDCCH_AGG_LVL_COMPR_SW",
     "0 → 1 (30/30)", "MIMO-015 Major", "Yes"),
    ("CR01-13", "Landed", "Step8 UL Boost", "NRDUCellFeatureSw",
     "MimoFeatureSwitch / UL_LOW_NOISE_SW",
     "0 → 1 (30/30). PHASE2 stays 0", "MIMO-019 Major", "Master only in this bitfield"),
    ("CR01-14", "Landed", "Step8 UL MU", "NRDUCellUlMimo",
     "UlMuMimoAlgoSwitch / GRP_PAIR + DIFF_WAVEFORM + CORR_ACCEL",
     "0 → 1 (30/30)", "MIMO-019 children", "Yes"),
    ("CR01-15", "Partial", "AHR Turbo child", "NRDUCellUlPcConfig",
     "UlPwrCtrlAlgoSwitch / SRS_JOINT_PC_SW",
     "0 → 1 (30/30) without SRS_IC", "MIMO-018 Major",
     "Joint PC landed; SRS_IC correctly held (tight MUX ON)"),
    ("CR01-16", "Missing", "RF azimuth", "NRDUCellTrpBeam",
     "Azimuth(degree)",
     "0° on 30/30 — no Proposed column", "MIMO-002 Critical",
     "Not in CR01. Audit vs RF design. Tilts are 2–9° (no 255)."),
    ("CR01-17", "Hold", "PMI / open-loop weight", "NRDUCellBeamAlgo",
     "PMI_WEIGHT_OPT_SW / OPEN_LOOP_WEIGHT_OPT_SW",
     "stay 0", "MIMO-010 Hold", "Do not mix with SRS weight this night"),
    ("CR01-18", "Hold", "AHR SRS-IC", "NRDUCellSrs",
     "SrsDetectionAlgoSwitch / SRS_IC_SW",
     "stay 0", "MIMO-018", "Do not add while SRS_TIGHT_MULTIPLEXING is ON"),
    ("CR01-19", "Hold", "Step4 multilayer", "NRDUCellDlMimo",
     "HighLayerMuMimoSw / MMIMO_MULTILAYER_ENHANCE + PAIRING_PREFERRED + RANK_BOOSTING",
     "stay 0", "MIMO-016 Major", "Next CR after iBeam 1.0 KPI-green"),
    ("CR01-20", "Hold", "AHR CU 2.0", "NRDUCellFeatureSw",
     "AhrSwitch / AHR_CAPC_UPGRADE_PHASE2_SW",
     "stay 0 (Phase1+Turbo already 1)", "MIMO-017 Major", "Later loaded-hour wave"),
    ("CR01-21", "Hold", "iBeam 2.0/3.0", "NRDUCellFeatureSw",
     "HighPrecisionBeamPhase2Sw",
     "OFF (30/30)", "MIMO-027 Introduce", "After 1.0 green"),
]


def patch_cover(wb):
    ws = wb["0. Cover & Index"]
    ws["A1"].value = "  5G MIMO (all features together)  —  Deployment Workbook  v4.0"
    last = ws.max_row + 2
    r = last
    r = section(ws, r, 10, "v4.0 addition — CR01 10-site execution pack + counters on every sheet")
    r = note_bar(ws, r, 10,
                 "Sheet 16 is the Huawei CME change request CR01_MIMO Performance improvemnet_Change Request.xlsx "
                 "(10 macros × 3 sectors = 30 NR DU cells 101/102/103) taken from "
                 "5G-RAN-BASIC-TO-ADVANCE_OPTIMIZATION NSA report v2.1. "
                 "Every sheet now ends with section “Performance and Monitoring Counter”. "
                 "File: MIMO_Deployment_v4.0.xlsx  ·  v3.0 is unchanged.")
    r = headers(ws, r, ["#", "Sheet", "Maps to", "What you will find"] + [""] * 6)
    r = table_row(ws, r,
                  ["16", SHEET_NAME,
                   "CR01 CME Planned Value  ·  10 gNB 32T n41",
                   "Site list, Live vs Proposed, mapping to sheet 14, 11-col MML with real NrDuCellId 101/102/103"] + [""] * 6,
                  fills=[PALE_ORANGE] * 10, height=36)
    merge(ws, r - 1, 4, r - 1, 10)
    r = table_row(ws, r,
                  ["*", "All sheets (end)", "FPD Counter Changes / MAE KPI",
                   "New last section: Performance and Monitoring Counter (MAE 15 min, trial vs control)"] + [""] * 6,
                  fills=[PALE_GREEN] * 10, height=32)
    merge(ws, r - 1, 4, r - 1, 10)
    return r


def patch_incon(wb):
    if INCON not in wb.sheetnames:
        return
    ws = wb[INCON]
    r = ws.max_row + 2
    r = section(ws, r, COLS, "H.  v4.0 — CR01 is the 10-site execution of this sheet’s Enable list")
    r = note_bar(ws, r, COLS,
                 "User uploaded CR01 (NSA v2.1 → this repo). It realises sheet-14 W1+W3+W4 on 10 loaded 32T macros "
                 "in one CME file (not the original 3–5 site weekly split). SRS_IC is correctly omitted. "
                 "Multilayer / AHR CU 2.0 / iBeam 2.0 / PMI weight still Hold. Jump to sheet 16 for site IDs and MML.")
    r = headers(ws, r, ["#", "Sheet", "Sites", "What CR01 sends"] + [""] * 7)
    r = put11(ws, r,
              ["16", SHEET_NAME,
               "DHGUL66 DHGULV8 DHBDD13 DHGULO9 DHGUL40 DHBDD34 DHGULQ1 DHGUL05 DHBDD69 DHGULU2  (101/102/103)",
               "SRS weight + beam tracking + SSB adapt + iBeam 1.0 pack + UL_LOW_NOISE + SRS_JOINT_PC  ·  30 cells"]
              + [""] * 7,
              fills=[PALE_ORANGE] * 11, height=40)
    merge(ws, r - 1, 4, r - 1, COLS)
    href_sheet(ws.cell(r - 1, 2), SHEET_NAME, SHEET_NAME)


def build_cr01_sheet(wb):
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = wb.create_sheet(SHEET_NAME)
    setup_sheet(ws, SHEET_NAME)
    set_widths(ws, WIDTHS)
    ws.oddHeader.left.text = "CR01 MIMO performance improvement — 10-site CME execution pack (v4.0)"
    ws.oddFooter.left.text = "Source: CR01_MIMO Performance improvemnet_Change Request.xlsx  ·  NSA report v2.1 DHK 9 Sep 2026  ·  dump 8 Sep 2026"
    ws.freeze_panes = "A4"
    ws.sheet_properties.tabColor = "C00000"

    r = 1
    r = banner(ws, r, COLS,
               "  16.  CR01 MIMO Performance Improvement  —  10-site execution pack (CME Planned Value)",
               size=16, height=32)
    r = note_bar(ws, r, COLS,
                 "This sheet is the change request the user uploaded to this repo as "
                 f"{CR01_NAME} (filename typo “improvemnet” kept). Origin: "
                 "miles4482/5G-RAN-BASIC-TO-ADVANCE_OPTIMIZATION  ·  "
                 "5G_NSA_Huawei_Inconsistency_Report_v2.1_DHK_9Sep26.xlsx. "
                 "Format = Huawei CME planned-value (MO sheets + Proposed column). "
                 "Proposed bitfields list ONLY the bits being turned ON; live bits that stay 1 are not repeated. "
                 "Execute the CR01 xlsx in CME, or send the MML below. 32T only. Keep 2T2R indoor OFF. "
                 "Do not send LAYER_8. mmWave/DAS/Fusion not in CR01 (n41 FR1).")

    r = section(ws, r, COLS, "A.  CR01 identity")
    r = headers(ws, r, ["Item", "Value"] + [""] * 9)
    identity = [
        ("CR file (this repo)", CR01_NAME),
        ("Origin", "5G-RAN-BASIC-TO-ADVANCE_OPTIMIZATION / NSA Incon report v2.1 DHK 9 Sep 2026 (Incon_Agent)"),
        ("Live dump", "5G CME 8 Sep 2026 DHK  ·  564×32T32R n41 FR1 TDD 40 MHz NSA Option 3x"),
        ("Scope", "10 gNodeB × 3 sectors = 30 NR DU cells  (NrDuCellId / TRP 101, 102, 103)"),
        ("Coverage / RF", "SCENARIO_8 all 30. Azimuth 0° all 30 (no Proposed). Tilts 2–9° — no 255."),
        ("How CME Proposed works", "Only the bits in the Proposed column are set to 1. Other bits in the live bitfield remain. Do not paste Proposed as a full replacement of the live string."),
        ("Vs sheet 14 waves", "CR01 packs W1 (weights/tracking/SSB) + W3 (iBeam 1.0) + W4 (UL Boosting) + SRS_JOINT_PC in one file. Weekly split is dropped. SRS_IC omitted (correct vs SRS_TIGHT_MULTIPLEXING)."),
        ("Licenses to LST first", "NR0S00BEAM00 (iBeam 1.0) and NR0S00UAHR00 (UL Boosting) on each of the 10 gNB. AHR CU / iBeam 2.0 licenses not required tonight."),
        ("Exclude", "DHTIAA1 / DHAPT11 / DHTEJ34 (2T2R). Not in this CR."),
    ]
    for a, b in identity:
        r = put11(ws, r, [a, b] + [""] * 9, fills=[PALE_BLUE, WHITE] + [WHITE] * 9,
                  bolds=[True] + [False] * 10, height=32)
        merge(ws, r - 1, 2, r - 1, COLS)

    r = section(ws, r, COLS, "B.  Trial sites  (10 gNB × 3 cells — tilts from CR01 NRDUCellTrpBeam, no Proposed)")
    r = headers(ws, r, ["SN", "gNodeB", "NrDuCellId 101 tilt", "102 tilt", "103 tilt",
                        "Azimuth", "CoverageScenario", "TxRxMode (dump)", "In CR01", "Notes"] + [""])
    for i, (gnb, tilts) in enumerate(GNBS, 1):
        vals = [i, gnb, f"{tilts[101]}°", f"{tilts[102]}°", f"{tilts[103]}°",
                "0° / 0° / 0°", "SCENARIO_8", "32T32R (trial)", "3 cells",
                "Tilt ≠ 255. Audit azimuth vs RF design (sheet 14 MIMO-002)."]
        vals += [""]
        r = put11(ws, r, vals, fills=[alt_fill(i)] * 11, bolds=[False, True] + [False] * 9,
                  center={1, 3, 4, 5}, height=22)

    r = section(ws, r, COLS, "C.  Live vs Proposed  (30/30 identical — one row per CR01 change)")
    r = note_bar(ws, r, COLS,
                 "Parsed from CR01 MO sheets. “Landed” = Proposed turns the bit ON. "
                 "“Hold/Missing” = sheet 14 still open after this CR. Click sheet 14 for the cluster-wide dump counts.")
    r = headers(ws, r, ["ID", "Status", "Family", "MO", "Parameter / switch",
                        "Live → Proposed (30 cells)", "Sheet 14", "Note"] + [""] * 3)
    for rec in DELTAS:
        st = rec[1]
        fh = SEV_FILL.get(st, WHITE)
        vals = list(rec) + [""] * (11 - len(rec))
        r = put11(ws, r, vals, fills=[fh] * 11, bolds=[True, True] + [False] * 9,
                  center={1, 2}, height=36)
        href_sheet(ws.cell(r - 1, 7), INCON, rec[6])

    r = section(ws, r, COLS, "D.  What still remains after CR01  (do not add to this night)")
    r = bullets(ws, r, COLS, [
        "RF: electrical azimuth still 0° on all 30 CR01 cells (cluster-wide 0° was MIMO-002). Tilts are healthy (2–9°, no 255).",
        "Hold same night: PMI_WEIGHT_OPT / OPEN_LOOP_WEIGHT_OPT (SRS weight is ON). SRS_IC (tight multiplexing is ON).",
        "Next CR after iBeam 1.0 KPI-green: MMIMO_MULTILAYER_ENHANCE + PAIRING_PREFERRED + MU_RANK_BOOSTING (LAYER_16 quota already set).",
        "Later: AHR_CAPC_UPGRADE_PHASE2 + INTEL_PRCS_DL_MU_MIMO_PAIR / PDCCH_MULTI_DIM_JOINT_SCH; then iBeam 2.0/3.0.",
        "Not applicable: mmWave / DAS / Fusion / 2T2R MU / FPD sample LAYER_8.",
    ], fill_hex=PALE_GOLD)
    ws.row_dimensions[r - 1].height = 88

    r = section(ws, r, COLS, "E.  Template MML  (11-col)  ·  replace {NrDuCellId} or use section F")
    r = note_bar(ws, r, COLS,
                 "Action Enable = in CR01 Proposed. Hold/Skip = do not send. "
                 "CME file is the execution artefact; these MOD lines are the equivalent U2000/MAE MML. "
                 "One Proposed column = one command (only the bits being set to 1).")
    r = mml_header(ws, r)
    mml_start = r
    for rec in template_mml():
        r = mml_row(ws, r, rec)
    ws.auto_filter.ref = f"A{mml_start - 1}:K{r - 1}"

    r = section(ws, r, COLS, "F.  Site-specific MML  (real gNB + NrDuCellId 101/102/103 — Enable only)")
    r = note_bar(ws, r, COLS,
                 "Each Enable command from section E expanded per gNodeB. "
                 "MML cell contains three lines (cells 101, 102, 103). Filter Action=Enable. "
                 "Preferred execution remains the CR01 CME xlsx (already has 30 rows per MO).")
    r = mml_header(ws, r)
    sn = 1
    enable_cmds = [rec for rec in template_mml() if rec[3] == "Enable"]
    for gnb, _tilts in GNBS:
        r = mml_row(ws, r, (sn, "TDD", gnb, "Seq", "", "", "", "", "",
                            f"Apply the next {len(enable_cmds)} commands on {gnb} cells 101/102/103.", ""))
        sn += 1
        for rec in enable_cmds:
            _sn, rat, doc, action, cmd, mo, pid, live, prop, purpose, lic = rec
            lines = []
            for cid in CELLS:
                lines.append(cmd.replace("{NrDuCellId}", str(cid)))
            mml = "\n".join(lines)
            r = mml_row(ws, r, (sn, rat, gnb, action, mml, mo, pid, live, prop,
                                f"{gnb}: {purpose}", lic))
            sn += 1

    r = section(ws, r, COLS, "G.  Trial rules / rollback")
    r = bullets(ws, r, COLS, [
        "Pre: LST LICENSE NR0S00BEAM00 and NR0S00UAHR00 on all 10 gNB. Confirm TxRxMode=32T32R. Confirm AHR Phase1+Turbo still ON. Confirm tilt ≠ 255 (already true here).",
        "Execute: import CR01 CME planned-value (or send section F MML). One window — user chose a combined pack from NSA v2.1, not the sheet-14 weekly split.",
        "KPI: MAE 15 min, busy hour, D−7 / D+1 / D+7. Trial = these 10 gNB. Control = neighbour 32T not in CR01. Counters = last section of this sheet.",
        "Rollback: CME reverse (Proposed bits back to 0 / HighPrecisionBeamSwitch=OFF / UL_LOW_NOISE_SW-0). Do not touch already-ON SU/MU/AHR Phase1/Turbo/LAYER_16.",
        "Hard stop: SRS NI jump, PDCCH blocking up, DL/UL IBLER jump, drop/HO against control → rollback that gNB, do not add SRS_IC or iBeam 2.0.",
        "After green: next CR = multilayer (MIMO-016) then AHR CU 2.0 (MIMO-017). Still never LAYER_8, never 2T2R, never mmWave on n41.",
    ], fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 130

    add_perf_monitor(ws, r=ws.max_row + 2, cols=COLS, sheet_key=SHEET_NAME)
    return ws


def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"missing {SRC} — build v3.0 first")
    print("copy v3.0 → v4.0")
    shutil.copy2(SRC, OUT)
    wb = load_workbook(OUT)
    print("patch cover...")
    patch_cover(wb)
    print("patch incon note...")
    patch_incon(wb)
    print("CR01 sheet...")
    build_cr01_sheet(wb)
    print("append Performance and Monitoring Counter to every sheet...")
    append_counters_to_all_sheets(wb)
    colors = ["1F4E79", "2E75B6", "0D7377", "C00000", "C65911", "548235",
              "7030A0", "1F4E79", "2E75B6", "0D7377", "C00000", "C65911",
              "548235", "7030A0", "C00000", "C9A227", "C00000"]
    for i, ws in enumerate(wb.worksheets):
        ws.sheet_view.showGridLines = False
        if i < len(colors):
            ws.sheet_properties.tabColor = colors[i]
    print("saving", OUT)
    wb.save(OUT)
    print("ok", os.path.getsize(OUT), "sheets", len(wb.worksheets), "last", wb.sheetnames[-1])
    print("names:", wb.sheetnames)
    with zipfile.ZipFile(ZIP_OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(OUT, os.path.basename(OUT))
    print("zip", ZIP_OUT, os.path.getsize(ZIP_OUT))


if __name__ == "__main__":
    main()
