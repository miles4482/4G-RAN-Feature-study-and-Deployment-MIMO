#!/usr/bin/env python3
"""Build MIMO_Deployment_v6.0.xlsx from v4.0.

Combines former sheets 14 (Incon) + 15 (Suggestions) into one sheet:

  Section 1  Document suggestions (boxes; each MML row has dump Live / status / Enabled? / Action
             plus Counter monitor / Impact on KPI / short Notes / Jump to Basic)
  Section 2  Final proposal — trial/implement what is NOT enabled
  Section 3  Performance counter and Monitoring KPI

No separate dump-incon section — that check lives on the MML row.
CR01 stays as sheet 15 (was 16). v1–v5 files are unchanged.
"""
import os, re, shutil, zipfile, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font
from openpyxl.worksheet.hyperlink import Hyperlink
from mimo_excel_style import *
from mimo_common import add_kpi_header, add_kpi
from build_suggestions_sheet import suggestions, FAMILY_COLOR, mml_line, PID, TID
from build_incon_report import mml_rows as incon_mml_rows, mml_row as incon_mml_row
from mimo_counters import append_counters_to_all_sheets

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "MIMO_Deployment_v4.0.xlsx")
OUT = os.path.join(ROOT, "MIMO_Deployment_v6.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v6.0.zip")

VERSION = "v6.0"
SHEET_NAME = "14. MIMO Suggest + Incon"  # 24 chars
FINDINGS_SHEET = "16. Master Findings"  # v7.0
ACTION_SHEET = "17. Action Plan Ph1-Ph7"  # v7.0
CR01_OLD = "16. CR01 MIMO Exec Pack"
CR01_NEW = "15. CR01 MIMO Exec Pack"
OLD14 = "14. MIMO Incon Report"
OLD15 = "15. MIMO Suggestions from Doc"

WITH_PREREQ = False  # v7.0 inserts the pre-requisite columns before the dump columns
COLS = 14
WIDTHS = []
COL = {}


def set_layout(with_prereq=False):
    """Sheet-14 column layout. v6.0 has no pre-req columns; v7.0 has them."""
    global WITH_PREREQ, COLS, WIDTHS, COL
    WITH_PREREQ = with_prereq
    if with_prereq:
        COL = dict(sn=1, cmd=2, cmd_end=3, prereq=4, prereq_mml=5, live=6, status=7,
                   enabled=8, action=9, ctr=10, kpi=11, notes=12, notes_end=13,
                   jump=14, jump_end=15)
        COLS = 15
        WIDTHS = [6, 30, 30, 30, 36, 19, 13, 11, 22, 25, 20, 17, 13, 16, 12]
    else:
        COL = dict(sn=1, cmd=2, cmd_end=4, prereq=None, prereq_mml=None, live=5,
                   status=6, enabled=7, action=8, ctr=9, kpi=10, notes=11,
                   notes_end=12, jump=13, jump_end=14)
        COLS = 14
        WIDTHS = [7, 16, 18, 22, 20, 14, 12, 24, 28, 22, 16, 18, 12, 14]
    return COL


set_layout(False)


def pad(titles):
    """Pad a header list out to the current sheet width."""
    return list(titles) + [""] * max(0, COLS - len(titles))

BLUE_HDR = "5B9BD5"
YELLOW_HDR = "FFC000"
TEAL_HDR = "0D7377"
GOLD_HDR = "C9A227"
RED_HDR = "A93226"
LINK_BLUE = "0563C1"
MED = Border(
    left=Side(style="medium", color=NAVY),
    right=Side(style="medium", color=NAVY),
    top=Side(style="medium", color=NAVY),
    bottom=Side(style="medium", color=NAVY),
)

ST_FILL = {
    "Already ON": "C6EFCE",
    "Missing": "F8CBAD",
    "Hold": "FCE4D6",
    "Fix": "F4B183",
    "N/A": "E2D5F1",
    "Partial": "FFF2CC",
    "Check": "DDEBF7",
    "Enable": "F8CBAD",
    "Skip": "E2EFDA",
}


def href_sheet(cell, sheet, text=None):
    cell.value = text or sheet
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A1",
                               display=str(text or sheet))
    cell.font = Font(name="Calibri", size=9, color=LINK_BLUE, underline="single", bold=True)
    cell.alignment = align("left", "center", True)


def href_row(cell, sheet, row, text):
    cell.value = text
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A{row}", display=str(text))
    cell.font = Font(name="Calibri", size=9, color=LINK_BLUE, underline="single", bold=True)
    cell.alignment = align("left", "center", True)


def box_border(ws, r1, r2, cols=None):
    cols = COLS if cols is None else cols
    for r in range(r1, r2 + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            left = "medium" if c == 1 else "thin"
            right = "medium" if c == cols else "thin"
            top = "medium" if r == r1 else "thin"
            bottom = "medium" if r == r2 else "thin"
            cell.border = Border(
                left=Side(style=left, color=NAVY),
                right=Side(style=right, color=NAVY),
                top=Side(style=top, color=NAVY),
                bottom=Side(style=bottom, color=NAVY),
            )


def label_row(ws, r, label, text, fill_hex, height=None):
    put(ws, r, 1, label, size=10, bold=True, color=WHITE, fill_hex=NAVY, h="center", v="center", border=True)
    merge(ws, r, 2, r, COLS)
    put(ws, r, 2, text, size=10, fill_hex=fill_hex, h="left", v="top", border=True)
    for c in range(3, COLS + 1):
        ws.cell(r, c).fill = fill(fill_hex)
        ws.cell(r, c).border = thin
    if height is None:
        height = min(78, max(28, 16 + (text.count("\n") + 1) * 13 + len(text) // 140 * 8))
    ws.row_dimensions[r].height = height
    return r + 1


def putn(ws, r, values, fills=None, bolds=None, height=None, center=None, font_color=None):
    n = COLS
    vals = list(values) + [""] * (n - len(values))
    vals = vals[:n]
    for i, v in enumerate(vals, 1):
        fh = fills[i - 1] if fills and i - 1 < len(fills) else None
        b = bolds[i - 1] if bolds and i - 1 < len(bolds) else False
        h = "center" if center and i in center else "left"
        col = font_color if font_color else "000000"
        put(ws, r, i, v, size=9, bold=b, color=col, fill_hex=fh, h=h, v="top", border=True)
    if height is None:
        longest = max((len(str(v)) if v is not None else 0) for v in vals)
        lines = max((str(v).count("\n") + 1) for v in vals)
        height = min(72, max(22, 16 + longest // 90 * 10, 14 * lines))
    ws.row_dimensions[r].height = height
    return r + 1


# Keyword → (counter to monitor, impact on KPI, short note). First match wins.
MONITOR = [
    ("SRS_WEIGHT_ESTIMATE", "N.UL.SRS.PreSINR.* / N.ChMeas.PDSCH.MCS.k / N.ThpVol.DL",
     "User DL Average Throughput (DU)", "SRS weights for large-packet UEs."),
    ("PMI_WEIGHT_OPT", "N.RRC.ReEst.Att / N.ChMeas.PDSCH.MCS.k",
     "User DL tput; RRC reconfig", "Hold until SRS-weight KPI is green."),
    ("OPEN_LOOP_WEIGHT_OPT", "N.ChMeas.PDSCH.MCS.k / N.DL.SCH.*.ErrTB.Ibler",
     "DL MCS / IBLER", "Hold first night; do not mix weight sources."),
    ("DL_PMI_SRS_ADAPT", "N.ThpVol.DL / N.ChMeas.PDSCH.MCS.k",
     "User DL Average Throughput (DU)", "Parent of non-AS PMI weight. Already ON 32T."),
    ("DISTRIBUTED_MODE", "N.ThpVol.DL / N.UL.SRS.PreSINR.*",
     "User DL tput (weight-opt prerequisite)", "Sensing must stay DISTRIBUTED on AAU."),
    ("SrsWeightValidityPeriod", "N.UL.SRS.PreSINR.* / N.ThpVol.DL",
     "User DL tput", "How long an SRS weight stays valid."),
    ("SrsNonASFixedWeightType", "N.ChMeas.PDSCH.MCS.k",
     "User DL tput (non-AS UEs)", "Needs DL_PMI_SRS_ADAPT=1."),
    ("FR1MaxCellCsirsPortNum", "N.ChMeas.CQI.SingleCW.k",
     "CQI / CSI quality", "8PORT already the DHK baseline."),
    ("SRS_SINR_MEAS_OPT", "N.UL.SRS.PreSINR.* / N.SRS.NI.Avg",
     "User UL Average Throughput (DU)", "SRS SINR meas opt — UL + DL-weight input."),
    ("UL_RANK_FAST_DECREASE", "N.PUSCH.TbUl.Rank1–2 / N.UL.SCH.*.ErrTB.Ibler",
     "UL IBLER / UL tput", "Drops rank fast on poor SRS."),
    ("PUSCH_CE_SINR_LEVEL_ENH", "N.ChMeas.PUSCH.MCS.k / N.UL.SCH.*.ErrTB.Ibler",
     "User UL Average Throughput (DU)", "PUSCH CE in interference."),
    ("SrsPreSinrJudgeThld", "N.UL.SRS.PreSINR.*",
     "SRS quality", "Gate for using an SRS sample."),
    ("CHN_MEASURE_CPU_DEC", "N.User.RRCConn.Avg (CPU, not air KPI)",
     "Stability (not tput)", "gNodeB-level CPU decrease."),
    ("DlSchOptTimeThld", "N.ThpVol.DL / N.PRB.DL.Used.Avg",
     "User DL Average Throughput (DU)", "How long DL sch opt is held."),
    ("DlAdaptSchTimeThld", "N.ThpVol.DL",
     "User DL Average Throughput (DU)", "Adaptive-sch timer."),
    ("MaxMimoLayerCnt", "N.ChMeas.MIMO.UL.Trans.Layer.Max",
     "UL layer / UL tput", "Do not downgrade live 32T quota."),
    ("MaxMimoLayerNum", "N.ChMeas.MIMO.DL.Transmission.Layer.Max",
     "DL layer / DL tput", "Keep LAYER_16 — never send LAYER_8."),
    ("DL_RANK_ADAPT", "N.PDSCH.InitTbDl.Rank1–4",
     "User DL Average Throughput (DU)", "Already ON 32T."),
    ("SuMimoPwrCtrlProtectThld", "N.ThpVol.DL / N.DL.SCH.*.ErrTB.Ibler",
     "DL IBLER under SU power", "SU power-protect."),
    ("SU_DMRS_OH_ADAPT", "N.ThpVol.DL / N.PRB.DL.Used.Avg",
     "User DL tput (more PDSCH REs)", "DMRS overhead deduct."),
    ("SRS_PRECODE_OPT", "N.ChMeas.PDSCH.MCS.k",
     "User DL tput on SRS DTX", "Precoding when SRS is missing."),
    ("INTRA_GNB_DL_JT", "N.ThpVol.DL / N.UECntx.AbnormRel",
     "DL tput (only if JT intended)", "Confirm JT before sending."),
    ("UL_MU_MIMO_SW", "N.ChMeas.MIMO.UL.Pair.Layer / Pair.PRB",
     "User UL Average Throughput (DU)", "UL MU master. Already ON 32T."),
    ("DL_MU_MIMO_SW", "N.ChMeas.MIMO.DL.Pair.Layer.Avg / Pair.PRB",
     "User DL Average Throughput (DU)", "DL MU master. Already ON 32T."),
    ("PDCCH_MU_SW", "N.CCE.Used.Avg / N.CCE.DL.AllocReq.Num",
     "PDCCH CCE success / DL tput", "Watch CCE blocking."),
    ("DlPmiMuMimoSpaceIsoThld", "N.ChMeas.MIMO.DL.Pair.Layer.Avg / N.DL.SCH.*.Ibler",
     "MU pair quality vs IBLER", "Higher = fewer pairs."),
    ("DlSrsMuMimoSpaceIsoThld", "N.ChMeas.MIMO.DL.Pair.Layer.Avg",
     "MU pair quality", "SRS isolation gate."),
    ("DlMuMimoSrsPreSinrThld", "N.UL.SRS.PreSINR.* / N.ChMeas.MIMO.DL.Pair.PRB",
     "MU pairing start", "Live −50 matches FPD."),
    ("DlMuMimoGroupMode", "N.ChMeas.MIMO.DL.Pair.Layer.Avg",
     "MU pair count", "ISOLATION_CORRELATION already live."),
    ("UlMuMimoCorrThld", "N.ChMeas.MIMO.UL.Pair.Layer",
     "UL MU pairing", "Live 9 matches FPD."),
    ("UlMuMimoSinrThld", "N.ChMeas.MIMO.UL.Pair.Layer / N.UL.SCH.*.Ibler",
     "UL MU vs UL IBLER", "Live −20 matches FPD."),
    ("DlMuBackToSuSeThld", "N.ChMeas.MIMO.DL.Pair.Layer.Avg",
     "MU stay vs SU fallback", "5 live; multilayer sample uses 0."),
    ("MaxPairLayerNum", "N.CCE.DL.AllocReq.Num",
     "PDCCH MU layers", "PDCCH pair cap."),
    ("HEAVY_LOAD_SCH_PRI_OPT", "N.ThpVol.DL / N.PRB.DL.Used.Avg",
     "Fairness under load", "Already ON 32T."),
    ("MMIMO_MULTILAYER_ENHANCE", "N.ChMeas.MIMO.DL.Pair.Layer.Avg",
     "User DL Average Throughput (DU)", "Spends LAYER_16 quota. Not in CR01."),
    ("MU_RANK_BOOSTING", "N.PDSCH.InitTbDl.Rank3 / Rank4",
     "DL rank / DL tput", "After iBeam 1.0 green."),
    ("SRS_BLIND_IS_SW", "N.SRS.NI.Avg / N.UL.SRS.PreSINR.*",
     "SRS quality / DL MU", "Multilayer child."),
    ("MU_MIMO_PAIRING_PREFERRED", "N.ChMeas.MIMO.DL.Pair.PRB",
     "DL MU PRB share", "Prefer MU when gain exists."),
    ("SRS_MEAS_ACCELERATING", "N.UL.SRS.PreSINR.*",
     "SRS freshness / DL tput", "Faster SRS for multilayer."),
    ("DL_HYBRID_PRECODING", "N.ChMeas.PDSCH.MCS.k / N.DL.SCH.*.Ibler",
     "DL MCS under MU", "iBeam 3.0 / multilayer overlap."),
    ("TAIL_PKT_MCS_OPT", "N.DL.SCH.QPSK.ErrTB.Ibler / N.Thp.DL.Samp.Indexk",
     "Tail-user DL tput / IBLER", "iBeam 1.0 child. CR01 ON."),
    ("MULTILAYER_DEMOD_ENH", "N.ChMeas.MIMO.UL.Pair.Layer",
     "User UL Average Throughput (DU)", "UL multilayer. Hold."),
    ("MU_MIMO_FLEX_PAIR", "N.ChMeas.MIMO.UL.Pair.PRB",
     "UL MU pairing", "Hold with UL multilayer."),
    ("RES_BASED_ADAPT_ULSCH", "N.PRB.UL.Used.Avg / N.ThpVol.UL",
     "User UL tput under load", "Hold."),
    ("CoverageScenario", "N.User.OptimalSSBBeam.Avg / N.MAC.ThpVol.DL.OptimalSSB",
     "SSB coverage / HO", "Use RF-planned scenario; do not copy Tilt=255."),
    ("Tilt=255", "N.ThpVol.DL / N.UECntx.AbnormRel",
     "Coverage / interference", "255 on live AAU = BF not applied. Fix first."),
    ("DL_INITIAL_BEAM_SELECT", "N.User.OptimalSSBBeam.Avg",
     "Idle/connected start of BF", "Already ON 32T."),
    ("PDCCH_BEAM_ROBUST", "N.CCE.Used.Avg / N.RRC.SetupReq.Succ",
     "Access / PDCCH", "Control-beam robustness."),
    ("SRS_BEAM_SELECT_OPT", "N.UL.SRS.PreSINR.* / N.ThpVol.UL",
     "UL beam / UL tput", "SRS beam select."),
    ("BEAM_TRACKING", "N.ChMeas.PDSCH.MCS.k / N.ThpVol.DL",
     "User DL tput in mobility", "CR01 ON 30 cells. Watch drop."),
    ("INTELLIGENT_BEAM_SELECTION", "N.ChMeas.PDSCH.MCS.k / N.User.OptimalSSBBeam.Avg",
     "User DL tput in mobility", "With beam tracking."),
    ("SSB_BEAM_ADAPT", "N.User.OptimalSSBBeam.Avg / N.UECntx.AbnormRel",
     "Coverage / HO / drop", "CR01 ON. Rollback if drop rises."),
    ("SSB_BEAM_VERTICAL_COV", "N.User.OptimalSSBBeam.Avg",
     "High-rise coverage", "With SSB adapt."),
    ("BEAM_SELECT_OPT", "N.ChMeas.PDSCH.MCS.k / N.CCE.Used.Avg",
     "PDSCH/PDCCH beam quality", "iBeam 1.0 child. CR01 ON."),
    ("SCENARIO_BEAM_OPT", "N.MAC.ThpVol.DL.OptimalSSB",
     "SSB scenario tput", "After RF scenario freeze."),
    ("AHR_PHASE1", "N.ChMeas.CQI.SingleCW.k",
     "CQI / DL experience", "Already ON 32T. Do not re-send."),
    ("CSIRS_INTRF_STATIC_AVOID", "N.ChMeas.CQI.SingleCW.k",
     "CSI quality", "Already ON with Phase1."),
    ("FD_RESOURCE", "N.ChMeas.CQI.SingleCW.k",
     "CSI resource", "Already ON."),
    ("AHR_EXP_TURBO", "N.ThpVol.DL / N.ChMeas.PDSCH.MCS.k",
     "User DL Average Throughput (DU)", "Turbo master already ON."),
    ("SRS_IC_SW", "N.SRS.NI.Avg / N.UL.SRS.PreSINR.*",
     "SRS NI / DL weight quality", "Do not stack with tight MUX (CR01)."),
    ("SRS_JOINT_PC", "N.UL.SRS.PreSINR.* / N.UL.RSSI.Avg",
     "SRS quality / UL tput", "CR01 ON without SRS_IC."),
    ("TAIL_PKT_SCH_OPT", "N.Thp.DL.Samp.Indexk",
     "Tail DL tput", "Already ON (different from TAIL_PKT_MCS_OPT)."),
    ("AHR_CAPC_UPGRADE", "N.ChMeas.MIMO.DL.Pair.Layer.Avg / N.PRB.DL.Used.Avg",
     "Loaded-hour DL tput", "Hold until Turbo SRS-IC trial is green."),
    ("PDCCH_MULTI_DIM_JOINT_SCH", "N.CCE.Used.Avg / N.ChMeas.MIMO.DL.Pair.PRB",
     "PDCCH + MU capacity", "AHR CU child. Hold."),
    ("HighPrecisionBeamSwitch=ON", "N.ThpVol.DL / N.CCE.DL.AggLvl* / N.DL.SCH.*.Ibler",
     "User DL tput in interference", "iBeam 1.0 master. CR01 ON 10 sites. LST NR0S00BEAM00."),
    ("HighPrecisionBeamPhase2", "N.ChMeas.PDSCH.MCS.k / N.DL.SCH.*.Ibler",
     "DL robustness (2nd wave)", "After 1.0 KPI-green. Not in CR01."),
    ("HighPrecisionBeamPhase3", "N.ThpVol.DL / N.DL.SCH.*.Ibler",
     "Residual interference tput", "Last iBeam wave."),
    ("SRS_BLIND_IS_MEAS", "N.UL.SRS.PreSINR.* / N.SRS.NI.Avg",
     "SRS quality", "iBeam 1.0. CR01 ON."),
    ("SRS_TIGHT_MULTIPLEXING", "N.SRS.NI.Avg / N.UL.SRS.PreSINR.* / N.User.RRCConn.Avg",
     "SRS capacity vs NI", "CR01 ON. Keep SRS_IC OFF."),
    ("DL_BWP_HYBRID_INTRF_RANDOM", "N.BWP.ThpVol.DL / N.DL.SCH.*.Ibler",
     "DL IBLER / tput in ISI", "iBeam 1.0. CR01 ON."),
    ("DL_RLC_STAT_RPT_MERGE", "N.QoS.DL.PktDelayAirInterface.* / N.Traffic.DL.RlcFirstPktDelay.Time",
     "DL latency / first-packet delay", "CR01 ON."),
    ("PDCCH_AGG_LVL_COMPR", "N.CCE.DL.AllocReq.Num / N.CCE.DL.AggLvl*",
     "PDCCH blocking", "CR01 ON. Blocking up = fail."),
    ("DL_MU_PRECISE_SCH", "N.ChMeas.MIMO.DL.Pair.Layer.Avg / N.DL.SCH.256QAM.ErrTB.Ibler",
     "MU quality vs 256QAM IBLER", "CR01 ON."),
    ("DL_MU_ANTI_INTRF", "N.ChMeas.MIMO.DL.Pair.PRB / N.DL.SCH.*.Ibler",
     "MU under interference", "CR01 ON."),
    ("RES_BASED_DL_ADAPT_SCH", "N.PRB.DL.Used.Avg / N.ThpVol.DL",
     "Loaded-hour DL tput", "CR01 ON."),
    ("DL_ROBUST_WEIGHT", "N.ChMeas.PDSCH.MCS.k / N.DL.SCH.*.Ibler",
     "DL MCS in interference", "iBeam 2.0. Hold."),
    ("UL_LOW_NOISE_PHASE2", "N.ThpVol.UL / N.UL.SCH.*.Ibler",
     "User UL tput (2.0)", "After 1.0 green. Needs sync."),
    ("UL_LOW_NOISE", "N.ThpVol.UL / N.ChMeas.MIMO.UL.Pair.Layer / N.UL.SCH.*.Ibler",
     "User UL Average Throughput (DU)", "Boosting 1.0 master. CR01 ON. LST NR0S00UAHR00."),
    ("UL_MU_GRP_PAIR", "N.ChMeas.MIMO.UL.Pair.Layer",
     "UL MU pairing", "CR01 ON."),
    ("DIFF_WAVEFORM_PAIR", "N.ChMeas.MIMO.UL.Pair.PRB / N.UL.SCH.*.Ibler",
     "UL MU mixed waveform", "CR01 ON."),
    ("UL_CORR_ACCELERATION", "N.ChMeas.MIMO.UL.Pair.Layer",
     "UL MU pairing speed", "CR01 ON."),
    ("PUSCH_COORD_PWR_CTRL", "N.UL.RSSI.Avg / N.UL.SCH.*.Ibler",
     "UL IBLER (multi-cell PC)", "Needs DSP CLKTST."),
    ("MULTI_BEAM_RX_ENH", "N.ThpVol.UL / N.ChMeas.PUSCH.MCS.k",
     "User UL tput", "Boosting 2.0 child."),
    ("DM_MIMO", "N.ThpVol.DL.LastSlot / N.UL.NI.Avg.PRB*",
     "DAS coverage (N/A here)", "n41 single-TRP — do not enable."),
    ("INTRA_CELL_MIMO", "N.ChMeas.MIMO.DL.Pair.PRB / N.UL.PUSCH.SINR.Index*",
     "Fusion 128T (N/A here)", "Not deployed."),
    ("NRDUCELLTRPMMWAVBEAM", "N.PDSCH.InitTbDl.Rank1–4",
     "FR2 tput (N/A)", "n41 FR1 — skip."),
    ("VOL_BASED_BEAM_MULTIPLEX", "N.ThpVol.DL",
     "FR2 FDM (N/A)", "Do not send on n41."),
    ("ANTENNAPORTOPTDET", "DSP result CROSSED/CORRECT (not PM)",
     "Commissioning, not tput", "4T4R only — skip on 32T CR01."),
    ("FREQ_SEL_SCH", "N.ThpVol.DL / N.ChMeas.PDSCH.MCS.k",
     "User DL tput (freq-sel UEs)", "Hold until CSI stable."),
    ("LOAD_BASED_DL_EXP_SCH", "N.ThpVol.DL / N.PRB.DL.Used.Avg",
     "Busy-hour DL tput", "Hold for loaded 32T."),
    ("SMART_SCH_AND_LINK_ADAPT", "N.ThpVol.DL / N.DL.SCH.*.Ibler",
     "DL tput (performance pack)", "License + TAC. Hold."),
    ("HIGH_CAPACITY_EXP_IMP", "N.ThpVol.DL / N.PRB.DL.Used.Avg",
     "Loaded-hour DL tput", "License + TAC. Hold."),
]


def extras(cmd, note):
    for key, ctr, kpi, short in sorted(MONITOR, key=lambda x: len(x[0]), reverse=True):
        if key in cmd:
            return ctr, kpi, short
    return ("N.ThpVol.DL / N.ThpVol.UL",
            "User DL/UL Average Throughput (DU)",
            (note[:90] + ".") if note else "Watch BH vs neighbour control.")


# Pre-requisites: what must already be ON (and in what order) before a switch is accepted.
# Keyword → (pre-req switch / parameter, MML or LST to run first). Longest match wins.
# RETCODE 2147616329 on the 10-Sep trial proved the BEAM_TRACKING_SW entry.
PREREQ = [
    ("BEAM_TRACKING_SW",
     "MANDATORY: NRDUCellChnCovAlgo.DlCoverageAlgoSwitch = SUPER_COVERAGE_SW-1. "
     "Without it the MOD is rejected with RETCODE 2147616329.",
     f"1) MOD NRDUCELLCHNCOVALGO: NrDuCellId={PID}, DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1;   "
     "// run FIRST, then send BEAM_TRACKING_SW on its own line"),
    ("INTELLIGENT_BEAM_SELECTION_SW",
     "BEAM_TRACKING_SW-1 (which itself needs SUPER_COVERAGE_SW-1). Never bundle both bits in one MML — "
     "MOD is atomic, so one invalid bit rejects the whole line.",
     f"1) MOD NRDUCELLCHNCOVALGO: NrDuCellId={PID}, DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1;  "
     f"2) MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, BeamOptAlgoSwitch=BEAM_TRACKING_SW-1;"),
    ("SRS_WEIGHT_ESTIMATE_SW",
     "BeamPerceiveMode=DISTRIBUTED_MODE and AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-1 "
     "(both already ON in DHK). Send on its own MML line — do not bundle with beam bits.",
     f"LST NRDUCELLBEAMALGO: NrDuCellId={PID};   // confirm BeamPerceiveMode=DISTRIBUTED_MODE first"),
    ("PMI_WEIGHT_OPT_SW",
     "BeamPerceiveMode=DISTRIBUTED_MODE + SRS_WEIGHT_ESTIMATE KPI already green (do not mix weight sources).",
     f"MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, BeamPerceiveMode=DISTRIBUTED_MODE;"),
    ("OPEN_LOOP_WEIGHT_OPT_SW",
     "BeamPerceiveMode=DISTRIBUTED_MODE + SRS weight night already green.",
     f"MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, BeamPerceiveMode=DISTRIBUTED_MODE;"),
    ("SrsNonASFixedWeightType",
     "AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-1 must be ON or the weight type is ignored.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-1;"),
    ("SCENARIO_BEAM_OPT_SW",
     "NRDUCellTrpBeam.CoverageScenario already set to the RF-planned scenario.",
     f"MOD NRDUCELLTRPBEAM: NrDuCellTrpId={TID}, CoverageScenario=SCENARIO_n;"),
    ("SSB_BEAM_VERTICAL_COV_IMP_SW",
     "SSB_BEAM_ADAPT_SW-1 first (vertical improve is its child).",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, BeamOptSwitch=SSB_BEAM_ADAPT_SW-1;"),
    ("SSB_BEAM_ADAPT_SW",
     "DL_INITIAL_BEAM_SELECT_SW-1 (already ON) + Tilt/Azimuth per RF design. Never with Tilt=255.",
     f"LST NRDUCELLTRPBEAM: NrDuCellTrpId={TID};   // Tilt must not be 255"),
    ("HighPrecisionBeamPhase2Sw",
     "HighPrecisionBeamSwitch=ON and iBeam 1.0 already KPI-green.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("HighPrecisionBeamPhase3Sw",
     "HighPrecisionBeamPhase2Sw=ON and 2.0 KPI-green.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase2Sw=ON;"),
    ("HighPrecisionBeamSwitch=ON",
     "License NR0S00BEAM00 + MU-MIMO masters DL_MU_MIMO_SW-1 / UL_MU_MIMO_SW-1 (Step3).",
     "1) LST LICENSE;   2) LST NRDUCELLALGOSWITCH: NrDuCellId=" + PID + ";   // MuMimoSwitch must show DL+UL MU"),
    ("SRS_TIGHT_MULTIPLEXING_SW",
     "HighPrecisionBeamSwitch=ON  AND  SrsDetectionAlgoSwitch=SRS_IC_SW-0 in the same cell. "
     "Both SRS bits in one window is what degraded Ph1 pairing.",
     f"MOD NRDUCELLSRS: NrDuCellId={PID}, SrsDetectionAlgoSwitch=SRS_IC_SW-0;   // keep IC OFF"),
    ("SRS_IC_SW",
     "AhrSwitch=AHR_EXP_TURBO_PHASE2_SW-1 AND SrsAlgoSwitch=SRS_TIGHT_MULTIPLEXING_SW-0. "
     "Never the same night as tight multiplexing.",
     f"MOD NRDUCELLSRS: NrDuCellId={PID}, SrsAlgoSwitch=SRS_TIGHT_MULTIPLEXING_SW-0;"),
    ("SRS_JOINT_PC_SW",
     "AhrSwitch=AHR_EXP_TURBO_PHASE2_SW-1 (Turbo master).",
     f"LST NRDUCELLFEATURESW: NrDuCellId={PID};   // AhrSwitch shows AHR_EXP_TURBO_PHASE2"),
    ("AHR_EXP_TURBO_PHASE2_SW",
     "AhrSwitch=AHR_PHASE1_SW-1 first.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, AhrSwitch=AHR_PHASE1_SW-1;"),
    ("AHR_CAPC_UPGRADE_PHASE2_SW",
     "AHR_PHASE1_SW-1 + AHR_EXP_TURBO_PHASE2_SW-1 + DL_MU_MIMO_SW-1.",
     f"LST NRDUCELLFEATURESW: NrDuCellId={PID};   // both AHR phases must be 1"),
    ("PDCCH_MULTI_DIM_JOINT_SCH_SW",
     "AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1 (master).",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1;"),
    ("MMIMO_MULTILAYER_ENHANCE_SW",
     "DL_MU_MIMO_SW-1 + MaxMimoLayerNum ≥ LAYER_8 (DHK is LAYER_16) + license NR0S0DLEPU00.",
     "1) LST LICENSE;   2) LST NRDUCELLPDSCH: NrDuCellId=" + PID + ";   // MaxMimoLayerNum"),
    ("MU_RANK_BOOSTING_SW",
     "HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1 (master) first.",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;"),
    ("MU_MIMO_PAIRING_PREFERRED_SW",
     "HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1 (master) first.",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;"),
    ("SRS_MEAS_ACCELERATING_SW",
     "HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1 (master) first.",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;"),
    ("SRS_BLIND_IS_SW",
     "HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1 (master) first.",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;"),
    ("MULTILAYER_DEMOD_ENH_SW",
     "MuMimoSwitch=UL_MU_MIMO_SW-1 + license NR0S0ULEPU00.",
     f"LST NRDUCELLALGOSWITCH: NrDuCellId={PID};   // UL_MU_MIMO_SW must be 1"),
    ("MU_MIMO_FLEX_PAIR_SW",
     "UlHighLayerMuMimoSwitch=MULTILAYER_DEMOD_ENH_SW-1 first.",
     f"MOD NRDUCELLULMIMO: NrDuCellId={PID}, UlHighLayerMuMimoSwitch=MULTILAYER_DEMOD_ENH_SW-1;"),
    ("UL_LOW_NOISE_PHASE2_SW",
     "MimoFeatureSwitch=UL_LOW_NOISE_SW-1 + inter-gNB time sync.",
     "1) DSP CLKTST;   2) LST NRDUCELLFEATURESW: NrDuCellId=" + PID + ";"),
    ("PUSCH_COORD_PWR_CTRL_SW",
     "UL_LOW_NOISE_PHASE2_SW-1 + TIME_SYNC confirmed.",
     "DSP CLKTST;   // must report time sync before coordinated PC"),
    ("UL_MU_GRP_PAIR_SW",
     "MimoFeatureSwitch=UL_LOW_NOISE_SW-1 + MuMimoSwitch=UL_MU_MIMO_SW-1.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, MimoFeatureSwitch=UL_LOW_NOISE_SW-1;"),
    ("DIFF_WAVEFORM_PAIR_SW",
     "MimoFeatureSwitch=UL_LOW_NOISE_SW-1 + UL_MU_GRP_PAIR_SW-1.",
     f"MOD NRDUCELLULMIMO: NrDuCellId={PID}, UlMuMimoAlgoSwitch=UL_MU_GRP_PAIR_SW-1;"),
    ("UL_CORR_ACCELERATION_SW",
     "MimoFeatureSwitch=UL_LOW_NOISE_SW-1 + UL_MU_GRP_PAIR_SW-1.",
     f"MOD NRDUCELLULMIMO: NrDuCellId={PID}, UlMuMimoAlgoSwitch=UL_MU_GRP_PAIR_SW-1;"),
    ("UL_LOW_NOISE_SW",
     "License NR0S00UAHR00 + 32T cell (keep OFF on 2T2R indoor).",
     "LST LICENSE;   // NR0S00UAHR00 must be allocated"),
    ("DL_MU_PRECISE_SCH_SW",
     "HighPrecisionBeamSwitch=ON + MuMimoSwitch=DL_MU_MIMO_SW-1. Tightens pairing — needs healthy SRS.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("DL_MU_ANTI_INTRF_SCH_SW",
     "HighPrecisionBeamSwitch=ON + DL_MU_MIMO_SW-1. Tightens pairing — needs healthy SRS.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("PRECISE_MUMIMO_EVAL_SW",
     "HighPrecisionBeamPhase2Sw=ON (iBeam 2.0 master).",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase2Sw=ON;"),
    ("FAR_UE_RANK_OPT_SW",
     "HighPrecisionBeamPhase2Sw=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase2Sw=ON;"),
    ("DL_CORR_ACCELERATION_SW",
     "HighPrecisionBeamPhase2Sw=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase2Sw=ON;"),
    ("DL_ROBUST_WEIGHT_SW",
     "HighPrecisionBeamPhase2Sw=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase2Sw=ON;"),
    ("DL_SELF_FUSION_WEIGHT_SW",
     "HighPrecisionBeamPhase3Sw=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase3Sw=ON;"),
    ("DL_SMART_AMC_SW",
     "HighPrecisionBeamPhase3Sw=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase3Sw=ON;"),
    ("DYNAMIC_CLUSTER_GROUP",
     "HighPrecisionBeamPhase2Sw=ON (overrides Step3 ISOLATION_CORRELATION only while 2.0 is ON).",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamPhase2Sw=ON;"),
    ("SRS_BLIND_IS_MEAS_SW",
     "HighPrecisionBeamSwitch=ON (iBeam 1.0 master).",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("BEAM_SELECT_OPT_SW",
     "HighPrecisionBeamSwitch=ON. Send on its own MML line (it was lost in the Ph1 bundled command).",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("DL_BWP_HYBRID_INTRF_RANDOM_SW",
     "HighPrecisionBeamSwitch=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("DL_RLC_STAT_RPT_MERGE_SCH_SW",
     "HighPrecisionBeamSwitch=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("RES_BASED_DL_ADAPT_SCH_SW",
     "HighPrecisionBeamSwitch=ON.",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("PDCCH_AGG_LVL_COMPR_SW",
     "HighPrecisionBeamSwitch=ON. Set AggLvlComprCceUsageThld=60 with it.",
     f"MOD NRDUCELLPDCCHALGO: NrDuCellId={PID}, AggLvlComprCceUsageThld=60;"),
    ("TAIL_PKT_MCS_OPT_SW",
     "HighPrecisionBeamSwitch=ON (iBeam) or MMIMO_MULTILAYER_ENHANCE_SW-1 (multilayer).",
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, HighPrecisionBeamSwitch=ON;"),
    ("PDCCH_MU_SW",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 and UL_MU_MIMO_SW-1 first.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DL_MU_MIMO_SW",
     "Step2 SU layers set + license NR0S00MUMM00. Keep OFF on 2T2R.",
     "LST LICENSE;   // NR0S00MUMM00 + layer capacity units"),
    ("UL_MU_MIMO_SW",
     "Step2 SU layers set + license NR0S00MUMM00. Keep OFF on 2T2R.",
     "LST LICENSE;   // NR0S00MUMM00 + layer capacity units"),
    ("MaxMimoLayerNum",
     "DL layer capacity licence NR0S0DLEPU00. Never downgrade live LAYER_16.",
     f"LST NRDUCELLPDSCH: NrDuCellId={PID};   // read the live value before any MOD"),
    ("MaxMimoLayerCnt",
     "UL layer capacity licence NR0S0ULEPU00. Never downgrade the live value.",
     f"LST NRDUCELLPUSCH: NrDuCellId={PID};   // read the live value before any MOD"),
    ("SMART_SCH_AND_LINK_ADAPT_SW",
     "Performance-pack license + Huawei TAC approval.",
     "LST LICENSE;   // raise TAC case before enabling"),
    ("HIGH_CAPACITY_EXP_IMP_SW",
     "Performance-pack license + Huawei TAC approval.",
     "LST LICENSE;   // raise TAC case before enabling"),
    ("UL_SU_SINR_INTEL_PREDICT_SW",
     "AI model loaded and valid.",
     f"DSP NRDUCELLAISCHMODEL: NrDuCellId={PID};"),
    ("UL_PRECISE_MCS_OPT_SW",
     "AI model loaded and valid + UL_LOW_NOISE_PHASE2_SW-1.",
     f"DSP NRDUCELLAISCHMODEL: NrDuCellId={PID};"),
    ("DM_MIMO_SERVICE_SWITCH",
     "Slave TRP already added (TrpType=SLAVE) on the same NR DU cell.",
     "MOD NRDUCELLTRP: NrDuCellTrpId={SlaveTrpId}, TrpType=SLAVE;"),
    ("MUMIMO_SINR_ENH_SW",
     "GNBCLUSTER (INTRA_CELL_MIMO) + GNBMIMOCLUSTERCELL already added.",
     "1) ADD GNBCLUSTER: ClusterType=INTRA_CELL_MIMO;   2) ADD GNBMIMOCLUSTERCELL: ...;"),
    ("SINGLE_TRP_SSB_TRANS_SW",
     "Fusion cluster active + MUMIMO_SINR_ENH_SW-1.",
     "LST GNBCLUSTER;   // cluster must be INTRA_CELL_MIMO"),
    ("NRDUCELLTRPMMWAVBEAM",
     "FR2 cell. DEA NRCELL before the MOD and ACT NRCELL after. N/A on n41 FR1.",
     "1) DEA NRCELL: NrCellId=...;   2) MOD ...;   3) ACT NRCELL: NrCellId=...;"),
    ("VOL_BASED_BEAM_MULTIPLEX_SW",
     "FR2 multi-beam cell. Do not send on FR1 n41.",
     "LST NRDUCELL;   // confirm FR2 before sending"),
    ("ANTENNAPORTOPTDET",
     "TxRxMode=4T4R + NORMAL_CELL + at least 2 intra-frequency cells in service. No VSWR alarm.",
     "1) LST NRDUCELLTRP;   2) LST ALMAF;   // clear VSWR first"),
    ("SUPER_COVERAGE_SW",
     "None — this switch is itself the pre-requisite of BEAM_TRACKING_SW.",
     "Run this line first, then BEAM_TRACKING_SW."),
    ("BeamPerceiveMode",
     "None — this is itself the pre-requisite of SRS_WEIGHT_ESTIMATE_SW / PMI_WEIGHT_OPT_SW.",
     "Run this line before any weight switch."),
    ("DL_PMI_SRS_ADAPT_SW",
     "None — this is itself the pre-requisite of SrsNonASFixedWeightType and the SRS weight path.",
     "Run this line before the weight switches."),
    ("SrsWeightValidityPeriod",
     "WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1 — the validity period is ignored without it.",
     f"MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1;"),
    ("SrsPreSinrJudgeThld",
     "SRS-based weights or SRS MU pairing already ON, otherwise this threshold has no effect.",
     f"LST NRDUCELLBEAMALGO: NrDuCellId={PID};   // WeightAlgoSwitch must show SRS_WEIGHT_ESTIMATE_SW"),
    ("SrsBlindIsDegree",
     "SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-1 first.",
     f"MOD NRDUCELLSRSMEAS: NrDuCellId={PID}, SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-1;"),
    ("AggLvlComprCceUsageThld",
     "PdcchAlgoSwitch=PDCCH_AGG_LVL_COMPR_SW-1 first. This line was missing from the Ph1 work order.",
     f"MOD NRDUCELLPDCCH: NrDuCellId={PID}, PdcchAlgoSwitch=PDCCH_AGG_LVL_COMPR_SW-1;"),
    ("DlSchOptTimeThld",
     "DlSchAlgoSwitch=RES_BASED_DL_ADAPT_SCH_SW-1 first.",
     f"MOD NRDUCELLDLSCH: NrDuCellId={PID}, DlSchAlgoSwitch=RES_BASED_DL_ADAPT_SCH_SW-1;"),
    ("DlAdaptSchTimeThld",
     "DlSchAlgoSwitch=RES_BASED_DL_ADAPT_SCH_SW-1 first.",
     f"MOD NRDUCELLDLSCH: NrDuCellId={PID}, DlSchAlgoSwitch=RES_BASED_DL_ADAPT_SCH_SW-1;"),
    ("DlPrecodeOptOnSrsDtxSw",
     "SRS-based precoding path already ON (WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1).",
     f"LST NRDUCELLBEAMALGO: NrDuCellId={PID};   // confirm the SRS weight path is active"),
    ("CoverageScenario",
     "License NR0SBSC3DC00 and the RF-planned scenario per TRP. Tilt / Azimuth must be real "
     "values — never leave 255 on a cell that takes beam-shaping switches.",
     f"LST NRDUCELLTRPBEAM: NrDuCellTrpId={TID};   // read the live scenario and tilt first"),
    ("Tilt=255",
     "RF design values for tilt and azimuth. A cell left at Tilt=255 must be fixed before it takes "
     "any beam-shaping or SSB-adaptation switch, and must not join the trial set.",
     f"LST NRDUCELLTRPBEAM: NrDuCellTrpId={TID};   // fix RF first, then re-send this line"),
    ("FR1MaxCellCsirsPortNum",
     "8-port-capable AAU (32T is). CSI-RS port count affects every UE in the cell, so align it with "
     "MaxMimoLayerNum and the DL layer licence before changing it.",
     f"LST NRDUCELLPDSCH: NrDuCellId={PID};   // check MaxMimoLayerNum and licence first"),
    ("MaxPairLayerNum",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 — pairing layer caps do nothing while MU is OFF.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlMuMimoGroupMode",
     "MuMimoSwitch=DL_MU_MIMO_SW-1.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlPmiMuMimoSpaceIsoThld",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 (PMI pairing branch).",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlSrsMuMimoSpaceIsoThld",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 and the SRS weight path ON (SRS pairing branch).",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlMuMimoSrsPreSinrThld",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 and SRS-based pairing in use.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlMuBackToSuSeThld",
     "MuMimoSwitch=DL_MU_MIMO_SW-1.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlMuPmiBeamNumThld",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 (PMI pairing branch).",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlMuMimoSirScaleFactor",
     "MuMimoSwitch=DL_MU_MIMO_SW-1.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlMuEstRbPolicy",
     "MuMimoSwitch=DL_MU_MIMO_SW-1.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlSrsMuMimoRank",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 and the SRS weight path ON.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("DlPmiMuMimoRank",
     "MuMimoSwitch=DL_MU_MIMO_SW-1 (PMI pairing branch).",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=DL_MU_MIMO_SW-1;"),
    ("UlMuMimoCorrThld",
     "MuMimoSwitch=UL_MU_MIMO_SW-1.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=UL_MU_MIMO_SW-1;"),
    ("UlMuMimoSinrThld",
     "MuMimoSwitch=UL_MU_MIMO_SW-1.",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=UL_MU_MIMO_SW-1;"),
]

PREREQ_NONE = ("None — parent MO only (no dependent switch).", "—")


def prereq_for(cmd):
    """Pre-requisite switch + the MML/LST to run before this line."""
    for key, text, mml in sorted(PREREQ, key=lambda x: len(x[0]), reverse=True):
        if key in cmd:
            return text, mml
    return PREREQ_NONE


# Section 2: dump vs each suggestion (32T n=564 unless noted)
DUMP_ROWS = [
    # sid, family, param, live, status, enabled, gap, action, license
    ("S15-01", "Step1", "DL_PMI_SRS_ADAPT_SW", "ON 564/564", "Already ON", "Yes", "Matches commercial.", "Skip — do not re-send", "FBFD-010003"),
    ("S15-01", "Step1", "BeamPerceiveMode=DISTRIBUTED_MODE", "DISTRIBUTED_MODE 564/564", "Already ON", "Yes", "Prerequisite of weight opt is live.", "Skip", "FBFD-010003"),
    ("S15-01", "Step1", "FR1MaxCellCsirsPortNum=8PORT / TYPE0 / FD_RESOURCE", "8PORT+TYPE0+FD ON", "Already ON", "Yes", "CSI baseline live.", "Skip", "AHR/CSI"),
    ("S15-01", "Step1", "SrsWeightValidityPeriod=MS400", "See LST NRDUCELLPDSCHPRECODE", "Check", "See dump", "Timer, not a missing switch.", "LST only", "FBFD-010003"),
    ("S15-01", "Step1", "SrsNonASFixedWeightType=PMI_WEIGHT", "Needs ADAPT=ON (already ON)", "Already ON", "Yes", "Non-AS PMI when SRS missing.", "Skip", "FBFD-010003"),
    ("S15-02", "Step1", "SRS_WEIGHT_ESTIMATE_SW", "OFF 564/564", "Missing", "No", "Largest basic-MIMO leftover.", "Enable (CR01 30/30)", "FBFD-010003"),
    ("S15-02", "Step1", "PMI_WEIGHT_OPT_SW", "OFF 564/564", "Hold", "No", "Missing vs FPD but mix with SRS tonight.", "Hold first night", "FBFD-010003"),
    ("S15-02", "Step1", "OPEN_LOOP_WEIGHT_OPT_SW", "OFF 564/564", "Hold", "No", "Same mix rule.", "Hold first night", "FBFD-010003"),
    ("S15-03", "Step1", "SRS_SINR_MEAS_OPT_SW", "OFF 564/564", "Missing", "No", "FPD example ON.", "Enable (CR01 30/30)", "FBFD-010003"),
    ("S15-03", "Step1", "UL_RANK_FAST_DECREASE_SW", "OFF 564/564", "Missing", "No", "Protects UL BLER.", "Enable (CR01 30/30)", "FBFD-010003"),
    ("S15-03", "Step1", "PUSCH_CE_SINR_LEVEL_ENH_SW", "OFF 564/564", "Missing", "No", "PUSCH CE enhance.", "Enable (CR01 30/30)", "FBFD-010003"),
    ("S15-03", "Step1", "SrsPreSinrJudgeThld / CHN_MEASURE_CPU_DEC", "Not a sheet-14 Major", "Check", "See dump", "Threshold / gNB CPU bit — LST.", "LST only", "FBFD-010003"),
    ("S15-04", "Step1", "DlSchOptTimeThld / DlAdaptSchTimeThld", "Not a sheet-14 Major", "Check", "See dump", "Timers, not a missing switch pack.", "LST only if tput stuck", "—"),
    ("S15-05", "Step2", "MaxMimoLayerNum / DL_RANK_ADAPT / SU multi-layer", "LAYER_16 557/564; SU ON 564; rank adapt ON", "Already ON", "Yes", "Better than FPD LAYER_8 sample.", "Skip — never send LAYER_8", "FOFD-010020"),
    ("S15-06", "Step2", "SU_DMRS_OH_ADAPT / SRS_PRECODE_OPT / JT", "Not the DL-tput gap", "Check", "Partial", "Optional SU helpers; JT only if intended.", "Do not send JT cluster-wide", "FOFD-010020"),
    ("S15-07", "Step3", "DL_MU / UL_MU / PDCCH_MU + isolation −50 / group ISOLATION_CORRELATION", "ON 563/564; PreSINR −50; BackToSu 5", "Already ON", "Yes", "Core MU live.", "Skip on 32T; keep OFF on 2T2R", "FOFD-010010"),
    ("S15-08", "Step3", "HEAVY_LOAD_SCH_PRI_OPT + MU ranks / DMRS policy", "HEAVY_LOAD ON 563; EPF ON", "Already ON", "Yes", "Scheduler baseline live.", "Skip these; extras in S15-16/25", "—"),
    ("S15-09", "Step4", "MMIMO_MULTILAYER_ENHANCE / PAIRING_PREFERRED / MU_RANK_BOOSTING / SRS_MEAS_ACCELERATING", "All OFF 564/564", "Hold", "No", "Quota LAYER_16 unused.", "Next CR after iBeam 1.0 green", "NR0S0DLEPU00"),
    ("S15-10", "Step4", "UL MULTILAYER_DEMOD_ENH / FLEX_PAIR / RES_BASED_ULSCH", "OFF 564/564", "Hold", "No", "UL multilayer not started.", "After DL multilayer", "NR0S0ULEPU00"),
    ("S15-11", "Step5", "DL_INITIAL_BEAM_SELECT / SSB density adapt / CoverageScenario", "Initial-select ON; density adapt ON; SCENARIO_8 mix", "Already ON", "Yes (select)", "Do not copy FPD Tilt=255.", "Skip select; Fix tilt=255 cluster-wide (28 cells)", "FBFD-010015"),
    ("S15-11", "Step5", "Tilt / Azimuth (RF)", "Tilt 255 on 28 cells; Azimuth 0° on 580/581", "Fix", "No (RF)", "Beams not applied if tilt=255.", "P0 Fix before switches. CR01 tilts OK (2–9°).", "RF"),
    ("S15-12", "Step5", "BEAM_TRACKING + INTELLIGENT_BEAM_SELECTION", "OFF 564/564", "Missing", "No", "Connected-mode BF does not follow UE.", "Enable (CR01 30/30)", "FBFD-010015"),
    ("S15-12", "Step5", "SSB_BEAM_ADAPT + SSB_BEAM_VERTICAL_COV_IMP", "OFF 564/564", "Missing", "No", "SSB frozen after initial select.", "Enable (CR01 30/30)", "FOFD-010100"),
    ("S15-12", "Step5", "BEAM_SELECT_OPT_SW", "OFF 564/564", "Missing", "No", "Also iBeam 1.0 child.", "Enable with iBeam pack (CR01)", "FOFD-081201"),
    ("S15-13", "Step6", "AHR_PHASE1 + FD/TYPE0/8PORT + CSIRS_INTRF_STATIC_AVOID", "Phase1 ON 563/564", "Already ON", "Yes", "AHR introduction live.", "Skip", "FOFD-051301"),
    ("S15-14", "Step6", "AHR_EXP_TURBO_PHASE2 master", "ON 563/564", "Already ON", "Yes", "Turbo master live.", "Skip master", "FOFD-061201"),
    ("S15-14", "Step6", "SRS_IC_SW", "OFF 564/564", "Hold", "No", "Turbo child off. Tight MUX is in CR01.", "Hold this night (do not stack with CR01 tight MUX)", "FOFD-061201"),
    ("S15-14", "Step6", "SRS_JOINT_PC_SW", "OFF 564/564 (CR01 Proposed ON 30)", "Partial", "No (cluster) / Yes (CR01)", "Joint PC without IC.", "Enable on trial (CR01); not cluster yet", "FOFD-061201"),
    ("S15-15", "Step6", "AHR_CAPC_UPGRADE_PHASE2 + PDCCH_MULTI_DIM_JOINT_SCH", "OFF 564/564", "Hold", "No", "Capacity wave never started.", "After Turbo SRS-IC + iBeam 1.0 green", "NR0S00ACT200"),
    ("S15-16", "Step7", "HighPrecisionBeamSwitch", "OFF 564/564 (CR01 ON 30)", "Missing", "No (cluster)", "Main missing DL-interference master.", "Enable (CR01 10 sites). LST NR0S00BEAM00", "FOFD-081201"),
    ("S15-16", "Step7", "iBeam 1.0 children (tight MUX, precise/anti-intrf MU, tail MCS, res-based, agg-compress, blind IS, BWP hybrid IR, RLC merge)", "All OFF 564/564", "Missing", "No", "Commercial 32T runs the package.", "Enable as one pack (CR01)", "FOFD-081201"),
    ("S15-17", "Step7", "HighPrecisionBeamPhase2Sw / DL_ROBUST_WEIGHT", "OFF 564/564", "Hold", "No", "Correctly not started — 1.0 missing on cluster.", "Do not enable now", "FOFD-091200"),
    ("S15-18", "Step7", "HighPrecisionBeamPhase3Sw / DL_SELF_FUSION_WEIGHT", "OFF", "Hold", "No", "Last wave.", "After 2.0 green", "FOFD-100200"),
    ("S15-19", "Step8", "UL_LOW_NOISE_SW + UL_MU_GRP_PAIR + DIFF_WAVEFORM + CORR_ACCEL", "OFF 564/564 (CR01 ON 30)", "Missing", "No (cluster)", "UL Boosting 1.0 never started cluster-wide.", "Enable (CR01). LST NR0S00UAHR00", "FOFD-091201"),
    ("S15-20", "Step8", "UL_LOW_NOISE_PHASE2_SW", "OFF 564/564", "Hold", "No", "After 1.0 KPI-green + sync.", "Not this CR", "FOFD-100201"),
    ("S15-21", "Step9", "DM_MIMO_SERVICE_SWITCH", "OFF; n41 single-TRP", "N/A", "N/A", "No DAS.", "Keep OFF", "N/A"),
    ("S15-22", "Step9", "Fusion INTRA_CELL_MIMO", "No Fusion cluster", "N/A", "N/A", "No virtual 128T.", "Keep OFF", "N/A"),
    ("S15-23", "Step10", "FR2 mmWave MOs / VOL_BASED_BEAM_MULTIPLEX", "Empty / n41 FR1 40 MHz", "N/A", "N/A", "Not FR2.", "Keep OFF", "N/A"),
    ("S15-24", "Step11", "STR ANTENNAPORTOPTDET", "32T AAU (not 4T4R NORMAL_CELL)", "N/A", "N/A", "Feature is 4T4R commissioning.", "Skip on 32T trial", "FBFD-010025"),
    ("S15-25", "Step3", "FREQ_SEL_SCH / LOAD_BASED_DL_EXP_SCH / SMART_SCH / HIGH_CAPACITY_EXP", "All OFF 564/564", "Hold", "No", "Beyond iBeam 1.0 minimum.", "After CSI/iBeam 1.0 green; license+TAC for smart/high-cap", "Performance pack"),
]

# Extra keywords so each MML line hits the right dump row (param text is often a summary).
PARAM_ALIASES = {
    "iBeam 1.0 children (tight MUX, precise/anti-intrf MU, tail MCS, res-based, agg-compress, blind IS, BWP hybrid IR, RLC merge)": (
        "SRS_BLIND_IS_MEAS", "SRS_TIGHT_MULTIPLEXING", "DL_BWP_HYBRID_INTRF_RANDOM",
        "DL_RLC_STAT_RPT_MERGE", "PDCCH_AGG_LVL_COMPR", "BEAM_SELECT_OPT",
        "DL_MU_PRECISE_SCH", "DL_MU_ANTI_INTRF", "TAIL_PKT_MCS_OPT", "RES_BASED_DL_ADAPT_SCH",
    ),
    "Tilt / Azimuth (RF)": ("Tilt", "Azimuth"),
    "DL_INITIAL_BEAM_SELECT / SSB density adapt / CoverageScenario": (
        "DL_INITIAL_BEAM_SELECT", "CoverageScenario", "SsbPeriod", "CsiPeriod",
        "PDCCH_BEAM_ROBUST", "PdcchBeamRobDtxThld", "SRS_BEAM_SELECT_OPT", "PdcchPrecodeEnhPolicy",
    ),
    "AHR_PHASE1 + FD/TYPE0/8PORT + CSIRS_INTRF_STATIC_AVOID": (
        "AHR_PHASE1", "CSIRS_INTRF_STATIC_AVOID", "FD_RESOURCE", "TYPE0",
        "EXP_BASED_MM_ADAPT", "RES_BASED_MM_ADAPT", "FixedAmcStepValue", "DlInitialMcsAdjValue",
        "DlDelaySchBufferThld",
    ),
    "AHR_EXP_TURBO_PHASE2 master": (
        "AHR_EXP_TURBO_PHASE2", "TAIL_PKT_SCH_OPT", "DL_MCS_ADJ_OPT", "RANK_AND_SINR_ESTIMATE_OPT",
        "SrsIntrfThld", "IntrfUeSrsPcMinSinrTarget", "MaxSrsPoAdjustAmount",
    ),
    "UL_LOW_NOISE_SW + UL_MU_GRP_PAIR + DIFF_WAVEFORM + CORR_ACCEL": (
        "UL_LOW_NOISE_SW", "UL_MU_GRP_PAIR", "DIFF_WAVEFORM_PAIR", "UL_CORR_ACCELERATION",
        "PUSCH_RES_ADAPT_ALLOC", "UL_PREALLOCATION_PERIOD_ADJ", "UL_CELL_OLLA",
        "SMALL_PKT_OL_ADAPT", "SR_BASED_SCH_MCS_OPT", "SinrThldforWaveformSel",
        "PDCCH_SYM_SMART_ALLOC", "CCE_SYMBOL_ALLOC_OPT", "F1_ACK_CODE_CHN_INTRF",
        "IRC_BASED_PDP_DETECT",
    ),
    "HighPrecisionBeamPhase2Sw / DL_ROBUST_WEIGHT": (
        "HighPrecisionBeamPhase2", "DL_ROBUST_WEIGHT", "RobustWtPhaseCalcMethod",
        "DYNAMIC_CLUSTER_GROUP", "PRECISE_MUMIMO_EVAL", "FAR_UE_RANK_OPT",
        "DL_CORR_ACCELERATION", "INTER_CELL_INTRF_AVOID", "MuMimoIblerTarget",
    ),
    "HighPrecisionBeamPhase3Sw / DL_SELF_FUSION_WEIGHT": (
        "HighPrecisionBeamPhase3", "DL_SELF_FUSION_WEIGHT", "SelfFusionWt",
        "PDCCH_ROBUST_WEIGHT", "DL_SMART_AMC",
    ),
    "UL_LOW_NOISE_PHASE2_SW": (
        "UL_LOW_NOISE_PHASE2", "PUSCH_COORD_PWR_CTRL", "INTRF_SC_LINK_PERF_PC",
        "UlCpcUeSsbRsrpThld", "UL_RETRANS_PREC_RB", "UlFirstRetransMinRbPct",
        "UL_PRECISE_MCS_OPT", "MULTI_BEAM_RX_ENH", "INTRF_PREC_FREQ_OFS", "PREC_CHANNEL_EST",
    ),
    "MMIMO_MULTILAYER_ENHANCE / PAIRING_PREFERRED / MU_RANK_BOOSTING / SRS_MEAS_ACCELERATING": (
        "MMIMO_MULTILAYER_ENHANCE", "MU_RANK_BOOSTING", "SRS_BLIND_IS_SW", "SrsBlindIsDegree",
        "MU_MIMO_PAIRING_PREFERRED", "SRS_MEAS_ACCELERATING", "DL_HYBRID_PRECODING",
        "TAIL_PKT_MCS_OPT", "SmallPktType1RobustSchPol", "DlMuBackToSuSeThld",
        "PUCCH_INTRF_COORD", "CCE_RESOURCE_OPT", "UE_BWP0_PDSCH_RES_OPT", "PDCCH_BLIND_DET",
    ),
    "UL MULTILAYER_DEMOD_ENH / FLEX_PAIR / RES_BASED_ULSCH": (
        "MULTILAYER_DEMOD_ENH", "MU_MIMO_FLEX_PAIR", "RES_BASED_ADAPT_ULSCH",
        "LATENCY_BASED_ADAPT_ULSCH", "UL_SU_SINR_INTEL_PREDICT",
    ),
    "AHR_CAPC_UPGRADE_PHASE2 + PDCCH_MULTI_DIM_JOINT_SCH": (
        "AHR_CAPC_UPGRADE_PHASE2", "MuMimoOptSwith", "DlMuGatherOptSw",
        "PDCCH_MULTI_DIM_JOINT_SCH", "DlMuIRPrecodePol", "DlSrsMuMimoPreSinrThld",
    ),
    "DL_MU / UL_MU / PDCCH_MU + isolation −50 / group ISOLATION_CORRELATION": (
        "UL_MU_MIMO", "DL_MU_MIMO", "PDCCH_MU", "DlPmiMuMimoSpaceIsoThld",
        "DlSrsMuMimoSpaceIsoThld", "DlMuMimoSrsPreSinrThld", "ISOLATION_CORRELATION",
        "UlMuMimoCorrThld", "UlMuMimoSinrThld", "DlMuBackToSuSeThld",
    ),
    "HEAVY_LOAD_SCH_PRI_OPT + MU ranks / DMRS policy": (
        "HEAVY_LOAD_SCH_PRI_OPT", "MaxPairLayerNum", "DlMuMimoSirScaleFactor",
        "DlMuPmiBeamNumThld", "DlMuEstRbPolicy", "DlSrsMuMimoRank", "DlPmiMuMimoRank",
        "PrecodingIntrfSupprValue",
    ),
    "MaxMimoLayerNum / DL_RANK_ADAPT / SU multi-layer": (
        "MaxMimoLayerCnt", "MaxMimoLayerNum", "DL_RANK_ADAPT", "SuMimoPwrCtrlProtectThld",
    ),
    "SU_DMRS_OH_ADAPT / SRS_PRECODE_OPT / JT": (
        "SU_DMRS_OH_ADAPT", "SRS_PRECODE_OPT", "INTRA_GNB_DL_JT",
    ),
    "DM_MIMO_SERVICE_SWITCH": ("DM_MIMO", "TxRxMode", "TrpType"),
    "Fusion INTRA_CELL_MIMO": ("INTRA_CELL_MIMO", "MUMIMO_SINR_ENH", "FUSION_CALIB", "SINGLE_TRP_SSB"),
    "FR2 mmWave MOs / VOL_BASED_BEAM_MULTIPLEX": (
        "NRDUCELLTRPMMWAVBEAM", "VOL_BASED_BEAM_MULTIPLEX", "FLEXIBLE_DENSE_BEAM",
        "DYNAMIC_BEAM_ALLOC", "SHORT_TAC_PERIOD", "SSB_MEAS_POS_POLICY",
    ),
    "STR ANTENNAPORTOPTDET": ("ANTENNAPORTOPTDET",),
    "FREQ_SEL_SCH / LOAD_BASED_DL_EXP_SCH / SMART_SCH / HIGH_CAPACITY_EXP": (
        "FREQ_SEL_SCH", "LOAD_BASED_DL_EXP_SCH", "SMART_SCH_AND_LINK_ADAPT", "HIGH_CAPACITY_EXP_IMP",
    ),
    "BEAM_TRACKING + INTELLIGENT_BEAM_SELECTION": ("BEAM_TRACKING", "INTELLIGENT_BEAM_SELECTION"),
    "SSB_BEAM_ADAPT + SSB_BEAM_VERTICAL_COV_IMP": ("SSB_BEAM_ADAPT", "SSB_BEAM_VERTICAL_COV"),
    "SrsPreSinrJudgeThld / CHN_MEASURE_CPU_DEC": ("SrsPreSinrJudgeThld", "CHN_MEASURE_CPU_DEC"),
}

TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]{4,}")
DUMP_NONE = ("", "", "See LST", "Check", "See dump", "", "", "LST to confirm", "")


def _keys_for_param(param):
    keys = list(TOKEN_RE.findall(param)) + list(PARAM_ALIASES.get(param, ()))
    for part in re.split(r"[^A-Za-z0-9_]+", param):
        if len(part) >= 5:
            keys.append(part)
    return keys


def dump_for(sid, cmd):
    """Map one MML command to the dump row for that suggestion (already enabled or not)."""
    if "Tilt=255" in cmd:
        rows_sid = [row for row in DUMP_ROWS if row[0] == sid]
        if rows_sid and rows_sid[0][4] == "N/A":
            return rows_sid[0]
        same = [row for row in rows_sid if row[4] == "Fix"]
        anyfix = [row for row in DUMP_ROWS if row[4] == "Fix"]
        if same:
            return same[0]
        if anyfix:
            return anyfix[0]
    rows = [r for r in DUMP_ROWS if r[0] == sid]
    scored = []
    for row in rows:
        hitlen = 0
        for k in _keys_for_param(row[2]):
            if k and k in cmd:
                hitlen = max(hitlen, len(k))
        if hitlen:
            scored.append((hitlen, row))
    if scored:
        scored.sort(key=lambda x: -x[0])
        return scored[0][1]
    if len(rows) == 1:
        return rows[0]
    return rows[0] if rows else DUMP_NONE


def add_box(ws, r, rec, sheet_name):
    start = r
    fam = rec["family"]
    color = FAMILY_COLOR.get(fam, NAVY)
    merge(ws, r, 1, r, COLS - 3)
    put(ws, r, 1, f"  {rec['sid']}   ·   {rec['title']}", size=12, bold=True, color=WHITE,
        fill_hex=color, h="left", v="center")
    for c in range(2, COLS - 2):
        ws.cell(r, c).fill = fill(color)
    put(ws, r, COLS - 2, "Read benefit →", size=9, bold=True, color=WHITE, fill_hex=color,
        h="right", v="center")
    merge(ws, r, COLS - 1, r, COLS)
    href_sheet(ws.cell(r, COLS - 1), rec["jump"], rec["jump"])
    ws.cell(r, COLS - 1).fill = fill("FFF2CC")
    ws.cell(r, COLS).fill = fill("FFF2CC")
    ws.row_dimensions[r].height = 24
    r += 1
    merge(ws, r, 1, r, COLS)
    put(ws, r, 1, f"Document: {rec['doc']}    |    Replace {PID} / {TID} with live NR DU cell / TRP IDs.",
        size=9, italic=True, fill_hex=PALE_BLUE, h="left", v="center")
    for c in range(2, COLS + 1):
        ws.cell(r, c).fill = fill(PALE_BLUE)
    ws.row_dimensions[r].height = 18
    r += 1
    r = label_row(ws, r, "Principal", rec["principal"], PALE_GOLD)
    r = label_row(ws, r, "Benefit", rec["benefit"], PALE_GREEN)
    r = label_row(ws, r, "Parameter details", rec["params"], "DDEBF7")

    # MML header: Seq | command | [pre-req] | Live dump | status | Enabled? | Action | Counter | KPI | Notes | Jump
    for col in range(1, COLS + 1):
        put(ws, r, col, "", size=8, bold=True, fill_hex=YELLOW_HDR, h="center", v="center", border=True)
    put(ws, r, COL["sn"], "Seq  /  MML #" if WITH_PREREQ else "MML #",
        size=8, bold=True, fill_hex=BLUE_HDR, h="center", v="center", border=True)
    merge(ws, r, COL["cmd"], r, COL["cmd_end"])
    put(ws, r, COL["cmd"], "MML Command  (one parameter / one switch / one line)  +  note at end",
        size=8, bold=True, fill_hex=YELLOW_HDR, h="center", v="center", border=True)
    if WITH_PREREQ:
        put(ws, r, COL["prereq"], "Pre-requisite switch / parameter  (must be ON first)",
            size=8, bold=True, color=WHITE, fill_hex=RED_HDR, h="center", v="center", border=True)
        put(ws, r, COL["prereq_mml"], "Pre-requisite MML  (run BEFORE this line)",
            size=8, bold=True, color=WHITE, fill_hex=RED_HDR, h="center", v="center", border=True)
    put(ws, r, COL["live"], "Live DHK dump (32T)", size=8, bold=True, color=WHITE, fill_hex=NAVY, h="center", v="center", border=True)
    put(ws, r, COL["status"], "Dump status", size=8, bold=True, color=WHITE, fill_hex=NAVY, h="center", v="center", border=True)
    put(ws, r, COL["enabled"], "Enabled?", size=8, bold=True, color=WHITE, fill_hex=NAVY, h="center", v="center", border=True)
    put(ws, r, COL["action"], "Action", size=8, bold=True, color=WHITE, fill_hex=NAVY, h="center", v="center", border=True)
    put(ws, r, COL["ctr"], "Counter monitor", size=8, bold=True, color=WHITE, fill_hex=TEAL_HDR, h="center", v="center", border=True)
    put(ws, r, COL["kpi"], "Impact on KPI", size=8, bold=True, fill_hex=GOLD_HDR, h="center", v="center", border=True)
    merge(ws, r, COL["notes"], r, COL["notes_end"])
    put(ws, r, COL["notes"], "short Notes", size=8, bold=True, fill_hex="F4B183", h="center", v="center", border=True)
    ws.cell(r, COL["notes_end"]).fill = fill("F4B183")
    merge(ws, r, COL["jump"], r, COL["jump_end"])
    put(ws, r, COL["jump"], "Jump to Basic", size=8, bold=True, fill_hex=YELLOW_HDR, h="center", v="center", border=True)
    ws.cell(r, COL["jump_end"]).fill = fill(YELLOW_HDR)
    ws.row_dimensions[r].height = 28
    r += 1

    for i, (cmd, note) in enumerate(rec["mmls"], 1):
        line = mml_line(cmd, note)
        ctr, kpi, short = extras(cmd, note)
        drow = dump_for(rec["sid"], cmd)
        live, status, enabled, action = drow[3], drow[4], drow[5], drow[7]
        st_fill = ST_FILL.get(status, WHITE)
        fh = WHITE if i % 2 else ROW_ALT
        put(ws, r, COL["sn"], i, size=9, bold=True, fill_hex=fh, h="center", v="top", border=True)
        merge(ws, r, COL["cmd"], r, COL["cmd_end"])
        put(ws, r, COL["cmd"], line, size=8, fill_hex=fh, h="left", v="top", border=True)
        for c in range(COL["cmd"] + 1, COL["cmd_end"] + 1):
            ws.cell(r, c).fill = fill(fh)
            ws.cell(r, c).border = thin
        if WITH_PREREQ:
            pre_txt, pre_mml = prereq_for(cmd)
            none_pre = pre_txt == PREREQ_NONE[0]
            pf = "EAECEE" if none_pre else "FCE4E4"
            put(ws, r, COL["prereq"], pre_txt, size=8, bold=not none_pre,
                fill_hex=pf, h="left", v="top", border=True)
            put(ws, r, COL["prereq_mml"], pre_mml, size=8, fill_hex=pf, h="left", v="top", border=True)
        put(ws, r, COL["live"], live, size=8, fill_hex=st_fill, h="left", v="top", border=True)
        put(ws, r, COL["status"], status, size=8, bold=True, fill_hex=st_fill, h="center", v="top", border=True)
        put(ws, r, COL["enabled"], enabled, size=8, bold=True, fill_hex=st_fill, h="center", v="top", border=True)
        put(ws, r, COL["action"], action, size=8, fill_hex=st_fill, h="left", v="top", border=True)
        put(ws, r, COL["ctr"], ctr, size=8, fill_hex="D5F5E3", h="left", v="top", border=True)
        put(ws, r, COL["kpi"], kpi, size=8, fill_hex="FFF2CC", h="left", v="top", border=True)
        merge(ws, r, COL["notes"], r, COL["notes_end"])
        put(ws, r, COL["notes"], short, size=8, fill_hex="FDEBD0", h="left", v="top", border=True)
        ws.cell(r, COL["notes_end"]).fill = fill("FDEBD0")
        ws.cell(r, COL["notes_end"]).border = thin
        merge(ws, r, COL["jump"], r, COL["jump_end"])
        href_sheet(ws.cell(r, COL["jump"]), rec["jump"], rec["jump"])
        ws.cell(r, COL["jump"]).fill = fill("FFF2CC")
        ws.cell(r, COL["jump_end"]).fill = fill("FFF2CC")
        ws.row_dimensions[r].height = min(72, max(28, 16 + len(line) // 100 * 12
                                                 + (12 if WITH_PREREQ else 0)))
        r += 1
    end = r - 1
    box_border(ws, start, end)
    r = blank(ws, r, 10)
    return r, start


def patch_cover(wb):
    ws = wb["0. Cover & Index"]
    ws["A1"].value = f"  5G MIMO (all features together)  —  Deployment Workbook  {VERSION}"
    r = ws.max_row + 2
    if WITH_PREREQ:
        r = section(ws, r, 10,
                    f"{VERSION} — pre-requisite columns on every MML + Master Findings + phased Action Plan")
        r = note_bar(ws, r, 10,
                     f"{SHEET_NAME} Section 1: every MML now carries “Pre-requisite switch / parameter (must be ON "
                     "first)” and “Pre-requisite MML (run BEFORE this line)”, so the sequence is on the row itself. "
                     "This is the fix for the 10-Sep RETCODE 2147616329 rejection (BEAM_TRACKING_SW needs "
                     "DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1 first). New sheet 16 = Master Findings of the Ph1 trial "
                     f"(10–11 Sep 2026). New sheet 17 = Action Plan Ph1…Ph7. File: MIMO_Deployment_{VERSION}.xlsx")
    else:
        r = section(ws, r, 10, "v6.0 — dump status sits on each suggestion MML row (no separate incon section)")
        r = note_bar(ws, r, 10,
                     "Former “14. MIMO Incon Report” and “15. MIMO Suggestions from Doc” are one sheet: "
                     f"{SHEET_NAME}. Section 1 = FPD boxes; each MML has Live dump / Dump status / Enabled? / Action "
                     "plus Counter monitor / Impact on KPI / short Notes / Jump to Basic. "
                     "Section 2 = Enable/Fix trial list (still OFF). Section 3 = Performance counter and Monitoring KPI. "
                     "CR01 is sheet 15. File: MIMO_Deployment_v6.0.xlsx")
    r = headers(ws, r, ["#", "Sheet", "Maps to", "What you will find"] + [""] * 6)
    r = table_row(ws, r,
                  ["14", SHEET_NAME, "FPD suggestions + DHK dump 8 Sep 2026",
                   "Boxes with dump-on-MML + Enable/Fix proposal + counters"] + [""] * 6,
                  fills=[PALE_ORANGE] * 10, height=34)
    merge(ws, r - 1, 4, r - 1, 10)
    r = table_row(ws, r,
                  ["15", CR01_NEW, "CR01 10-site CME pack",
                   "Same 10 gNB execution pack as v4.0 (was sheet 16)"] + [""] * 6,
                  fills=[PALE_GOLD] * 10, height=30)
    merge(ws, r - 1, 4, r - 1, 10)
    if WITH_PREREQ:
        r = table_row(ws, r,
                      ["16", FINDINGS_SHEET, "Ph1 trial result 10–11 Sep 2026",
                       "What was executed, what was rejected, why DL tput flat and MU pairing fell, verdict per switch"]
                      + [""] * 6, fills=["F8CBAD"] * 10, height=34)
        merge(ws, r - 1, 4, r - 1, 10)
        r = table_row(ws, r,
                      ["17", ACTION_SHEET, "Phased rollout Ph1…Ph7",
                       "Ph1 complete; Ph2 fix + re-run the rejected MML in sequence; Ph3…Ph7 with KPI exit gates"]
                      + [""] * 6, fills=[PALE_GREEN] * 10, height=34)
        merge(ws, r - 1, 4, r - 1, 10)
    return r


def retarget_hyperlinks(wb, old_sheet, new_sheet):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                hl = cell.hyperlink
                if not hl:
                    continue
                loc = getattr(hl, "location", None) or ""
                if old_sheet in loc:
                    cell.hyperlink = Hyperlink(ref=cell.coordinate,
                                               location=loc.replace(old_sheet, new_sheet),
                                               display=str(cell.value or new_sheet))
                if cell.value and old_sheet in str(cell.value):
                    cell.value = str(cell.value).replace(old_sheet, new_sheet)


def build_combined(wb):
    ws = wb.create_sheet(SHEET_NAME, 14)
    setup_sheet(ws, SHEET_NAME)
    set_widths(ws, WIDTHS)
    ws.oddHeader.left.text = "5G MIMO Suggestions + Incon (v6.0) — dump status on each MML"
    ws.oddFooter.left.text = "Dump: 5G CME 8 Sep 2026 DHK 564×32T32R n41 · Suggestions from RAN10.1 FPDs · Jump to Basic = step sheet"
    ws.freeze_panes = "A4"
    ws.sheet_properties.tabColor = "C65911"

    items = suggestions()
    r = 1
    r = banner(ws, r, COLS,
               "  14.  MIMO Suggestions + Inconsistency  —  dump status on each MML, then trial proposal",
               fill_hex="C65911", size=16, height=30)
    r = note_bar(ws, r, COLS,
                 "One sheet instead of old 14 + 15. Section 1 keeps the FPD suggestion boxes. After each MML Command: "
                 "Live DHK dump / Dump status / Enabled? / Action (the old Section 2 check), then Counter monitor / "
                 "Impact on KPI / short Notes; Jump is Jump to Basic. There is no separate dump-incon section. "
                 "Section 2 is the Enable/Fix list (still OFF). Section 3 is the monitoring pack. "
                 "Replace {NrDuCellId} / {NrDuCellTrpId}. Keep 2T2R indoor OFF. mmWave/DAS/Fusion = N/A on n41. "
                 "CR01 (sheet 15) already sends most Section 2 Enables on 10 macros.")

    toc_row = r
    r = section(ws, r, COLS, "Jump  (click)")
    jumps = [("Section 1 — Suggestions + dump on MML", None),
             ("Section 2 — Final proposal", None),
             ("Section 3 — Counters / KPI", None)]
    jump_cells = []
    for i, (lab, _) in enumerate(jumps, 1):
        put(ws, r, i, lab, size=9, bold=True, fill_hex=PALE_BLUE, h="center", v="center", border=True)
        jump_cells.append((r, i))
    ws.row_dimensions[r].height = 20
    r += 1
    r = blank(ws, r, 8)

    # ----- Section 1 -----
    sec1 = r
    r = section(ws, r, COLS, "Section 1.  Document suggestions  (Principal · Benefit · Parameter · MML + dump)")
    r = note_bar(ws, r, COLS,
                 "Same boxes as v3.0 sheet 15. Dump columns on each MML (CME 8 Sep 2026 DHK 564×32T32R): "
                 "Already ON = Skip. Missing = still OFF (goes to Section 2 Enable). Hold = OFF but not this night. "
                 "Fix = RF first. N/A = not this network. Then Counter monitor / Impact on KPI / short Notes / Jump to Basic.")
    r = section(ws, r, COLS, "Index of suggestion boxes  (click ID to jump down this sheet)")
    r = headers(ws, r, pad(["ID", "Family", "Suggestion (click ID)", "Source step (Jump to Basic)"]))
    toc_start = r
    for rec in items:
        put(ws, r, 1, rec["sid"], size=10, bold=True, fill_hex=PALE_GOLD, h="center", v="center", border=True)
        put(ws, r, 2, rec["family"], size=10, fill_hex=WHITE, h="center", v="center", border=True)
        merge(ws, r, 3, r, 11)
        put(ws, r, 3, rec["title"][:90], size=9, fill_hex=WHITE, h="left", v="center", border=True)
        merge(ws, r, 12, r, 14)
        href_sheet(ws.cell(r, 12), rec["jump"], rec["jump"])
        ws.row_dimensions[r].height = 20
        r += 1
    r = blank(ws, r, 10)

    box_rows = {}
    r = subsection(ws, r, COLS, "Suggestion boxes", fill_hex=NAVY)
    for rec in items:
        r, start = add_box(ws, r, rec, SHEET_NAME)
        box_rows[rec["sid"]] = start
    for i, rec in enumerate(items):
        rr = toc_start + i
        href_row(ws.cell(rr, 1), SHEET_NAME, box_rows[rec["sid"]], rec["sid"])
        ws.cell(rr, 1).fill = fill(PALE_GOLD)
        href_row(ws.cell(rr, 3), SHEET_NAME, box_rows[rec["sid"]], rec["title"][:90])

    # ----- Section 2 (was 3) — final proposal -----
    sec2 = r
    r = section(ws, r, COLS, "Section 2.  Final proposal  —  trial / implement what is NOT enabled")
    r = note_bar(ws, r, COLS,
                 "Filtered from the dump on each Section 1 MML: Action = Enable or Fix only. This is the set still OFF vs a commercial "
                 "32T network with good DL user throughput. CR01 (sheet 15) already sends most Enable lines on 10 macros "
                 "× 3 cells (101/102/103). Do not send LAYER_8. Do not add SRS_IC the same night as SRS_TIGHT_MULTIPLEXING. "
                 "2T2R indoor DHTIAA1 / DHAPT11 / DHTEJ34 stay OFF. Yellow columns = MML through License.")
    titles = pad(["SN", "RAT", "Doc Name", "Action", "MML Command (Proposed)",
                  "MO Name", "Parameter ID", "Live Value (actual in Dump)",
                  "proposed Parameter Value", "Purpose/Short Notes", "License"])
    for i, t in enumerate(titles, 1):
        if not t:
            continue
        fh = BLUE_HDR if i <= 4 else YELLOW_HDR
        put(ws, r, i, t, size=8, bold=True, fill_hex=fh, h="center", v="center", border=True)
    ws.row_dimensions[r].height = 28
    r += 1
    mml_start = r
    sn = 1
    skip_enable = ("SRS_IC_SW",)  # CR01 tight MUX is in this proposal
    for rec in incon_mml_rows():
        if rec[3] not in ("Enable", "Fix"):
            continue
        cmd = str(rec[4] or "")
        if rec[3] == "Enable" and any(k in cmd for k in skip_enable):
            continue
        new = (sn,) + rec[1:]
        r = incon_mml_row(ws, r, new)
        for c in range(12, COLS + 1):
            ws.cell(r - 1, c).fill = fill(YELLOW_HDR if rec[3] == "Enable" else "F4B183")
            ws.cell(r - 1, c).border = thin
        sn += 1
    ws.auto_filter.ref = f"A{mml_start - 1}:K{r - 1}"

    r = blank(ws, r, 8)
    r = subsection(ws, r, COLS, "Still OFF — Hold this night (not in the trial MML above)", fill_hex="C65911")
    r = note_bar(ws, r, COLS,
                 "These are not enabled on the cluster either, but they are not the Section 2 send list. "
                 "SRS_IC stays Hold while CR01 SRS_TIGHT_MULTIPLEXING is ON. PMI/open-loop stay Hold while SRS weight is ON. "
                 "Dump status on the matching Section 1 MML is Hold.")
    r = headers(ws, r, pad(["SN", "Suggestion", "Parameter / switch", "Dump status", "Why Hold", "When"]))
    holds = [x for x in DUMP_ROWS if x[4] in ("Hold",)]
    for i, rec in enumerate(holds, 1):
        sid, fam, param, live, status, enabled, gap, action, lic = rec
        vals = [i, sid, param, status, gap, action]
        r = putn(ws, r, vals, fills=[ST_FILL["Hold"]] * COLS, bolds=[False, True, False, True],
                 center={1, 2, 4}, height=28)
        merge(ws, r - 1, 6, r - 1, COLS)
        if sid in box_rows:
            href_row(ws.cell(r - 1, 2), SHEET_NAME, box_rows[sid], sid)

    r = blank(ws, r, 8)
    r = bullets(ws, r, COLS, [
        "Cluster still OFF after CR01: same Enable list on the remaining 32T (not the 10 trial gNB).",
        "Hold (not in this proposal): PMI_WEIGHT_OPT / OPEN_LOOP_WEIGHT_OPT; SRS_IC while tight MUX is ON; multilayer; AHR CU 2.0; iBeam 2.0/3.0; freq-sel/smart sch; Boosting 2.0.",
        "Fix first if the cell is not a CR01 site: tilt=255 (28 cells cluster-wide) and azimuth vs RF design.",
        "Rollback: set the enabled bit back to 0 / HighPrecisionBeamSwitch=OFF / UL_LOW_NOISE_SW-0. Do not touch SU/MU/AHR Phase1+Turbo/LAYER_16.",
    ], fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 72

    # ----- Section 3 (was 4) -----
    sec3 = r
    r = section(ws, r, COLS, "Section 3.  Performance counter and Monitoring KPI", fill_hex=TEAL)
    r = note_bar(ws, r, COLS,
                 "MAE DU, 15 min, busy hour. Trial = CR01 10 gNB (or any cell where Section 2 was sent). "
                 "Control = neighbour 32T not in the trial. Pre D−7, post D+1 and D+7. "
                 "Source: FPD Counter Changes for MIMO TDD / Beam / AHR / iBeam / UL Boosting.")
    r = headers(ws, r,
                pad(["SN", "Counter ID", "Counter Name", "Function", "Use for Section 1 suggestions",
                     "KPI it feeds", "Granularity", "Pass / fail gate"]),
                fill_hex=TEAL, height=24)
    ctrs = [
        (1, "N.ThpVol.DL / N.RLC.ThpTime.DL.Cell", "User DL Average Throughput (DU)",
         "North-star commercial KPI for every DL box (S15-01..16, 25).",
         "Must not drop vs control. Target: lift vs pre-CR week.",
         "User DL Average Throughput (DU)", "15 min / cell",
         "Fail if trial < control trend."),
        (2, "N.ThpVol.UL / N.RLC.ThpTime.UL.Cell", "User UL Average Throughput (DU)",
         "UL boxes S15-03 / 10 / 19 / 20 and SRS quality that feeds DL weights.",
         "UL Boosting + PUSCH CE + UL rank.",
         "User UL Average Throughput (DU)", "15 min / cell",
         "Fail if UL regresses."),
        (3, "N.DL.SCH.*.ErrTB.Ibler / N.DL.SCH.*.TB", "DL IBLER by MCS",
         "iBeam precise/anti-intrf, tail-pkt MCS, hybrid IR (S15-16).",
         "Gate for MU/iBeam.",
         "DL IBLER %", "15 min / cell",
         "Fail if IBLER jumps vs control."),
        (4, "N.UL.SCH.*.ErrTB.Ibler / .TB", "UL IBLER",
         "S15-03 rank/CE and S15-19 UL Boosting.",
         "UL residual BLER.",
         "UL IBLER %", "15 min / cell",
         "Fail if UL IBLER jumps."),
        (5, "N.UL.SRS.PreSINR.Index* / N.SRS.NI.Avg", "SRS quality / SRS NI",
         "S15-02 weights, S15-14 joint PC, S15-16 tight MUX / blind IS.",
         "SRS that drives DL BF.",
         "SRS PreSINR / NI", "15 min / cell",
         "Fail if NI jumps — do not add SRS_IC."),
        (6, "N.ChMeas.PDSCH.MCS.k / N.PDSCH.InitTbDl.Rank*", "DL MCS / rank",
         "S15-02 / 12 tracking / 16 iBeam.",
         "BF gain = MCS upshift at same load.",
         "MCS / rank mix", "15 min / cell",
         "Informational with IBLER gate."),
        (7, "N.ChMeas.MIMO.DL.Pair.Layer.Avg / Pair.PRB", "DL MU pairing",
         "S15-07 already ON; S15-16 precise/anti-intrf; S15-09 later.",
         "Pairing quality not raw count.",
         "MU layers / PRB", "15 min / cell",
         "Pairing up + IBLER up = rollback MU extras."),
        (8, "N.ChMeas.MIMO.UL.Pair.Layer / Pair.PRB", "UL MU pairing",
         "S15-19 CR01 children.",
         "UL Boosting pairing.",
         "UL MU layers / PRB", "15 min / cell",
         "With UL IBLER."),
        (9, "N.CCE.DL.AllocReq.Num / N.CCE.DL.AggLvl* / N.CCE.Used.Avg", "PDCCH CCE / agg-level",
         "S15-16 PDCCH_AGG_LVL_COMPR.",
         "Blocking proxy.",
         "PDCCH CCE utilisation", "15 min / cell",
         "Fail if blocking rises."),
        (10, "N.User.OptimalSSBBeam.Avg / N.MAC.ThpVol.DL.OptimalSSB", "Optimal SSB users / volume",
         "S15-12 SSB adapt + tracking.",
         "Broadcast-beam follow.",
         "SSB coverage", "15 min / cell",
         "Drop/HO rise = rollback SSB bits."),
        (11, "N.ChMeas.CQI.SingleCW.k", "Wideband CQI",
         "S15-13 AHR Phase1 (already ON).",
         "CSI quality baseline.",
         "Average CQI", "15 min / cell",
         "Should already be commercial-like."),
        (12, "N.ChMeas.MIMO.DL.Transmission.Layer.Max", "Max DL layers on a PRB",
         "S15-05 SU quota (LAYER_16 live). S15-09 spends it later.",
         "Layer consumption.",
         "Max DL layers", "15 min / cell",
         "Do not drop after any MOD."),
        (13, "N.UECntx.AbnormRel / N.RRC.ReEst.Att", "Drop / re-establish",
         "Safety for S15-12 SSB + S15-16 iBeam in one window.",
         "Must not rise.",
         "Drop / re-est", "15 min / cell",
         "Any rise vs control → rollback that gNB."),
        (14, "N.PRB.DL.Used.Avg / N.PRB.UL.Used.Avg", "PRB usage (load)",
         "Normalise every tput/MCS comparison.",
         "Same BH, similar load.",
         "PRB %", "15 min / cell",
         "Do not compare empty vs loaded cells."),
        (15, "N.QoS.DL.PktDelayAirInterface.* / N.Traffic.DL.RlcFirstPktDelay.Time",
         "DL latency / first-packet delay",
         "S15-16 RLC merge + res-based adapt sch.",
         "Tail experience.",
         "Delay ms", "15 min / cell",
         "Should not rise."),
    ]
    for rec in ctrs:
        vals = list(rec) + [""] * (COLS - len(rec))
        fh = alt_fill(rec[0])
        r = putn(ws, r, vals, fills=[fh] * COLS, bolds=[False, True, True], center={1}, height=42)

    r = blank(ws, r, 8)
    r = subsection(ws, r, COLS, "MAE KPI (derived)", fill_hex=NAVY2)
    r = add_kpi_header(ws, r)
    kpis = [
        (1, "User DL Average Throughput (DU)", "N.ThpVol.DL / N.RLC.ThpTime.DL.Cell (MAE)", "Mbit/s",
         "Pass: no drop vs control. Target: lift vs D−7."),
        (2, "User UL Average Throughput (DU)", "N.ThpVol.UL / N.RLC.ThpTime.UL.Cell (MAE)", "Mbit/s",
         "Pass: no UL regression after Boosting/CE/rank."),
        (3, "DL IBLER", "Σ N.DL.SCH.*.ErrTB.Ibler / Σ N.DL.SCH.*.TB", "%",
         "Pass: within planned band vs control."),
        (4, "UL IBLER", "Σ N.UL.SCH.*.ErrTB.Ibler / Σ TB", "%",
         "Pass: no jump after UL_LOW_NOISE / rank fast decrease."),
        (5, "PDCCH CCE utilisation", "N.CCE.Used.Avg / N.CCE.Avail.*", "%",
         "Pass: used CCE falls or holds after agg-compress."),
        (6, "SSB optimal-beam users", "N.User.OptimalSSBBeam.Avg", "users",
         "Should rise after SSB adapt; watch drop."),
    ]
    for rec in kpis:
        r = add_kpi(ws, r, *rec)
        for c in range(11, COLS + 1):
            ws.cell(r - 1, c).fill = fill(alt_fill(rec[0]))
            ws.cell(r - 1, c).border = thin

    # patch TOC jump row
    href_row(ws.cell(jump_cells[0][0], jump_cells[0][1]), SHEET_NAME, sec1, "Section 1 — Suggestions + dump on MML")
    href_row(ws.cell(jump_cells[1][0], jump_cells[1][1]), SHEET_NAME, sec2, "Section 2 — Final proposal")
    href_row(ws.cell(jump_cells[2][0], jump_cells[2][1]), SHEET_NAME, sec3, "Section 3 — Counters / KPI")
    for _, i in jump_cells:
        ws.cell(jump_cells[0][0], i).fill = fill(PALE_BLUE)
    return ws


def main(out=None, zip_out=None, with_prereq=False, extra=None, version="v6.0"):
    """Build the combined workbook. extra(wb) may append further sheets before counters."""
    global VERSION
    out = out or OUT
    zip_out = zip_out or ZIP_OUT
    VERSION = version
    set_layout(with_prereq)
    if not os.path.exists(SRC):
        raise SystemExit(f"missing {SRC} — build v4.0 first")
    print("copy v4.0 →", os.path.basename(out))
    shutil.copy2(SRC, out)
    wb = load_workbook(out)
    if OLD14 in wb.sheetnames:
        print("remove", OLD14)
        del wb[OLD14]
    if OLD15 in wb.sheetnames:
        print("remove", OLD15)
        del wb[OLD15]
    print("combined sheet...")
    build_combined(wb)
    if CR01_OLD in wb.sheetnames:
        print("rename CR01", CR01_OLD, "→", CR01_NEW)
        wb[CR01_OLD].title = CR01_NEW
    print("patch cover...")
    patch_cover(wb)
    print("retarget hyperlinks to old 14/15...")
    retarget_hyperlinks(wb, OLD14, SHEET_NAME)
    retarget_hyperlinks(wb, OLD15, SHEET_NAME)
    retarget_hyperlinks(wb, CR01_OLD, CR01_NEW)
    if extra:
        print("extra sheets...")
        extra(wb)
    append_counters_to_all_sheets(wb)
    colors = ["1F4E79", "2E75B6", "0D7377", "C00000", "C65911", "548235",
              "7030A0", "1F4E79", "2E75B6", "0D7377", "C00000", "C65911",
              "548235", "7030A0", "C65911", "C00000", "A93226", "548235"]
    for i, ws in enumerate(wb.worksheets):
        ws.sheet_view.showGridLines = False
        if i < len(colors):
            ws.sheet_properties.tabColor = colors[i]
    print("saving", out)
    wb.save(out)
    print("ok", os.path.getsize(out), "sheets", len(wb.worksheets), wb.sheetnames)
    with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(out, os.path.basename(out))
    print("zip", zip_out, os.path.getsize(zip_out))
    return out


if __name__ == "__main__":
    main()
