#!/usr/bin/env python3
"""Build MIMO_Deployment_v5.0.xlsx from v4.0.

Combines former sheets 14 (Incon) + 15 (Suggestions) into one sheet:

  Section 1  Document suggestions (boxes; MML + Counter monitor / Impact on KPI / short Notes;
             Jump renamed to Jump to Basic)
  Section 2  Incon for those suggestions (dump: already enabled or not)
  Section 3  Final proposal — trial/implement what is NOT enabled
  Section 4  Performance counter and Monitoring KPI

CR01 stays as sheet 15 (was 16). v1–v4 files are unchanged.
"""
import os, shutil, zipfile, sys
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
OUT = os.path.join(ROOT, "MIMO_Deployment_v5.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v5.0.zip")

SHEET_NAME = "14. MIMO Suggest + Incon"  # 24 chars
CR01_OLD = "16. CR01 MIMO Exec Pack"
CR01_NEW = "15. CR01 MIMO Exec Pack"
OLD14 = "14. MIMO Incon Report"
OLD15 = "15. MIMO Suggestions from Doc"

COLS = 13
WIDTHS = [8, 22, 16, 16, 20, 22, 34, 26, 18, 24, 14, 12, 16]
# 1=SN, 2-6=MML, 7=Counter, 8=KPI impact, 9-10=short Notes, 11-13=Jump to Basic

BLUE_HDR = "5B9BD5"
YELLOW_HDR = "FFC000"
TEAL_HDR = "0D7377"
GOLD_HDR = "C9A227"
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


def box_border(ws, r1, r2, cols=COLS):
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


# Section 2: dump vs each suggestion (32T n=564 unless noted)
DUMP_ROWS = [
    # sid, family, param, live, status, enabled, gap, action, license
    ("S15-01", "Step1", "DL_PMI_SRS_ADAPT_SW", "ON 564/564", "Already ON", "Yes", "Matches commercial.", "Skip — do not re-send", "FBFD-010003"),
    ("S15-01", "Step1", "BeamPerceiveMode=DISTRIBUTED_MODE", "DISTRIBUTED_MODE 564/564", "Already ON", "Yes", "Prerequisite of weight opt is live.", "Skip", "FBFD-010003"),
    ("S15-01", "Step1", "FR1MaxCellCsirsPortNum=8PORT / TYPE0 / FD_RESOURCE", "8PORT+TYPE0+FD ON", "Already ON", "Yes", "CSI baseline live.", "Skip", "AHR/CSI"),
    ("S15-02", "Step1", "SRS_WEIGHT_ESTIMATE_SW", "OFF 564/564", "Missing", "No", "Largest basic-MIMO leftover.", "Enable (CR01 30/30)", "FBFD-010003"),
    ("S15-02", "Step1", "PMI_WEIGHT_OPT_SW", "OFF 564/564", "Hold", "No", "Missing vs FPD but mix with SRS tonight.", "Hold first night", "FBFD-010003"),
    ("S15-02", "Step1", "OPEN_LOOP_WEIGHT_OPT_SW", "OFF 564/564", "Hold", "No", "Same mix rule.", "Hold first night", "FBFD-010003"),
    ("S15-03", "Step1", "SRS_SINR_MEAS_OPT_SW", "OFF 564/564", "Missing", "No", "FPD example ON.", "Enable (CR01 30/30)", "FBFD-010003"),
    ("S15-03", "Step1", "UL_RANK_FAST_DECREASE_SW", "OFF 564/564", "Missing", "No", "Protects UL BLER.", "Enable (CR01 30/30)", "FBFD-010003"),
    ("S15-03", "Step1", "PUSCH_CE_SINR_LEVEL_ENH_SW", "OFF 564/564", "Missing", "No", "PUSCH CE enhance.", "Enable (CR01 30/30)", "FBFD-010003"),
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
    ("S15-14", "Step6", "SRS_IC_SW", "OFF 564/564", "Missing", "No", "Turbo child off. Tight MUX is in CR01.", "Hold this night (do not stack with CR01 tight MUX)", "FOFD-061201"),
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


def add_box(ws, r, rec, sheet_name):
    start = r
    fam = rec["family"]
    color = FAMILY_COLOR.get(fam, NAVY)
    merge(ws, r, 1, r, 10)
    put(ws, r, 1, f"  {rec['sid']}   ·   {rec['title']}", size=12, bold=True, color=WHITE,
        fill_hex=color, h="left", v="center")
    for c in range(2, 11):
        ws.cell(r, c).fill = fill(color)
    put(ws, r, 11, "Read benefit →", size=9, bold=True, color=WHITE, fill_hex=color, h="right", v="center")
    merge(ws, r, 12, r, 13)
    href_sheet(ws.cell(r, 12), rec["jump"], rec["jump"])
    ws.cell(r, 12).fill = fill("FFF2CC")
    ws.cell(r, 13).fill = fill("FFF2CC")
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

    # MML header: # | command | Counter monitor | Impact on KPI | short Notes | Jump to Basic
    hdr = [
        (1, "MML #", BLUE_HDR, False, False),
        (2, "MML Command  (one parameter / one switch / one line)  +  note at end", YELLOW_HDR, 2, 6),
        (7, "Counter monitor", TEAL_HDR, False, False),
        (8, "Impact on KPI", GOLD_HDR, False, False),
        (9, "short Notes", "F4B183", 9, 10),
        (11, "Jump to Basic", YELLOW_HDR, 11, 13),
    ]
    for col in range(1, COLS + 1):
        put(ws, r, col, "", size=9, bold=True, fill_hex=YELLOW_HDR, h="center", v="center", border=True)
    put(ws, r, 1, "MML #", size=9, bold=True, fill_hex=BLUE_HDR, h="center", v="center", border=True)
    merge(ws, r, 2, r, 6)
    put(ws, r, 2, "MML Command  (one parameter / one switch / one line)  +  note at end",
        size=9, bold=True, fill_hex=YELLOW_HDR, h="center", v="center", border=True)
    put(ws, r, 7, "Counter monitor", size=9, bold=True, color=WHITE, fill_hex=TEAL_HDR, h="center", v="center", border=True)
    put(ws, r, 8, "Impact on KPI", size=9, bold=True, fill_hex=GOLD_HDR, h="center", v="center", border=True)
    merge(ws, r, 9, r, 10)
    put(ws, r, 9, "short Notes", size=9, bold=True, fill_hex="F4B183", h="center", v="center", border=True)
    ws.cell(r, 10).fill = fill("F4B183")
    merge(ws, r, 11, r, 13)
    put(ws, r, 11, "Jump to Basic", size=9, bold=True, fill_hex=YELLOW_HDR, h="center", v="center", border=True)
    ws.cell(r, 12).fill = fill(YELLOW_HDR)
    ws.cell(r, 13).fill = fill(YELLOW_HDR)
    ws.row_dimensions[r].height = 24
    r += 1

    for i, (cmd, note) in enumerate(rec["mmls"], 1):
        line = mml_line(cmd, note)
        ctr, kpi, short = extras(cmd, note)
        fh = WHITE if i % 2 else ROW_ALT
        put(ws, r, 1, i, size=9, bold=True, fill_hex=fh, h="center", v="top", border=True)
        merge(ws, r, 2, r, 6)
        put(ws, r, 2, line, size=8, fill_hex=fh, h="left", v="top", border=True)
        for c in range(3, 7):
            ws.cell(r, c).fill = fill(fh)
            ws.cell(r, c).border = thin
        put(ws, r, 7, ctr, size=8, fill_hex="D5F5E3", h="left", v="top", border=True)
        put(ws, r, 8, kpi, size=8, fill_hex="FFF2CC", h="left", v="top", border=True)
        merge(ws, r, 9, r, 10)
        put(ws, r, 9, short, size=8, fill_hex="FDEBD0", h="left", v="top", border=True)
        ws.cell(r, 10).fill = fill("FDEBD0")
        ws.cell(r, 10).border = thin
        merge(ws, r, 11, r, 13)
        href_sheet(ws.cell(r, 11), rec["jump"], rec["jump"])
        ws.cell(r, 11).fill = fill("FFF2CC")
        ws.cell(r, 12).fill = fill("FFF2CC")
        ws.cell(r, 13).fill = fill("FFF2CC")
        ws.row_dimensions[r].height = min(56, max(26, 16 + len(line) // 120 * 12))
        r += 1
    end = r - 1
    box_border(ws, start, end)
    r = blank(ws, r, 10)
    return r, start


def patch_cover(wb):
    ws = wb["0. Cover & Index"]
    ws["A1"].value = "  5G MIMO (all features together)  —  Deployment Workbook  v5.0"
    r = ws.max_row + 2
    r = section(ws, r, 10, "v5.0 — sheets 14 + 15 combined (suggestion + dump incon in one place)")
    r = note_bar(ws, r, 10,
                 "Former “14. MIMO Incon Report” and “15. MIMO Suggestions from Doc” are now one sheet: "
                 f"{SHEET_NAME}. Section 1 = FPD suggestion boxes (MML + Counter monitor / Impact on KPI / short Notes; "
                 "Jump to Basic). Section 2 = dump check (already enabled or not). Section 3 = final trial proposal "
                 "for what is still OFF. Section 4 = Performance counter and Monitoring KPI. "
                 "CR01 is sheet 15. File: MIMO_Deployment_v5.0.xlsx")
    r = headers(ws, r, ["#", "Sheet", "Maps to", "What you will find"] + [""] * 6)
    r = table_row(ws, r,
                  ["14", SHEET_NAME, "FPD suggestions + DHK dump 8 Sep 2026",
                   "Boxes + dump enabled/not + Enable/Fix MML + counters"] + [""] * 6,
                  fills=[PALE_ORANGE] * 10, height=34)
    merge(ws, r - 1, 4, r - 1, 10)
    r = table_row(ws, r,
                  ["15", CR01_NEW, "CR01 10-site CME pack",
                   "Same 10 gNB execution pack as v4.0 (was sheet 16)"] + [""] * 6,
                  fills=[PALE_GOLD] * 10, height=30)
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
    ws.oddHeader.left.text = "5G MIMO Suggestions + Incon (v5.0) — dump check + trial proposal"
    ws.oddFooter.left.text = "Dump: 5G CME 8 Sep 2026 DHK 564×32T32R n41 · Suggestions from RAN10.1 FPDs · Jump to Basic = step sheet"
    ws.freeze_panes = "A4"
    ws.sheet_properties.tabColor = "C65911"

    items = suggestions()
    r = 1
    r = banner(ws, r, COLS,
               "  14.  MIMO Suggestions + Inconsistency  —  document boxes, dump check, trial proposal",
               fill_hex="C65911", size=16, height=30)
    r = note_bar(ws, r, COLS,
                 "One sheet instead of old 14 + 15. Section 1 keeps the FPD suggestion boxes and adds "
                 "Counter monitor / Impact on KPI / short Notes beside each MML; Jump is renamed Jump to Basic. "
                 "Section 2 is the dump incon for those same suggestions (already enabled or not on DHK 8 Sep 2026). "
                 "Section 3 is the final Enable/Fix list (what is still OFF). Section 4 is the monitoring pack. "
                 "Replace {NrDuCellId} / {NrDuCellTrpId}. Keep 2T2R indoor OFF. mmWave/DAS/Fusion = N/A on n41. "
                 "CR01 (sheet 15) already sends most Section 3 Enables on 10 macros.")

    toc_row = r
    r = section(ws, r, COLS, "Jump  (click)")
    jumps = [("Section 1 — Suggestions", None),
             ("Section 2 — Dump incon", None),
             ("Section 3 — Final proposal", None),
             ("Section 4 — Counters / KPI", None)]
    jump_cells = []
    for i, (lab, _) in enumerate(jumps, 1):
        put(ws, r, i, lab, size=9, bold=True, fill_hex=PALE_BLUE, h="center", v="center", border=True)
        jump_cells.append((r, i))
    ws.row_dimensions[r].height = 20
    r += 1
    r = blank(ws, r, 8)

    # ----- Section 1 -----
    sec1 = r
    r = section(ws, r, COLS, "Section 1.  Document suggestions  (Principal · Benefit · Parameter · MML)")
    r = note_bar(ws, r, COLS,
                 "Same boxes as v3.0 sheet 15. New columns after MML Command: Counter monitor, Impact on KPI, "
                 "short Notes (one sentence). Last column = Jump to Basic (the Step sheet). "
                 "This section is document-based — dump check is Section 2.")
    r = section(ws, r, COLS, "Index of suggestion boxes  (click ID to jump down this sheet)")
    r = headers(ws, r, ["ID", "Family", "Suggestion (click ID)", "Source step (Jump to Basic)"] + [""] * 9)
    toc_start = r
    for rec in items:
        put(ws, r, 1, rec["sid"], size=10, bold=True, fill_hex=PALE_GOLD, h="center", v="center", border=True)
        put(ws, r, 2, rec["family"], size=10, fill_hex=WHITE, h="center", v="center", border=True)
        merge(ws, r, 3, r, 10)
        put(ws, r, 3, rec["title"][:90], size=9, fill_hex=WHITE, h="left", v="center", border=True)
        merge(ws, r, 11, r, 13)
        href_sheet(ws.cell(r, 11), rec["jump"], rec["jump"])
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

    # ----- Section 2 -----
    sec2 = r
    r = section(ws, r, COLS, "Section 2.  Incon report for Section 1  (dump: already enabled or not)")
    r = note_bar(ws, r, COLS,
                 "Live = 5G CME ConfigurationData 8 Sep 2026 DHK, 564×32T32R n41 NSA (2T2R indoor excluded from trials). "
                 "Already ON = Skip. Missing = still OFF cluster-wide (CR01 may have it on 10 sites). "
                 "Hold = OFF but not this night. Fix = RF before switches. N/A = not this network. "
                 "Click Suggestion ID to jump to the Section 1 box.")
    r = headers(ws, r,
                ["SN", "Suggestion", "Family", "Parameter / switch", "Live DHK dump (32T)",
                 "Dump status", "Enabled?", "Gap vs commercial", "Action", "License",
                 "Jump to box", "", "Jump to Basic"],
                fill_hex=NAVY, height=26)
    for i, rec in enumerate(DUMP_ROWS, 1):
        sid, fam, param, live, status, enabled, gap, action, lic = rec
        fh = ST_FILL.get(status, WHITE)
        vals = [i, sid, fam, param, live, status, enabled, gap, action, lic, sid, "", ""]
        r = putn(ws, r, vals, fills=[fh] * COLS, bolds=[False, True, False, False, False, True],
                 center={1, 2, 6, 7}, height=40)
        if sid in box_rows:
            href_row(ws.cell(r - 1, 2), SHEET_NAME, box_rows[sid], sid)
            href_row(ws.cell(r - 1, 11), SHEET_NAME, box_rows[sid], sid)
        # jump to step from suggestion family
        jump = next((x["jump"] for x in items if x["sid"] == sid), None)
        if jump:
            merge(ws, r - 1, 12, r - 1, 13)
            href_sheet(ws.cell(r - 1, 12), jump, jump)

    r = blank(ws, r, 8)
    r = subsection(ws, r, COLS, "Section 2 counts (this dump check)", fill_hex=TEAL)
    counts = {}
    for rec in DUMP_ROWS:
        counts[rec[4]] = counts.get(rec[4], 0) + 1
    r = headers(ws, r, ["Dump status", "Rows", "Meaning"] + [""] * 10)
    meanings = {
        "Already ON": "Matches commercial — Skip / do not re-send.",
        "Missing": "OFF on cluster. Trial/implement (Section 3 Enable). CR01 covers most on 10 sites.",
        "Hold": "OFF but not this night (mix rule, next wave, or license).",
        "Fix": "RF first (tilt 255 / azimuth audit).",
        "Partial": "Master ON / child OFF, or CR01-only not cluster.",
        "Check": "Optional / LST — not a primary DL-tput gap.",
        "N/A": "DAS / Fusion / mmWave / 4T4R cable — not this n41 32T network.",
    }
    for st, n in counts.items():
        fh = ST_FILL.get(st, WHITE)
        r = putn(ws, r, [st, n, meanings.get(st, "")] + [""] * 10,
                 fills=[fh] * COLS, bolds=[True, True], height=22)
        merge(ws, r - 1, 3, r - 1, COLS)

    # ----- Section 3 -----
    sec3 = r
    r = section(ws, r, COLS, "Section 3.  Final proposal  —  trial / implement what is NOT enabled")
    r = note_bar(ws, r, COLS,
                 "Filtered from the dump incon: Action = Enable or Fix only. This is the set still OFF vs a commercial "
                 "32T network with good DL user throughput. CR01 (sheet 15) already sends most Enable lines on 10 macros "
                 "× 3 cells (101/102/103). Do not send LAYER_8. Do not add SRS_IC the same night as SRS_TIGHT_MULTIPLEXING. "
                 "2T2R indoor DHTIAA1 / DHAPT11 / DHTEJ34 stay OFF. Yellow columns = MML through License.")
    # 11-col MML in first 11 columns
    titles = ["SN", "RAT", "Doc Name", "Action", "MML Command (Proposed)",
              "MO Name", "Parameter ID", "Live Value (actual in Dump)",
              "proposed Parameter Value", "Purpose/Short Notes", "License", "", ""]
    for i, t in enumerate(titles, 1):
        if not t:
            continue
        fh = BLUE_HDR if i <= 4 else YELLOW_HDR
        put(ws, r, i, t, size=8, bold=True, fill_hex=fh, h="center", v="center", border=True)
    ws.row_dimensions[r].height = 28
    r += 1
    mml_start = r
    sn = 1
    for rec in incon_mml_rows():
        if rec[3] not in ("Enable", "Fix"):
            continue
        new = (sn,) + rec[1:]
        # reuse 11-col painter then pad cols 12-13
        r = incon_mml_row(ws, r, new)
        for c in range(12, COLS + 1):
            ws.cell(r - 1, c).fill = fill(YELLOW_HDR if rec[3] == "Enable" else "F4B183")
            ws.cell(r - 1, c).border = thin
        sn += 1
    ws.auto_filter.ref = f"A{mml_start - 1}:K{r - 1}"

    r = blank(ws, r, 8)
    r = bullets(ws, r, COLS, [
        "Cluster still OFF after CR01: same Enable list on the remaining 32T (not the 10 trial gNB).",
        "Hold (not in this proposal): PMI_WEIGHT_OPT / OPEN_LOOP_WEIGHT_OPT; SRS_IC while tight MUX is ON; multilayer; AHR CU 2.0; iBeam 2.0/3.0; freq-sel/smart sch; Boosting 2.0.",
        "Fix first if the cell is not a CR01 site: tilt=255 (28 cells cluster-wide) and azimuth vs RF design.",
        "Rollback: set the enabled bit back to 0 / HighPrecisionBeamSwitch=OFF / UL_LOW_NOISE_SW-0. Do not touch SU/MU/AHR Phase1+Turbo/LAYER_16.",
    ], fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 72

    # ----- Section 4 -----
    sec4 = r
    r = section(ws, r, COLS, "Section 4.  Performance counter and Monitoring KPI", fill_hex=TEAL)
    r = note_bar(ws, r, COLS,
                 "MAE DU, 15 min, busy hour. Trial = CR01 10 gNB (or any cell where Section 3 was sent). "
                 "Control = neighbour 32T not in the trial. Pre D−7, post D+1 and D+7. "
                 "Source: FPD Counter Changes for MIMO TDD / Beam / AHR / iBeam / UL Boosting.")
    r = headers(ws, r,
                ["SN", "Counter ID", "Counter Name", "Function", "Use for Section 1 suggestions",
                 "KPI it feeds", "Granularity", "Pass / fail gate", "", "", "", "", ""],
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
    href_row(ws.cell(jump_cells[0][0], jump_cells[0][1]), SHEET_NAME, sec1, "Section 1 — Suggestions")
    href_row(ws.cell(jump_cells[1][0], jump_cells[1][1]), SHEET_NAME, sec2, "Section 2 — Dump incon")
    href_row(ws.cell(jump_cells[2][0], jump_cells[2][1]), SHEET_NAME, sec3, "Section 3 — Final proposal")
    href_row(ws.cell(jump_cells[3][0], jump_cells[3][1]), SHEET_NAME, sec4, "Section 4 — Counters / KPI")
    for _, i in jump_cells:
        ws.cell(jump_cells[0][0], i).fill = fill(PALE_BLUE)
    return ws


def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"missing {SRC} — build v4.0 first")
    print("copy v4.0 → v5.0")
    shutil.copy2(SRC, OUT)
    wb = load_workbook(OUT)
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
    append_counters_to_all_sheets(wb)
    colors = ["1F4E79", "2E75B6", "0D7377", "C00000", "C65911", "548235",
              "7030A0", "1F4E79", "2E75B6", "0D7377", "C00000", "C65911",
              "548235", "7030A0", "C65911", "C00000"]
    for i, ws in enumerate(wb.worksheets):
        ws.sheet_view.showGridLines = False
        if i < len(colors):
            ws.sheet_properties.tabColor = colors[i]
    print("saving", OUT)
    wb.save(OUT)
    print("ok", os.path.getsize(OUT), "sheets", len(wb.worksheets), wb.sheetnames)
    with zipfile.ZipFile(ZIP_OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(OUT, os.path.basename(OUT))
    print("zip", ZIP_OUT, os.path.getsize(ZIP_OUT))


if __name__ == "__main__":
    main()
