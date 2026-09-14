#!/usr/bin/env python3
"""Add sheet 15: document-based MIMO performance suggestions (box layout).

Copies MIMO_Deployment_v2.0.xlsx → MIMO_Deployment_v3.0.xlsx.
Does not compare the live dump. Every enable-able FPD feature from sheets 0–13
is one box: Principal, Benefit, Parameter details, MML (one switch per line + note)
and a hyperlink back to the source step sheet.
"""
import os, shutil, zipfile, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from mimo_excel_style import *

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "MIMO_Deployment_v2.0.xlsx")
OUT = os.path.join(ROOT, "MIMO_Deployment_v3.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v3.0.zip")

SHEET_NAME = "15. MIMO Suggestions from Doc"  # 30 chars
COLS = 10
WIDTHS = [16, 10, 14, 12, 52, 18, 28, 24, 20, 22]
PID = "{NrDuCellId}"
TID = "{NrDuCellTrpId}"

LINK_BLUE = "0563C1"
MED = Border(
    left=Side(style="medium", color=NAVY),
    right=Side(style="medium", color=NAVY),
    top=Side(style="medium", color=NAVY),
    bottom=Side(style="medium", color=NAVY),
)
THIN_NAVY = Border(
    left=Side(style="thin", color="1F4E79"),
    right=Side(style="thin", color="1F4E79"),
    top=Side(style="thin", color="1F4E79"),
    bottom=Side(style="thin", color="1F4E79"),
)

# Jump targets (workbook sheet names)
S0, S1, S2 = "0. Cover & Index", "1. Overview of MIMO", "2. Type of MIMO Config"
S3, S4, S5 = "3. Step1 Basic MIMO", "4. Step2 SU-MIMO", "5. Step3 MU-MIMO"
S6, S7, S8 = "6. Step4 MM Multi-Layer", "7. Step5 Beam Management", "8. Step6 AHR"
S9, S10, S11 = "9. Step7 iBeam", "10. Step8 UL Boosting", "11. Step9 DAS+Fusion Cell"
S12, S13 = "12. Step10 mmWave", "13. Step11 Cable Sequence"
S14 = "14. MIMO Incon Report"

FAMILY_COLOR = {
    "Step1": "1F4E79", "Step2": "0D7377", "Step3": "C65911", "Step4": "C00000",
    "Step5": "7030A0", "Step6": "548235", "Step7": "BF8F00", "Step8": "2E75B6",
    "Step9": "595959", "Step10": "C00000", "Step11": "1F4E79",
}


def href_sheet(cell, sheet, text=None):
    cell.value = text or sheet
    # Internal Excel jump: location='Sheet'!A1  (do not put this in target)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A1", display=str(text or sheet))
    cell.font = Font(name="Calibri", size=10, color=LINK_BLUE, underline="single", bold=True)
    cell.alignment = align("left", "center", True)


def href_row(cell, sheet, row, text):
    cell.value = text
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A{row}", display=str(text))
    cell.font = Font(name="Calibri", size=10, color=LINK_BLUE, underline="single", bold=True)
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


def mml_line(cmd, note):
    """One parameter / one switch / one line. Note sits at the end of the command."""
    cmd = cmd.rstrip().rstrip(";")
    return f"{cmd};    // {note}"


# ---------------------------------------------------------------------------
# One box per FPD feature. MML = every activation line from the uploaded docs.
# ---------------------------------------------------------------------------
def suggestions():
    p, t = PID, TID
    return [
        dict(sid="S15-01", family="Step1", jump=S3,
             title="Basic MIMO — PMI/SRS weight adaptation + MetaAAU beam sensing",
             doc="MIMO TDD Ch.4.1 / 4.4.1.2  FBFD-010003",
             principal="NR TDD cells are already multi-antenna. DL beamforming uses PMI-based and/or SRS-based weights (FR1). MetaAAU beam sensing (BeamPerceiveMode=DISTRIBUTED_MODE) is the prerequisite of PMI-based and open-loop weight optimization. DL_PMI_SRS_ADAPT_SW must be ON before SrsNonASFixedWeightType takes effect.",
             benefit="Raises DL beamforming quality when SRS/PMI is available; non-AS UEs still get a PMI weight. Sensing lets later weight-opt bits work. KPI: User DL Average Throughput (DU).",
             params="NRDUCellAlgoSwitch.AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW; NRDUCellBeamAlgo.BeamPerceiveMode=DISTRIBUTED_MODE; NRDUCellPdschPrecode.SrsNonASFixedWeightType=PMI_WEIGHT; NRDUCellCsirs.FR1MaxCellCsirsPortNum=8PORT; NRDUCellPdschPrecode.SrsWeightValidityPeriod=MS400. License: none (basic functions).",
             mmls=[
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-1;",
                  "Enable SRS- and PMI-based weight adaptation (parent of non-AS weight type)."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, BeamPerceiveMode=DISTRIBUTED_MODE;",
                  "MetaAAU beam sensing. Prerequisite of PMI/open-loop weight opt. Use CENTRALIZED_MODE only if Hyper Cell."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, SrsWeightValidityPeriod=MS400;",
                  "How long an SRS weight stays valid (FPD sample MS400)."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, SrsNonASFixedWeightType=PMI_WEIGHT;",
                  "Weight type when SRS is not available. Needs DL_PMI_SRS_ADAPT_SW=1."),
                 (f"MOD NRDUCELLCSIRS: NrDuCellId={p}, FR1MaxCellCsirsPortNum=8PORT;",
                  "Max CSI-RS ports in the FR1 cell (FPD sample 8PORT)."),
             ]),
        dict(sid="S15-02", family="Step1", jump=S3,
             title="Basic MIMO — SRS / PMI / open-loop weight optimization (main BF gain)",
             doc="MIMO TDD Ch.4.1.3 / Table 4-4 / Ch.4.4.1.2  FBFD-010003",
             principal="After DISTRIBUTED_MODE is set, enable WeightAlgoSwitch bits one by one: SRS_WEIGHT_ESTIMATE_SW (SRS-based weights for DL large-packet UEs), then PMI_WEIGHT_OPT_SW, then OPEN_LOOP_WEIGHT_OPT_SW. Do not mix PMI + SRS + open-loop sources on the same night. Network impact: PMI-weight opt increases RRC reconfigurations; open-loop changes MCS/rank/IBLER and can raise scheduled UE count.",
             benefit="Largest basic-MIMO DL gain: large-packet UEs get estimated SRS weights instead of a frozen template. Mobility MCS/rank follow the channel.",
             params="NRDUCellBeamAlgo.WeightAlgoSwitch bits SRS_WEIGHT_ESTIMATE_SW / PMI_WEIGHT_OPT_SW / OPEN_LOOP_WEIGHT_OPT_SW = 1. Prerequisite: S15-01 sensing. License: none.",
             mmls=[
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1;",
                  "Enable SRS-based weight estimation for DL large-packet UEs (first weight bit)."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, WeightAlgoSwitch=PMI_WEIGHT_OPT_SW-1;",
                  "Enable PMI-based weight optimization. Hold until SRS-weight KPI is green (do not mix sources same night)."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, WeightAlgoSwitch=OPEN_LOOP_WEIGHT_OPT_SW-1;",
                  "Enable open-loop weight optimization. Needs DISTRIBUTED_MODE. Hold first night."),
             ]),
        dict(sid="S15-03", family="Step1", jump=S3,
             title="Basic MIMO — UL optimization (SRS SINR, fast UL rank drop, PUSCH CE)",
             doc="MIMO TDD Ch.4.1.3 UL optimization / Ch.4.4.1.2  FBFD-010003",
             principal="UL receive diversity is always on for multi-antenna cells. Extra UL bits: SRS_SINR_MEAS_OPT_SW, UL_RANK_FAST_DECREASE_SW, PUSCH_CE_SINR_LEVEL_ENH_SW. Also set SRS pre-SINR judge and board channel-measurement CPU decrease.",
             benefit="Protects UL BLER when the channel drops; better SRS quality feeds DL weights. KPI: User UL Average Throughput (DU).",
             params="NRDUCellSrs.SrsAlgoSwitch=SRS_SINR_MEAS_OPT_SW; NRDUCellUlRank.UlRankAlgoSw=UL_RANK_FAST_DECREASE_SW; NRDUCellPusch.PuschPerformanceSwitch=PUSCH_CE_SINR_LEVEL_ENH_SW; NRDUCellPdsch.SrsPreSinrJudgeThld=-100; GNodeBParam.NrBoardPerformanceSw=CHN_MEASURE_CPU_DEC_SW.",
             mmls=[
                 (f"MOD NRDUCELLSRS: NrDuCellId={p}, SrsAlgoSwitch=SRS_SINR_MEAS_OPT_SW-1;",
                  "Enable SRS SINR measurement optimization (UL enhancement)."),
                 (f"MOD NRDUCELLULRANK: NrDuCellId={p}, UlRankAlgoSw=UL_RANK_FAST_DECREASE_SW-1;",
                  "Enable fast UL rank decrease on a poor channel (protects UL BLER)."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, PuschPerformanceSwitch=PUSCH_CE_SINR_LEVEL_ENH_SW-1;",
                  "Enable PUSCH channel-estimation SINR-level enhancement."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, SrsPreSinrJudgeThld=-100;",
                  "Set SRS pre-SINR judge threshold (FPD sample −100 dB)."),
                 ("MOD GNODEBPARAM: NrBoardPerformanceSw=CHN_MEASURE_CPU_DEC_SW-1;",
                  "Enable channel-measurement CPU decrease (gNodeB-level, not per cell)."),
             ]),
        dict(sid="S15-04", family="Step1", jump=S3,
             title="Basic MIMO — DL scheduling time thresholds",
             doc="MIMO TDD Ch.4.4.1.2  FBFD-010003",
             principal="DlSchOptTimeThld and DlAdaptSchTimeThld bound how long DL scheduling optimization / adaptive scheduling run. FPD activation uses 300 and 30; setting DlSchOptTimeThld=0 is the document deactivation of that timer.",
             benefit="Lets DL scheduler hold an optimized allocation long enough to help average user throughput without freezing forever.",
             params="NRDUCellPdsch.DlSchOptTimeThld=300; NRDUCellDlSch.DlAdaptSchTimeThld=30.",
             mmls=[
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlSchOptTimeThld=300;",
                  "Set DL scheduling optimization time threshold (FPD sample 300; 0 disables)."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlAdaptSchTimeThld=30;",
                  "Set adaptive DL scheduling time threshold (FPD sample 30)."),
             ]),
        dict(sid="S15-05", family="Step2", jump=S4,
             title="SU-MIMO multiple layers — layer quota + DL rank adapt",
             doc="MIMO TDD Ch.5  FOFD-010020",
             principal="If a UE supports N layers, UL/DL peak is theoretically N × single-layer. Set MaxMimoLayerCnt / MaxMimoLayerNum after Step1. License FOFD-010020 + 2-layer capacity units (NR0S0DLEPU00 / NR0S0ULEPU00). Do not enable MU yet if you want SU-only layer counters. FPD sample uses LAYER_2 UL and LAYER_DEFAULT DL — live 32T may already be LAYER_16; never downgrade.",
             benefit="Unlocks SU peak layers (rank > 1). Counter N.ChMeas.MIMO.DL.Transmission.Layer.Max / UL.Trans.Layer.Max.",
             params="NRDUCellPusch.MaxMimoLayerCnt; NRDUCellPdsch.MaxMimoLayerNum; NRDUCellPdsch.DlLinkAdaptEnhancementSw=DL_RANK_ADAPT_SW; NRDUCellDlAmc.SuMimoPwrCtrlProtectThld=100. License NR0S0PREUM00 per cell + layer units.",
             mmls=[
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, MaxMimoLayerCnt=LAYER_2;",
                  "Set max UL SU-MIMO layers (FPD sample LAYER_2). Keep live LAYER_8 on 32T if already set — do not downgrade."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, MaxMimoLayerNum=LAYER_DEFAULT;",
                  "Set max DL SU-MIMO layers (FPD sample LAYER_DEFAULT). Keep live LAYER_16 on 32T — never send LAYER_8 as a downgrade."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlLinkAdaptEnhancementSw=DL_RANK_ADAPT_SW-1;",
                  "Enable DL rank adaptation."),
                 (f"MOD NRDUCELLDLAMC: NrDuCellId={p}, SuMimoPwrCtrlProtectThld=100;",
                  "Set SU-MIMO power-control protect threshold (FPD sample 100)."),
             ]),
        dict(sid="S15-06", family="Step2", jump=S4,
             title="SU-MIMO — DMRS overhead deduct + SRS-DTX precoding + optional JT",
             doc="MIMO TDD Ch.5.1.2 / 5.4.1.2  FOFD-010020",
             principal="SU DMRS overhead can be deducted adaptively. Precoding can be optimized when SRS is DTX. Intra-gNB DL JT (INTRA_GNB_DL_JT_SW) is in the FPD activation example — only if JT is intended.",
             benefit="More PDSCH REs for data (DMRS OH deduct) and more stable precoding when SRS is missing (DTX).",
             params="NRDUCellDmrs.DlDmrsSwitch=SU_DMRS_OH_ADAPT_DEDUCT_SW; NRDUCellPdschPrecode.DlPrecodeOptOnSrsDtxSw=SRS_PRECODE_OPT_SW; NRDUCellAlgoSwitch.CompSwitch=INTRA_GNB_DL_JT_SW.",
             mmls=[
                 (f"MOD NRDUCELLDMRS: NrDuCellId={p}, DlDmrsSwitch=SU_DMRS_OH_ADAPT_DEDUCT_SW-1;",
                  "Enable adaptive DMRS overhead deduction for SU."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, DlPrecodeOptOnSrsDtxSw=SRS_PRECODE_OPT_SW-1;",
                  "Enable precoding optimization on SRS DTX."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, CompSwitch=INTRA_GNB_DL_JT_SW-1;",
                  "Enable intra-gNB DL joint-transmission interaction (FPD sample — confirm JT is intended before sending)."),
             ]),
        dict(sid="S15-07", family="Step3", jump=S5,
             title="MU-MIMO basic pairing — master switch + isolation gates",
             doc="MIMO TDD Ch.6.1 / 6.4.1.2  FOFD-010010",
             principal="gNodeB pairs UEs with enough spatial isolation onto the same time-frequency resource. Pairing starts when MuMimoSwitch UL+DL is ON and isolation/SINR gates pass (DlPmiMuMimoSpaceIsoThld, DlSrsMuMimoSpaceIsoThld, DlMuMimoSrsPreSinrThld, UlMuMimoCorrThld, UlMuMimoSinrThld). UE falls back to SU when DlMuBackToSuSeThld trips. Requires Step2. Impacts: IBLER/MCS fluctuation; PDCCH MU can lower CCE success.",
             benefit="Cell spectral efficiency and average UE throughput in medium/heavy load (the core massive-MIMO capacity step).",
             params="MuMimoSwitch UL_MU_MIMO_SW & DL_MU_MIMO_SW; isolation 140 / 50; SRS PreSINR −50; group ISOLATION_CORRELATION; UL corr 9 / SINR −20; fallback 5. License NR0S00MUMM00 + layer units. Keep OFF on 2T2R.",
             mmls=[
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, MuMimoSwitch=UL_MU_MIMO_SW-1;",
                  "Enable UL MU-MIMO pairing (first master bit)."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, MuMimoSwitch=DL_MU_MIMO_SW-1;",
                  "Enable DL MU-MIMO pairing (second master bit)."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, MuMimoSwitch=PDCCH_MU_SW-1;",
                  "Enable PDCCH MU-MIMO (control-channel pairing; watch CCE success)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlPmiMuMimoSpaceIsoThld=140;",
                  "Set PMI-based DL MU isolation threshold (FPD sample 140; higher = fewer pairs)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlSrsMuMimoSpaceIsoThld=50;",
                  "Set SRS-based DL MU isolation threshold (FPD sample 50)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DlMuMimoSrsPreSinrThld=-50;",
                  "Set SRS pre-SINR gate for DL MU (FPD sample −50)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlMuMimoGroupMode=ISOLATION_CORRELATION;",
                  "Set how MU groups are built (FPD sample ISOLATION_CORRELATION)."),
                 (f"MOD NRDUCELLULMIMO: NrDuCellId={p}, UlMuMimoCorrThld=9;",
                  "Set UL MU correlation threshold (FPD sample 9)."),
                 (f"MOD NRDUCELLULMIMO: NrDuCellId={p}, UlMuMimoSinrThld=-20;",
                  "Set UL MU SINR threshold in dB (FPD sample −20)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DlMuBackToSuSeThld=5;",
                  "Set spectral-efficiency threshold to fall back from MU to SU (FPD sample 5)."),
             ]),
        dict(sid="S15-08", family="Step3", jump=S5,
             title="MU-MIMO — pairing helpers (rank, DMRS, SIR, heavy-load sch, PDCCH pair layers)",
             doc="MIMO TDD Ch.6.1.3 / 6.4.1.2  FOFD-010010",
             principal="Document activation also sets PDCCH MaxPairLayerNum, UL MaxMimoLayerCnt, SIR scale, PMI beam-number, DMRS RB policy, MU ranks, precoding IS value, and heavy-load scheduling priority.",
             benefit="Cleaner MU pairs and fairer scheduling under load; PDCCH can pair two users for grants.",
             params="MaxPairLayerNum=LAYER_2; MaxMimoLayerCnt=LAYER_4 (FPD); DlMuMimoSirScaleFactor=10; DlMuPmiBeamNumThld=3; DlMuEstRbPolicy=PSEUDOORTHOG_DMRS_ADAPT_DEDUCT; DlSrs/Pmi MuMimoRank=RANK_2; PrecodingIntrfSupprValue=DEFAULT; HEAVY_LOAD_SCH_PRI_OPT_SW.",
             mmls=[
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, MaxPairLayerNum=LAYER_2;",
                  "Set max PDCCH MU paired layers (FPD sample LAYER_2)."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, MaxMimoLayerCnt=LAYER_4;",
                  "Set UL layer cap during MU (FPD sample LAYER_4). Keep live LAYER_8 on 32T if already set."),
                 (f"MOD NRDUCELLDLAMC: NrDuCellId={p}, DlMuMimoSirScaleFactor=10;",
                  "Set DL MU SIR scale factor (FPD sample 10)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DlMuPmiBeamNumThld=3;",
                  "Set PMI beam-number threshold for MU (FPD sample 3)."),
                 (f"MOD NRDUCELLDMRS: NrDuCellId={p}, DlMuEstRbPolicy=PSEUDOORTHOG_DMRS_ADAPT_DEDUCT;",
                  "Set DL MU DMRS RB policy (FPD sample)."),
                 (f"MOD NRDUCELLDLRANK: NrDuCellId={p}, DlSrsMuMimoRank=RANK_2;",
                  "Set SRS-based MU pairing rank (FPD sample RANK_2)."),
                 (f"MOD NRDUCELLDLRANK: NrDuCellId={p}, DlPmiMuMimoRank=RANK_2;",
                  "Set PMI-based MU pairing rank (FPD sample RANK_2)."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, PrecodingIntrfSupprValue=DEFAULT;",
                  "Set precoding interference-suppression value (FPD sample DEFAULT)."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=HEAVY_LOAD_SCH_PRI_OPT_SW-1;",
                  "Enable heavy-load scheduling priority optimization."),
             ]),
        dict(sid="S15-09", family="Step4", jump=S6,
             title="Massive MIMO multi-layer (DL) — spend the layer quota",
             doc="MIMO TDD Ch.7.1  FR1 only",
             principal="As MU paired-layer count grows, interference, scheduling and weight accuracy become the bottleneck. HighLayerMuMimoSw master MMIMO_MULTILAYER_ENHANCE_SW plus MU_RANK_BOOSTING_SW, SRS_BLIND_IS_SW, hybrid precoding, tail-pkt MCS, PDCCH CCE/BWP0/blind-detect, PUCCH IF coord. FPD raises MaxMimoLayerNum to LAYER_8 and relaxes DlMuBackToSuSeThld to 0. Not for FR2. After Step3 is stable on 32T/64T.",
             benefit="Converts 32T/64T hardware into higher MU layers (the quota MaxMimoLayerNum cannot be consumed without these bits).",
             params="HighLayerMuMimoSw: MMIMO_MULTILAYER_ENHANCE / MU_RANK_BOOSTING / SRS_BLIND_IS; SrsBlindIsDegree=4; DL_HYBRID_PRECODING_SW; TAIL_PKT_MCS_OPT_SW; SmallPktType1RobustSchPol=LEVEL3. Layer licenses NR0S0DLEPU00.",
             mmls=[
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;",
                  "Enable master DL multi-layer enhance (parent of other HighLayer bits)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, HighLayerMuMimoSw=MU_RANK_BOOSTING_SW-1;",
                  "Enable MU rank boosting."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, HighLayerMuMimoSw=SRS_BLIND_IS_SW-1;",
                  "Enable SRS blind interference suppression."),
                 (f"MOD NRDUCELLSRS: NrDuCellId={p}, SrsBlindIsDegree=4;",
                  "Set degree of SRS blind IS (FPD sample 4)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, HighLayerMuMimoSw=MU_MIMO_PAIRING_PREFERRED_SW-1;",
                  "Enable prefer-MU pairing when gain exists (companion HighLayer bit)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, HighLayerMuMimoSw=SRS_MEAS_ACCELERATING_SW-1;",
                  "Enable faster SRS measurement for multilayer."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, PrecodingAlgoSwitch=DL_HYBRID_PRECODING_SW-1;",
                  "Enable DL hybrid precoding."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=TAIL_PKT_MCS_OPT_SW-1;",
                  "Enable tail-packet MCS optimization."),
                 (f"MOD NRDUCELLDLSCHRES: NrDuCellId={p}, SmallPktType1RobustSchPol=LEVEL3;",
                  "Set small-packet type1 robust scheduling policy (FPD sample LEVEL3)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DlMuBackToSuSeThld=0;",
                  "Relax SU-fallback so high-layer MU can stay (FPD multi-layer sample 0; was 5 in Step3)."),
                 (f"MOD NRDUCELLPUCCH: NrDuCellId={p}, PucchAlgoSwitch=PUCCH_INTRF_COORD_SW-1;",
                  "Enable PUCCH interference coordination."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoEnhSwitch=CCE_RESOURCE_OPT_SW-1;",
                  "Enable CCE resource optimization."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoEnhSwitch=UE_BWP0_PDSCH_RES_OPT_SW-1;",
                  "Enable BWP0 PDSCH resource optimization."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoSwitch=PDCCH_BLIND_DET_ASSIGN_OPT_SW-1;",
                  "Enable PDCCH blind-detection assignment optimization."),
             ]),
        dict(sid="S15-10", family="Step4", jump=S6,
             title="Massive MIMO multi-layer (UL) — demod, flex pair, adaptive UL sch",
             doc="MIMO TDD Ch.7.2  FR1 only",
             principal="UL chapter: multilayer demod enhancement, flexible MU pairing, resource- and latency-based adaptive UL scheduling, optional AI UL SU SINR predict (DSP NRDUCELLAISCHMODEL).",
             benefit="More UL MU layers and better UL scheduling under load/latency. KPI: User UL Average Throughput (DU).",
             params="UlHighLayerMuMimoSwitch: MULTILAYER_DEMOD_ENH_SW / MU_MIMO_FLEX_PAIR_SW; UlSchAlgoSwitch RES_BASED_ADAPT_ULSCH_SW / LATENCY_BASED_ADAPT_ULSCH_SW; optional AiAmcAlgoSwitch=UL_SU_SINR_INTEL_PREDICT_SW.",
             mmls=[
                 (f"MOD NRDUCELLULMIMO: NrDuCellId={p}, UlHighLayerMuMimoSwitch=MULTILAYER_DEMOD_ENH_SW-1;",
                  "Enable UL multilayer demodulation enhancement."),
                 (f"MOD NRDUCELLULMIMO: NrDuCellId={p}, UlHighLayerMuMimoSwitch=MU_MIMO_FLEX_PAIR_SW-1;",
                  "Enable flexible UL MU pairing."),
                 (f"MOD NRDUCELLULSCH: NrDuCellId={p}, UlSchAlgoSwitch=RES_BASED_ADAPT_ULSCH_SW-1;",
                  "Enable resource-based adaptive UL scheduling."),
                 (f"MOD NRDUCELLULSCH: NrDuCellId={p}, UlSchAlgoSwitch=LATENCY_BASED_ADAPT_ULSCH_SW-1;",
                  "Enable latency-based adaptive UL scheduling."),
                 (f"MOD NRDUCELLAIALGO: NrDuCellId={p}, AiAmcAlgoSwitch=UL_SU_SINR_INTEL_PREDICT_SW-1;",
                  "Optional AI UL SU SINR predict. Verify model with DSP NRDUCELLAISCHMODEL before cluster."),
             ]),
        dict(sid="S15-11", family="Step5", jump=S7,
             title="Beam Management — broadcast SSB/CSI + control-beam robustness",
             doc="Beam Mgmt Ch.4  FBFD-010015",
             principal="Broadcast beams (SSB/CSI-RS) set coverage. Control beams: PDCCH initial beam select, PDCCH beam robustness (DTX threshold), SRS beam select opt. 32T/64T use CoverageScenario + tilt/azimuth; 2/4/8T use RET/RVD. FPD sample uses NrDuCellId=1 / TrpId=1, Tilt=255 meaning product default in that sample — live network must use RF-designed tilt (255 on a live AAU is illegal).",
             benefit="SSB/CSI coverage + more robust PDCCH/SRS beams → idle camping, access, and connected-mode start of BF.",
             params="CoverageScenario / Tilt / Azimuth; SsbPeriod=MS20; CsiPeriod=SLOT40; PdcchPrecodeEnhPolicy=ENH_PMI; BeamOptSwitch=DL_INITIAL_BEAM_SELECT_SW; PdcchAlgoSwitch=PDCCH_BEAM_ROBUST_SW; PdcchBeamRobDtxThld=4; SrsAlgoSwitch=SRS_BEAM_SELECT_OPT_SW.",
             mmls=[
                 (f"MOD NRDUCELLTRPBEAM: NrDuCellTrpId={t}, CoverageScenario=DEFAULT, Tilt=255, Azimuth=0;",
                  "FPD sample default coverage. On live AAU replace Tilt/Azimuth with RF design — do not leave Tilt=255."),
                 (f"MOD NRDUCELL: NrDuCellId={p}, DuplexMode=CELL_TDD, SsbPeriod=MS20, SsbTimePos=DEFAULT;",
                  "Set SSB broadcast timing (FPD sample 20 ms)."),
                 (f"MOD NRDUCELLCSIRS: NrDuCellId={p}, CsiPeriod=SLOT40, CsirsCellResourceNum=4_RESOURCE;",
                  "Set CSI-RS period/resources for beam management (FPD sample). AHR later uses FD_RESOURCE."),
                 (f"MOD NRDUCELLPDCCHALGO: NrDuCellId={p}, PdcchPrecodeEnhPolicy=ENH_PMI;",
                  "Set PDCCH precoding policy to ENH_PMI."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, BeamOptSwitch=DL_INITIAL_BEAM_SELECT_SW-1;",
                  "Enable DL initial beam selection."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoSwitch=PDCCH_BEAM_ROBUST_SW-1;",
                  "Enable PDCCH beam robustness."),
                 (f"MOD NRDUCELLPDCCHALGO: NrDuCellId={p}, PdcchBeamRobDtxThld=4;",
                  "Set DTX count before robust beam action (FPD sample 4)."),
                 (f"MOD NRDUCELLSRS: NrDuCellId={p}, SrsAlgoSwitch=SRS_BEAM_SELECT_OPT_SW-1;",
                  "Enable SRS-based beam select optimization."),
             ]),
        dict(sid="S15-12", family="Step5", jump=S7,
             title="Beam Management — 3D coverage pattern + SSB adapt / beam tracking",
             doc="Beam Mgmt Ch.5  FOFD-010100  + BeamOptSwitch (SSB adapt / tracking)",
             principal="3D Coverage Pattern applies a scenario catalogue (traditional AAU, MetaAAU, 8T) or customized broadcast beams (ADD NRDUCELLTRPCUSTBEAM then CUSTOMIZED_BEAM_SCENARIO). SCENARIO_31/37 do not consume NR0SBSC3DC00. Connected-mode tracking: BEAM_TRACKING_SW + INTELLIGENT_BEAM_SELECTION_SW. SSB follow: SSB_BEAM_ADAPT_SW + SSB_BEAM_VERTICAL_COV_IMP_SW. Optional SCENARIO_BEAM_OPT_SW.",
             benefit="SSB grid matches the site (hotspot vs dense urban). Tracking follows the UE on the road — main mobility DL-tput beam lever after weights.",
             params="CoverageScenario=SCENARIO_n; WeightAlgoSwitch=SCENARIO_BEAM_OPT_SW; BeamOptAlgoSwitch BEAM_TRACKING / INTELLIGENT_BEAM_SELECTION; BeamOptSwitch SSB_BEAM_ADAPT / SSB_BEAM_VERTICAL_COV_IMP. License NR0SBSC3DC00 per cell.",
             mmls=[
                 (f"MOD NRDUCELLTRPBEAM: NrDuCellTrpId={t}, CoverageScenario=SCENARIO_1, Tilt=255, Azimuth=0;",
                  "Apply 3D coverage scenario 1 (FPD sample). Replace with RF-planned SCENARIO_n. Consumes NR0SBSC3DC00."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, WeightAlgoSwitch=SCENARIO_BEAM_OPT_SW-1;",
                  "Enable weight optimization for the coverage scenario."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, BeamOptAlgoSwitch=BEAM_TRACKING_SW-1;",
                  "Enable connected-mode beam tracking (follow the UE)."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, BeamOptAlgoSwitch=INTELLIGENT_BEAM_SELECTION_SW-1;",
                  "Enable intelligent beam selection."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, BeamOptSwitch=SSB_BEAM_ADAPT_SW-1;",
                  "Enable SSB beam adaptation so SSB is not a frozen grid."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, BeamOptSwitch=SSB_BEAM_VERTICAL_COV_IMP_SW-1;",
                  "Enable SSB vertical coverage improve (high-rise / downtilt shaping)."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1;",
                  "Enable beam selection optimization (also an iBeam 1.0 child — see S15-16)."),
             ]),
        dict(sid="S15-13", family="Step6", jump=S8,
             title="AHR Introduction (Phase1) — adaptive CSI-RS multi-beam",
             doc="AHR Ch.4  FOFD-051301",
             principal="Adaptive CSI-RS multi-beam adjustment + adaptive multi-stream perception. Master AHR_PHASE1_SW. Companions: CSIRS_INTRF_STATIC_AVOID, FD_RESOURCE + TYPE0, experience-based / resource-based MM adapt sch, AMC start, delay buffer, sch timers.",
             benefit="CSI beams and stream perception follow the UE — baseline of later Turbo / Capacity Upgrade. KPI: DL experience, CSI beam adapt.",
             params="AhrSwitch=AHR_PHASE1_SW; CsiSwitch=CSIRS_INTRF_STATIC_AVOID_SW; CsirsCellResourceNum=FD_RESOURCE; CsirsBeamType=TYPE0; ServiceExpAlgoSwitch EXP_BASED_MM_ADAPT_SCH / RES_BASED_MM_ADAPT_SCH. License NR0S00AET100 per cell.",
             mmls=[
                 (f"MOD NRDUCELLCSIRS: NrDuCellId={p}, CsiSwitch=CSIRS_INTRF_STATIC_AVOID_SW-1;",
                  "Enable CSI-RS static interference avoid (Phase1 companion)."),
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, AhrSwitch=AHR_PHASE1_SW-1;",
                  "Enable AHR Introduction master switch."),
                 (f"MOD NRDUCELLCSIRS: NrDuCellId={p}, CsirsCellResourceNum=FD_RESOURCE;",
                  "Set CSI-RS cell resource to FD_RESOURCE (AHR CSI baseline)."),
                 (f"MOD NRDUCELLCSIRS: NrDuCellId={p}, CsirsBeamType=TYPE0;",
                  "Set CSI-RS beam type to TYPE0 (AHR CSI baseline)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, FixedAmcStepValue=40;",
                  "Set fixed AMC step (FPD sample 40)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlInitialMcsAdjValue=0;",
                  "Set DL initial MCS adjust value (FPD sample 0)."),
                 (f"MOD NRDUCELLSERVEXP: NrDuCellId={p}, ServiceExpAlgoSwitch=EXP_BASED_MM_ADAPT_SCH_SW-1;",
                  "Enable experience-based massive-MIMO adaptive scheduling."),
                 (f"MOD NRDUCELLSERVEXP: NrDuCellId={p}, ServiceExpAlgoSwitch=RES_BASED_MM_ADAPT_SCH_SW-1;",
                  "Enable resource-based massive-MIMO adaptive scheduling."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlDelaySchBufferThld=1;",
                  "Set DL delay-schedule buffer threshold (FPD sample 1)."),
             ]),
        dict(sid="S15-14", family="Step6", jump=S8,
             title="AHR Experience Turbo 2.0 — SRS IC + precise AMC",
             doc="AHR Ch.5  FOFD-061201",
             principal="After Phase1 is green: SRS interference coordination + precise AMC. Master AHR_EXP_TURBO_PHASE2_SW. Do not enable LOW_SNR_CHN_DENOISE together with Turbo (FPD sets it 0). Do not stack new SRS_IC with iBeam SRS_TIGHT_MULTIPLEXING the same night.",
             benefit="DL tput under SRS collision; more accurate MCS/rank. Turbo master without SRS_IC is only half-activated.",
             params="AhrSwitch=AHR_EXP_TURBO_PHASE2_SW; SrsIntrfThld=5; SRS_JOINT_PC_SW; IntrfUeSrsPcMinSinrTarget=50; MaxSrsPoAdjustAmount=6; TAIL_PKT_SCH_OPT_SW; optional DL_SU_MCS_INTEL_OPT. License NR0S00AET200.",
             mmls=[
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, AhrSwitch=AHR_EXP_TURBO_PHASE2_SW-1;",
                  "Enable AHR Experience Turbo 2.0 master (only after Phase1 is green)."),
                 (f"MOD NRDUCELLSRS: NrDuCellId={p}, SrsDetectionAlgoSwitch=SRS_IC_SW-1;",
                  "Enable SRS interference coordination (Turbo leftover if master is already ON)."),
                 (f"MOD NRDUCELLSRSMEAS: NrDuCellId={p}, SrsIntrfThld=5;",
                  "Set SRS interference threshold (FPD Turbo sample 5)."),
                 (f"MOD NRDUCELLULPCCONFIG: NrDuCellId={p}, UlPwrCtrlAlgoSwitch=SRS_JOINT_PC_SW-1;",
                  "Enable SRS joint power control (recommended with SRS IC)."),
                 (f"MOD NRDUCELLULPCCONFIG: NrDuCellId={p}, IntrfUeSrsPcMinSinrTarget=50;",
                  "Set interfered-UE SRS PC min SINR target (FPD sample 50)."),
                 (f"MOD NRDUCELLULPCCONFIG: NrDuCellId={p}, MaxSrsPoAdjustAmount=6;",
                  "Set max SRS power-adjust amount (FPD sample 6)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlPdschAlgoSwitch=TAIL_PKT_SCH_OPT_SW-1;",
                  "Enable tail-packet scheduling opt (Turbo precise-AMC helper)."),
                 (f"MOD NRDUCELLDLAMC: NrDuCellId={p}, DlAmcAlgoSw=DL_MCS_ADJ_OPT_SW-1;",
                  "Enable DL MCS adjust optimization (Turbo helper)."),
                 (f"MOD NRDUCELLDLRANK: NrDuCellId={p}, DlRankSelAlgoSw=RANK_AND_SINR_ESTIMATE_OPT_SW-1;",
                  "Enable rank-and-SINR estimate optimization (Turbo helper)."),
             ]),
        dict(sid="S15-15", family="Step6", jump=S8,
             title="AHR Capacity Upgrade 2.0 — high-res MU pairing / joint sch",
             doc="AHR Ch.6  FOFD-061202",
             principal="Loaded-hour capacity wave after Phase1+Turbo. Master AHR_CAPC_UPGRADE_PHASE2_SW. Needs MU-MIMO (Step3). Children: MuMimoOptSwith (FPD spelling), gathering MU, multi-dimension joint PDCCH/PDSCH sch, MU IR precoding.",
             benefit="Loaded-hour DL cell capacity and average user tput — the AHR wave that actually upgrades pairing resolution.",
             params="AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW; MuMimoOptSwith=ON; DlMuGatherOptSw=ON; PDCCH_MULTI_DIM_JOINT_SCH_SW; DlMuIRPrecodePol=MU_ALLUSER_POWER_ENH. License NR0S00ACT200. Needs MU.",
             mmls=[
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1;",
                  "Enable AHR Capacity Upgrade 2.0 master (after Turbo SRS-IC is KPI-green)."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, MuMimoOptSwith=ON;",
                  "Enable MU-MIMO optimization (FPD spelling MuMimoOptSwith)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlMuGatherOptSw=ON;",
                  "Enable DL gathering MU-MIMO optimization."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlSrsMuMimoPreSinrThld=0;",
                  "Set DL SRS MU pre-SINR threshold for capacity phase (FPD sample 0)."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, MuMimoSwitch=PDCCH_MULTI_DIM_JOINT_SCH_SW-1;",
                  "Enable multi-dimension joint PDCCH/PDSCH scheduling."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, DlMuIRPrecodePol=MU_ALLUSER_POWER_ENH;",
                  "Set MU interference-reject precoding policy (FPD sample MU_ALLUSER_POWER_ENH)."),
             ]),
        dict(sid="S15-16", family="Step7", jump=S9,
             title="iBeam 1.0 — high-precision beams + interference-aware MU/scheduling",
             doc="iBeam Ch.3  FOFD-081201",
             principal="Main DL-interference package on loaded urban 32T. Master HighPrecisionBeamSwitch=ON then children: SRS blind IS measurement, tight SRS multiplexing, hybrid BWP interference-random, RLC merge sch, PDCCH agg compress, beam select opt, precise + anti-intrf MU sch, tail-pkt MCS, resource-based adapt sch. Requires MU (Step3). Do not enable 2.0/3.0 without 1.0. Do not stack SRS_TIGHT_MULTIPLEXING with a new AHR SRS_IC the same night.",
             benefit="DL user tput in interference, MU quality, PDCCH blocking, tail packets — what commercial interference-limited 32T runs.",
             params="HighPrecisionBeamSwitch=ON; SRS_BLIND_IS_MEAS_SW; SRS_TIGHT_MULTIPLEXING_SW; DL_BWP_HYBRID_INTRF_RANDOM_SW; DL_RLC_STAT_RPT_MERGE_SCH_SW; PDCCH_AGG_LVL_COMPR_SW (thld 60); BEAM_SELECT_OPT_SW; DL_MU_PRECISE_SCH_SW; DL_MU_ANTI_INTRF_SCH_SW; TAIL_PKT_MCS_OPT_SW; RES_BASED_DL_ADAPT_SCH_SW. License NR0S00BEAM00. LST LICENSE first.",
             mmls=[
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, HighPrecisionBeamSwitch=ON;",
                  "Enable iBeam 1.0 master (LST LICENSE NR0S00BEAM00 first)."),
                 (f"MOD NRDUCELLSRSMEAS: NrDuCellId={p}, SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-1;",
                  "Enable blind IS for SRS measurement."),
                 (f"MOD NRDUCELLSRS: NrDuCellId={p}, SrsAlgoSwitch=SRS_TIGHT_MULTIPLEXING_SW-1;",
                  "Enable tight SRS multiplexing. Do not stack with new SRS_IC the same night."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=DL_BWP_HYBRID_INTRF_RANDOM_SW-1;",
                  "Enable DL multi-BWP hybrid interference randomization."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=DL_RLC_STAT_RPT_MERGE_SCH_SW-1;",
                  "Enable merged DL RLC status-report scheduling."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoSwitch=PDCCH_AGG_LVL_COMPR_SW-1;",
                  "Enable PDCCH aggregation-level compression (set AggLvlComprCceUsageThld=60 if not already)."),
                 (f"MOD NRDUCELLBEAMALGO: NrDuCellId={p}, ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1;",
                  "Enable beam selection optimization."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-1;",
                  "Enable DL MU precise scheduling."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DLMuMimoSchSupplementSw=DL_MU_ANTI_INTRF_SCH_SW-1;",
                  "Enable DL MU anti-interference scheduling."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=TAIL_PKT_MCS_OPT_SW-1;",
                  "Enable tail-packet MCS opt (recommended with precise/anti-intrf MU)."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=RES_BASED_DL_ADAPT_SCH_SW-1;",
                  "Enable resource-based adaptive DL scheduling."),
             ]),
        dict(sid="S15-17", family="Step7", jump=S9,
             title="iBeam 2.0 — robust weights + dynamic MU cluster",
             doc="iBeam Ch.4  FOFD-091200",
             principal="After iBeam 1.0 is KPI-green. Master HighPrecisionBeamPhase2Sw. Robust weights (DL_ROBUST_WEIGHT_SW, RobustWtPhaseCalcMethod=SRS_H_BASED_CALCULATION), dynamic cluster grouping, correlation acceleration, far-UE rank, precise MU eval, optional inter-cell interference avoid.",
             benefit="Second-wave DL robustness in interference; better pairing for far UEs.",
             params="HighPrecisionBeamPhase2Sw=ON; PrecodingAlgoSwitch=DL_ROBUST_WEIGHT_SW; RobustWtPhaseCalcMethod=SRS_H_BASED_CALCULATION; SrsIntrfThld=3; SrsBlindIsDegree=7; DlMuMimoGroupMode=DYNAMIC_CLUSTER_GROUP; PRECISE_MUMIMO_EVAL / FAR_UE_RANK_OPT / DL_CORR_ACCELERATION; optional INTER_CELL_INTRF_AVOID_SW. License NR0S0DLEHR00.",
             mmls=[
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, HighPrecisionBeamPhase2Sw=ON;",
                  "Enable iBeam 2.0 master (only after 1.0 KPI-green)."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, PrecodingAlgoSwitch=DL_ROBUST_WEIGHT_SW-1;",
                  "Enable robust weight."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, RobustWtPhaseCalcMethod=SRS_H_BASED_CALCULATION;",
                  "Set robust-weight phase calculation to SRS_H_BASED_CALCULATION."),
                 (f"MOD NRDUCELLSRSMEAS: NrDuCellId={p}, SrsIntrfThld=3;",
                  "Set SRS interference threshold (iBeam 2.0 sample 3; Turbo used 5 — do not fight Turbo on the same night)."),
                 (f"MOD NRDUCELLSRS: NrDuCellId={p}, SrsBlindIsDegree=7;",
                  "Set SRS blind IS degree (iBeam 2.0 sample 7)."),
                 (f"MOD NRDUCELLPDSCH: NrDuCellId={p}, DlMuMimoGroupMode=DYNAMIC_CLUSTER_GROUP;",
                  "Switch MU grouping to dynamic cluster (overrides Step3 ISOLATION_CORRELATION while 2.0 is ON)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, MuMimoIblerTarget=15;",
                  "Set MU-MIMO IBLER target (FPD sample 15)."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DLMuMimoSchSupplementSw=PRECISE_MUMIMO_EVAL_SW-1;",
                  "Enable precise MU-MIMO evaluation."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DLMuMimoSchSupplementSw=FAR_UE_RANK_OPT_SW-1;",
                  "Enable far-UE rank optimization."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DLMuMimoSchSupplementSw=DL_CORR_ACCELERATION_SW-1;",
                  "Enable DL correlation-acceleration pairing."),
                 (f"MOD NRDUCELLCOLLABSERV: NrDuCellId={p}, MultiCellMimoSwitch=INTER_CELL_INTRF_AVOID_SW-1;",
                  "Optional inter-cell interference avoid (needs cluster/collab ready)."),
             ]),
        dict(sid="S15-18", family="Step7", jump=S9,
             title="iBeam 3.0 — self-fusion weights + smart AMC",
             doc="iBeam Ch.5  FOFD-100200",
             principal="Last iBeam wave after 2.0. Master HighPrecisionBeamPhase3Sw. Self-fusion weights, PDCCH robust weights, smart AMC, optional AI DL MU MCS.",
             benefit="Residual interference-limited DL tput after 1.0+2.0.",
             params="HighPrecisionBeamPhase3Sw=ON; DL_SELF_FUSION_WEIGHT_SW; SelfFusionWtSrsPresinrThld=2; SelfFusionWtRhoThld=80; PDCCH_ROBUST_WEIGHT_SW; DL_SMART_AMC_SW; optional DL_MU_MCS_INTEL_OPT_SW. License NR0S00BEAM30. Do not enable now if 1.0 is still OFF.",
             mmls=[
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, HighPrecisionBeamPhase3Sw=ON;",
                  "Enable iBeam 3.0 master (only after 2.0 KPI-green)."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, PrecodingAlgoSwitch=DL_SELF_FUSION_WEIGHT_SW-1;",
                  "Enable self-fusion weight."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, SelfFusionWtSrsPresinrThld=2;",
                  "Set self-fusion weight SRS pre-SINR threshold (FPD sample 2)."),
                 (f"MOD NRDUCELLPDSCHPRECODE: NrDuCellId={p}, SelfFusionWtRhoThld=80;",
                  "Set self-fusion weight rho threshold (FPD sample 80)."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoSwitch=PDCCH_ROBUST_WEIGHT_SW-1;",
                  "Enable PDCCH robust weights."),
                 (f"MOD NRDUCELLDLMIMO: NrDuCellId={p}, DLMuMimoSchSupplementSw=DL_SMART_AMC_SW-1;",
                  "Enable smart AMC (iBeam 3.0)."),
             ]),
        dict(sid="S15-19", family="Step8", jump=S10,
             title="Uplink Boosting 1.0 — PUSCH/PUCCH/SRS interference reduction",
             doc="UL Boosting Ch.4  FOFD-091201",
             principal="Master UL_LOW_NOISE_SW. Phase1: UL MU group/flex/corr pairing, PUSCH resource adapt, OLLA, waveform SINR, PDCCH symbol smart alloc, PUCCH/CSI period opt, IRC-based PDP detect. Skip RedCap / BWP2 if those MOs are not on the NSA layer. Skip PKT_LEN_BASED_SCH_OPT if CCE is tight.",
             benefit="UL user tput and SRS quality that feeds DL weights. Secondary for DL but required for UL KPI.",
             params="MimoFeatureSwitch=UL_LOW_NOISE_SW; UL_MU_GRP_PAIR / DIFF_WAVEFORM_PAIR / UL_CORR_ACCELERATION; PUSCH_RES_ADAPT_ALLOC; UL_CELL_OLLA; SinrThldforWaveformSel=32; PDCCH_SYM_SMART_ALLOC; F1_ACK_CODE_CHN_INTRF_OPT; IRC_BASED_PDP_DETECT. License NR0S00UAHR00. 32T only.",
             mmls=[
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, MimoFeatureSwitch=UL_LOW_NOISE_SW-1;",
                  "Enable Uplink Boosting 1.0 master (LST LICENSE NR0S00UAHR00)."),
                 (f"MOD NRDUCELLULMIMO: NrDuCellId={p}, UlMuMimoAlgoSwitch=UL_MU_GRP_PAIR_SW-1;",
                  "Enable UL MU grouping/pairing."),
                 (f"MOD NRDUCELLULMIMO: NrDuCellId={p}, UlMuMimoAlgoSwitch=DIFF_WAVEFORM_PAIR_SW-1;",
                  "Enable mixed-waveform UL MU pairing."),
                 (f"MOD NRDUCELLULMIMO: NrDuCellId={p}, UlMuMimoAlgoSwitch=UL_CORR_ACCELERATION_SW-1;",
                  "Enable UL correlation-acceleration pairing."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, PuschAlgoExtSwitch=PUSCH_RES_ADAPT_ALLOC_SW-1;",
                  "Enable adaptive PUSCH resource allocation."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, UlPreallocationSwitch=UL_PREALLOCATION_PERIOD_ADJ_SW-1;",
                  "Enable UL preallocation period adjust."),
                 (f"MOD NRDUCELLULAMC: NrDuCellId={p}, UlAmcAlgoSw=UL_CELL_OLLA_SW-1;",
                  "Enable UL cell-level OLLA."),
                 (f"MOD NRDUCELLULAMC: NrDuCellId={p}, UlAmcPerformanceSw=SMALL_PKT_OL_ADAPT_ADJ_SW-1;",
                  "Enable small-packet outer-loop adapt."),
                 (f"MOD NRDUCELLULAMC: NrDuCellId={p}, UlAmcPerformanceSw=SR_BASED_SCH_MCS_OPT_SW-1;",
                  "Enable SR-based MCS optimization."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, SinrThldforWaveformSel=32;",
                  "Set waveform-select SINR threshold (FPD sample 32)."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoEnhSwitch=PDCCH_SYM_SMART_ALLOC_SW-1;",
                  "Enable smart PDCCH symbol allocation."),
                 (f"MOD NRDUCELLPDCCH: NrDuCellId={p}, PdcchAlgoSwitch=CCE_SYMBOL_ALLOC_OPT_SW-1;",
                  "Enable CCE symbol allocation opt (FPD optional for 40–70 MHz)."),
                 (f"MOD NRDUCELLPUCCH: NrDuCellId={p}, PucchAlgoExtSwitch=F1_ACK_CODE_CHN_INTRF_OPT_SW-1;",
                  "Enable PUCCH F1 ACK code-channel interference opt."),
                 (f"MOD NRDUCELLDMRS: NrDuCellId={p}, UlDmrsSwitch=IRC_BASED_PDP_DETECT_SW-1;",
                  "Enable IRC-based PDP detect (RS interference reduction)."),
             ]),
        dict(sid="S15-20", family="Step8", jump=S10,
             title="Uplink Boosting 2.0 — coordinated PC + precise RB/MCS + multi-beam RX",
             doc="UL Boosting Ch.5  FOFD-100201",
             principal="After Boosting 1.0 is green. Master UL_LOW_NOISE_PHASE2_SW. Coordinated PUSCH PC needs inter-gNB time sync (FPD mentions DSP CLKTST). Adds link-perf PC, precise retrans RB, precise MCS / aware sch, multi-beam RX, precise freq-offset and channel estimation.",
             benefit="UL tput at interference and more efficient UL retransmission.",
             params="MimoFeatureSwitch=UL_LOW_NOISE_PHASE2_SW; PUSCH_COORD_PWR_CTRL_SW; INTRF_SC_LINK_PERF_PC_SW; UlCpcUeSsbRsrpThld=-120; UL_RETRANS_PREC_RB_CTRL_SW; UlFirstRetransMinRbPct=20; UL_PRECISE_MCS_OPT_SW; MULTI_BEAM_RX_ENH_SW. License NR0S00UPBT20.",
             mmls=[
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, MimoFeatureSwitch=UL_LOW_NOISE_PHASE2_SW-1;",
                  "Enable Uplink Boosting 2.0 master (after 1.0)."),
                 (f"MOD NRDUCELLULPCCONFIG: NrDuCellId={p}, UlPwrCtrlAlgoExtSwitch=PUSCH_COORD_PWR_CTRL_SW-1;",
                  "Enable multi-cell PUSCH coordinated PC (needs TIME_SYNC — DSP CLKTST first)."),
                 (f"MOD NRCELLMEASCONFIG: NrCellId={p}, UlCpcUeSsbRsrpThld=-120;",
                  "Set coordinated-PC UE SSB RSRP threshold (FPD sample −120)."),
                 (f"MOD NRDUCELLULPCCONFIG: NrDuCellId={p}, UlPwrCtrlAlgoExtSwitch=INTRF_SC_LINK_PERF_PC_SW-1;",
                  "Enable interference-scenario link-performance PC."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, PuschAlgoExtSwitch=UL_RETRANS_PREC_RB_CTRL_SW-1;",
                  "Enable precise UL retransmission RB control."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, UlFirstRetransMinRbPct=20;",
                  "Set first-retrans min RB percent (FPD sample 20)."),
                 (f"MOD NRDUCELLAIALGO: NrDuCellId={p}, AiAmcAlgoSwitch=UL_PRECISE_MCS_OPT_SW-1;",
                  "Enable precise UL MCS (optional AI)."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, PuschAlgoExtSwitch=MULTI_BEAM_RX_ENH_SW-1;",
                  "Enable multi-beam RX enhancement."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, PuschAlgoExtSwitch=INTRF_PREC_FREQ_OFS_EST_SW-1;",
                  "Enable precise frequency-offset estimation under interference."),
                 (f"MOD NRDUCELLPUSCH: NrDuCellId={p}, PuschAlgoExtSwitch=PREC_CHANNEL_EST_SW-1;",
                  "Enable precise channel estimation."),
             ]),
        dict(sid="S15-21", family="Step9", jump=S11,
             title="Distributed Massive MIMO (DAS) — multi-TRP one cell",
             doc="DAS FPD Ch.4  FOFD-050202 / FOFD-071211",
             principal="Site architecture after air-interface MIMO is green. One NR DU cell, multiple TRPs (master + slave). Master switch DmMimoSwitch=DM_MIMO_SERVICE_SWITCH. Optional PDCCH TRP trans, TRP-select weight, UL beam boost. Do not enable on a single-AAU-per-sector cluster.",
             benefit="Coverage fill and UL boost in indoor LampSite or macro split-TRP sites.",
             params="TxRxMode / TrpType=SLAVE; DM_MIMO_SERVICE_SWITCH; optional DM_MIMO_PDCCH_TRP_TRANS_SW / DM_MIMO_TRP_SEL_WEIGHT_SW. License NR0SDMMIMO00 (LampSite) or NR0S00VMMM00 (macro).",
             mmls=[
                 (f"MOD NRDUCELLTRP: NrDuCellTrpId={t}, NrDuCellId={p}, TxRxMode=4T4R;",
                  "Set master TRP (FPD sample 4T4R — use planned TxRxMode)."),
                 (f"MOD NRDUCELLTRP: NrDuCellTrpId={{SlaveTrpId}}, NrDuCellId=65535, TxRxMode=4T4R, TrpType=SLAVE;",
                  "Add slave TRP (FPD sample NrDuCellId=65535)."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, DmMimoSwitch=DM_MIMO_SERVICE_SWITCH-1;",
                  "Enable Distributed MIMO service master."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, DmMimoSwitch=DM_MIMO_PDCCH_TRP_TRANS_SW-1;",
                  "Optional: enable PDCCH TRP transmission."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, DmMimoSwitch=DM_MIMO_TRP_SEL_WEIGHT_SW-1;",
                  "Optional: enable TRP selection weights."),
             ]),
        dict(sid="S15-22", family="Step9", jump=S11,
             title="Fusion Cell / Virtual 128T — two 64T TRPs, one cell",
             doc="Fusion Cell FPD Ch.4  FBFD-091101",
             principal="Fuse two 64T TRPs into one cell (virtual 128T). ClusterType=INTRA_CELL_MIMO. Enhancements: MUMIMO_SINR_ENH_SW, SINGLE_TRP_SSB_TRANS_SW, FUSION_CALIB_INTRF_AVOID_SW. Service-affecting ADD/ACT. License by TRP count (two TRPs → two units). Do not mix with DAS blindly.",
             benefit="Higher-order MU-MIMO SINR and virtual 128T capacity where two 64T AAUs share one cell.",
             params="ADD GNBCLUSTER ClusterType=INTRA_CELL_MIMO; FusionCellAlgoSwitch MUMIMO_SINR_ENH_SW / SINGLE_TRP_SSB_TRANS_SW; GNODEBALGO ChnCalibPolSw=FUSION_CALIB_INTRF_AVOID_SW.",
             mmls=[
                 ("ADD GNBCLUSTER: ClusterId=90, ClusterType=INTRA_CELL_MIMO;",
                  "Create Fusion cluster (FPD sample ClusterId=90 — replace)."),
                 ('ADD GNBMIMOCLUSTERCELL: ClusterId=90, Mcc="302", Mnc="220", gNodeBId=1, CellId=90;',
                  "Bind cell to cluster (replace PLMN/gNodeBId/CellId)."),
                 (f"MOD NRDUCELLMULTITRP: NrDuCellId={p}, FusionCellAlgoSwitch=MUMIMO_SINR_ENH_SW-1;",
                  "Enable Fusion Cell MU-MIMO SINR enhancement."),
                 ("MOD GNODEBALGO: ChnCalibPolSw=FUSION_CALIB_INTRF_AVOID_SW-1;",
                  "Enable Fusion calibration interference avoid (gNodeB-level)."),
                 (f"MOD NRDUCELLMULTITRP: NrDuCellId={p}, FusionCellAlgoSwitch=SINGLE_TRP_SSB_TRANS_SW-1;",
                  "Optional: enable single-TRP SSB transmission."),
             ]),
        dict(sid="S15-23", family="Step10", jump=S12,
             title="mmWave (FR2) — beams, 3D coverage, MU-MIMO, multi-beam FDM",
             doc="mmWave Beam FPD + MIMO TDD Ch.8–9  FOFD-030201 / FOFD-010010",
             principal="Parallel FR2 track — not a substitute for FR1 Steps 1–9. FR2 uses PMI-based weights only (no SRS-based DL weights). Sequence: basic mmWave beams → 3D coverage / dense / TA / dynamic beam → mmWave MU-MIMO → multi-beam FDM. Cell DEA/ACT around NRDUCELLTRPMMWAVBEAM. Not applicable on n41 FR1 40 MHz — listed so the FPD line is not missed.",
             benefit="FR2 coverage + UL MU + volume-based beam multiplex. Skip on this n41 network.",
             params="CoverageScenario on NRDUCELLTRPMMWAVBEAM; FLEXIBLE_DENSE_BEAM_SW; DYNAMIC_BEAM_ALLOC_SW; MuMimoSwitch=UL_MU_MIMO_SW; BeamMultiplexSwitch=VOL_BASED_BEAM_MULTIPLEX_SW. License NR0SMMW3DCP0 / NR0SMMWULE00.",
             mmls=[
                 (f"MOD NRDUCELLTRPMMWAVBEAM: NrDuCellTrpId={t}, CoverageScenario=DEFAULT, Tilt=255;",
                  "FR2 basic beams (FPD sample). DEA NRCELL before this MOD, ACT after. N/A on n41 FR1."),
                 (f"MOD NRDUCELLTRPMMWAVBEAM: NrDuCellTrpId={t}, CoverageScenario=SCENARIO_101, Tilt=255;",
                  "FR2 3D coverage scenario 101. Consumes NR0SMMW3DCP0."),
                 (f"MOD NRCELLALGOSWITCH: NrCellId={p}, MeasPolicySwitch=SSB_MEAS_POS_POLICY_SW-1;",
                  "Enable SSB measurement-position policy (mmWave 3D)."),
                 (f"MOD NRDUCELLTRPMMWAVBEAM: NrDuCellTrpId={t}, BeamPerformanceSw=FLEXIBLE_DENSE_BEAM_SW-1;",
                  "Enable flexible dense beam."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, ShortTaCommandPeriodSW=SHORT_TAC_PERIOD-1;",
                  "Enable short TA command period."),
                 (f"MOD NRDUCELLMOBILEBHALGO: NrDuCellId={p}, MobileBackhaulAlgoSw=DYNAMIC_BEAM_ALLOC_SW-1;",
                  "Enable dynamic beam allocation."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, MuMimoSwitch=UL_MU_MIMO_SW-1;",
                  "Enable mmWave UL MU-MIMO (MIMO TDD Ch.8)."),
                 (f"MOD NRDUCELLALGOSWITCH: NrDuCellId={p}, BeamMultiplexSwitch=VOL_BASED_BEAM_MULTIPLEX_SW-1;",
                  "Enable volume-based beam multiplex (FR2 multi-beam FDM, Ch.9). Do not send on FR1 n41."),
             ]),
        dict(sid="S15-24", family="Step11", jump=S13,
             title="Inter-cell cable sequence detection — commissioning (MIMO gain depends on it)",
             doc="MIMO TDD Ch.10  FBFD-010025",
             principal="Detects crossed antenna cables between intra-gNodeB intra-frequency 4T4R NORMAL_CELL cells. Polarization of physical vs logical ports must match (Ch.4.3.4) or MIMO collapses. Run after RF is up; run again if MIMO KPI is poor and VSWR/DAS is not the cause. No license. Inaccurate in indoor DAS, VSWR alarms, or NI > −90 dBm.",
             benefit="Finds crossed feeders that silently destroy MIMO. Result CROSSED vs CORRECT. Not a throughput switch — a prerequisite of every MIMO KPI.",
             params="STR ANTENNAPORTOPTDET AntPortOptDetPolicy=INTER_CELL_DETECT; DSP ANTENNAPORTOPTDET; STP to stop. Constraint: TxRxMode=4T4R, NORMAL_CELL, ≥2 intra-freq cells ON.",
             mmls=[
                 ("STR ANTENNAPORTOPTDET: AntPortOptDetPolicy=INTER_CELL_DETECT;",
                  "Start inter-cell cable sequence detection for the gNodeB (4T4R NORMAL_CELL)."),
                 ("DSP ANTENNAPORTOPTDET;",
                  "Read result: CROSSED / CORRECT / ERROR_*. Fix jumpers if CROSSED, then re-run STR."),
                 ("STP ANTENNAPORTOPTDET: AntPortOptDetPolicy=INTER_CELL_DETECT;",
                  "Stop detection (document deactivation)."),
             ]),
        dict(sid="S15-25", family="Step3", jump=S5,
             title="Optional DL scheduling extras (freq-sel / load-based / smart sch)",
             doc="iBeam / MIMO TDD scheduling chapters + NRDUCellFeatureSw.PerformanceAlgoSwitch",
             principal="Beyond iBeam 1.0 minimum: FREQ_SEL_SCH_SW and LOAD_BASED_DL_EXP_SCH_SW on loaded cells; PerformanceAlgoSwitch SMART_SCH_AND_LINK_ADAPT_SW + HIGH_CAPACITY_EXP_IMP_SW (license + TAC). DlBeamPriBasedSch only after SSB adapt is trusted.",
             benefit="Loaded-hour and frequency-selective DL user tput on top of EPF + heavy-load pri.",
             params="NRDUCellDlSch.DlSchAlgoSwitch FREQ_SEL_SCH_SW / LOAD_BASED_DL_EXP_SCH_SW; NRDUCellFeatureSw.PerformanceAlgoSwitch SMART_SCH_AND_LINK_ADAPT_SW / HIGH_CAPACITY_EXP_IMP_SW.",
             mmls=[
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=FREQ_SEL_SCH_SW-1;",
                  "Enable frequency-selective DL scheduling (after CSI quality is stable)."),
                 (f"MOD NRDUCELLDLSCH: NrDuCellId={p}, DlSchAlgoSwitch=LOAD_BASED_DL_EXP_SCH_SW-1;",
                  "Enable load-based DL experience scheduling (busy-hour 32T)."),
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, PerformanceAlgoSwitch=SMART_SCH_AND_LINK_ADAPT_SW-1;",
                  "Enable smart scheduling and link adaptation (confirm performance-algorithm license)."),
                 (f"MOD NRDUCELLFEATURESW: NrDuCellId={p}, PerformanceAlgoSwitch=HIGH_CAPACITY_EXP_IMP_SW-1;",
                  "Enable high-capacity experience improve (confirm license + Huawei TAC before cluster)."),
             ]),
    ]


def add_box(ws, r, rec, sheet_name):
    start = r
    fam = rec["family"]
    color = FAMILY_COLOR.get(fam, NAVY)
    # title
    merge(ws, r, 1, r, 8)
    put(ws, r, 1, f"  {rec['sid']}   ·   {rec['title']}", size=12, bold=True, color=WHITE,
        fill_hex=color, h="left", v="center")
    for c in range(2, 9):
        ws.cell(r, c).fill = fill(color)
    put(ws, r, 9, "Read benefit →", size=9, bold=True, color=WHITE, fill_hex=color, h="right", v="center")
    href_sheet(ws.cell(r, 10), rec["jump"], rec["jump"])
    ws.cell(r, 10).fill = fill("FFF2CC")
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
    # MML header
    titles = ["MML #", "MML Command  (one parameter / one switch / one line)  +  note at end"] + [""] * 6 + ["Jump"] + [""]
    for i, t in enumerate(titles, 1):
        put(ws, r, i, t, size=9, bold=True, color="000000", fill_hex=YELLOW_HDR if i > 1 else BLUE_HDR,
            h="center", v="center", border=True)
    merge(ws, r, 2, r, 8)
    merge(ws, r, 9, r, 10)
    ws.row_dimensions[r].height = 22
    r += 1
    for i, (cmd, note) in enumerate(rec["mmls"], 1):
        line = mml_line(cmd, note)
        fh = WHITE if i % 2 else ROW_ALT
        put(ws, r, 1, i, size=9, bold=True, fill_hex=fh, h="center", v="top", border=True)
        merge(ws, r, 2, r, 8)
        put(ws, r, 2, line, size=9, fill_hex=fh, h="left", v="top", border=True)
        for c in range(3, 9):
            ws.cell(r, c).fill = fill(fh)
            ws.cell(r, c).border = thin
        merge(ws, r, 9, r, 10)
        href_sheet(ws.cell(r, 9), rec["jump"], rec["jump"])
        ws.cell(r, 9).fill = fill("FFF2CC")
        ws.cell(r, 10).fill = fill("FFF2CC")
        ws.row_dimensions[r].height = min(52, max(22, 16 + len(line) // 110 * 12))
        r += 1
    end = r - 1
    box_border(ws, start, end)
    r = blank(ws, r, 10)
    return r, start


BLUE_HDR = "5B9BD5"
YELLOW_HDR = "FFC000"


def patch_cover(wb):
    ws = wb["0. Cover & Index"]
    ws["A1"].value = "  5G MIMO (all features together)  —  Deployment Workbook  v3.0"
    last = ws.max_row + 2
    r = last
    r = section(ws, r, 10, "v3.0 addition — document suggestions (NEW last-but-check sheet 15)")
    r = note_bar(ws, r, 10,
                 "Sheet 15 lists every FPD feature that can still be enabled to improve massive MIMO "
                 "(Principal / Benefit / Parameter details / MML one-switch-one-line). "
                 "Hyperlinks jump to sheets 0–13 to read the full benefit. This is document-based — not the dump comparison (sheet 14). "
                 "File: MIMO_Deployment_v3.0.xlsx")
    r = headers(ws, r, ["#", "Sheet", "Maps to", "What you will find"] + [""] * 6)
    r = table_row(ws, r,
                  ["15", SHEET_NAME, "All 8 FPDs (sheets 0–13)",
                   "Box suggestions: Principal, Benefit, Parameter, MML + notes + hyperlinks"] + [""] * 6,
                  fills=[PALE_GOLD] * 10, height=36)
    merge(ws, r - 1, 4, r - 1, 10)
    return r


def build_sheet(wb):
    ws = wb.create_sheet(SHEET_NAME)
    setup_sheet(ws, SHEET_NAME)
    set_widths(ws, WIDTHS)
    ws.oddHeader.left.text = "5G MIMO Suggestions from FPD (v3.0) — Principal / Benefit / Parameter / MML"
    ws.oddFooter.left.text = "Source: eight Huawei 5G RAN10.1 FPDs in this workbook (sheets 0–13). One switch one line. Click Jump."
    ws.freeze_panes = "A4"
    ws.sheet_properties.tabColor = GOLD

    r = 1
    r = banner(ws, r, COLS,
               "  15.  MIMO Suggestions from uploaded documents  —  features / switches to enable further",
               fill_hex="C9A227", size=16, height=30)
    r = note_bar(ws, r, COLS,
                 "Read every FPD activation line already extracted in sheets 0–13. This sheet is NOT a live-dump check "
                 "(that is sheet 14). Goal: which feature (parameter or switch) can still be enabled to improve massive MIMO "
                 "(beam, scheduling, pairing, AHR, iBeam, UL boosting). Each box = Principal + Benefit + Parameter details + "
                 "MML (one parameter / one switch / one line, note at end of command). Yellow Jump cells hyperlink to the "
                 "source step sheet so you can read the full benefit. Replace {NrDuCellId} / {NrDuCellTrpId}. "
                 "Keep 2T2R indoor OFF. mmWave/DAS/Fusion boxes are site-architecture / FR2 — skip if not deployed.")

    r = section(ws, r, COLS, "Jump to workbook sheets  (click)")
    sheets = [S0, S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14]
    # two rows of links
    row_sheets = [sheets[:8], sheets[8:]]
    for pack in row_sheets:
        for i, name in enumerate(pack, 1):
            href_sheet(ws.cell(r, i), name, name.split(". ", 1)[-1][:22])
            ws.cell(r, i).fill = fill(PALE_BLUE)
            ws.cell(r, i).border = thin
        ws.row_dimensions[r].height = 20
        r += 1
    r = blank(ws, r, 8)

    r = section(ws, r, COLS, "Index of suggestion boxes  (click ID to jump down this sheet)")
    toc_header_row = r
    r = headers(ws, r, ["ID", "Family", "Suggestion (click ID)", "Source sheet (click)"] + [""] * 6)
    toc_start = r
    items = suggestions()
    # placeholder rows — fill after boxes know their start rows
    for rec in items:
        put(ws, r, 1, rec["sid"], size=10, bold=True, fill_hex=PALE_GOLD, h="center", v="center", border=True)
        put(ws, r, 2, rec["family"], size=10, fill_hex=WHITE, h="center", v="center", border=True)
        put(ws, r, 3, rec["title"][:80], size=9, fill_hex=WHITE, h="left", v="center", border=True)
        merge(ws, r, 3, r, 8)
        href_sheet(ws.cell(r, 9), rec["jump"], rec["jump"])
        merge(ws, r, 9, r, 10)
        ws.row_dimensions[r].height = 20
        r += 1
    toc_end = r - 1
    r = blank(ws, r, 12)

    box_rows = {}
    r = section(ws, r, COLS, "Suggestion boxes  (Principal · Benefit · Parameter details · MML)")
    for rec in items:
        r, start = add_box(ws, r, rec, SHEET_NAME)
        box_rows[rec["sid"]] = start

    # patch TOC hyperlinks to box title rows
    for i, rec in enumerate(items):
        rr = toc_start + i
        href_row(ws.cell(rr, 1), SHEET_NAME, box_rows[rec["sid"]], rec["sid"])
        ws.cell(rr, 1).fill = fill(PALE_GOLD)
        href_row(ws.cell(rr, 3), SHEET_NAME, box_rows[rec["sid"]], rec["title"][:90])

    return ws


def main():
    src = SRC
    if not os.path.exists(src):
        src = os.path.join(ROOT, "MIMO_Deployment.xlsx")
    if not os.path.exists(src):
        raise SystemExit("missing v2.0 / v1 workbook")
    print("copy", src, "→", OUT)
    shutil.copy2(src, OUT)
    wb = load_workbook(OUT)
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    print("patch cover...")
    patch_cover(wb)
    print("sheet 15...")
    build_sheet(wb)
    print("saving", OUT)
    wb.save(OUT)
    print("ok", os.path.getsize(OUT), "sheets", len(wb.worksheets), "last", wb.sheetnames[-1])
    with zipfile.ZipFile(ZIP_OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(OUT, os.path.basename(OUT))
    print("zip", ZIP_OUT, os.path.getsize(ZIP_OUT))


if __name__ == "__main__":
    main()
