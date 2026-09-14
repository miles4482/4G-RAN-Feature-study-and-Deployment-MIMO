"""
MIMO_Deployment v10.0 — Action plan rebuilt for TWO equal targets:

  1. Raise N.ChMeas.MIMO.DL.MuPairing  (all-layer and ≥4-layer paired RB)
  2. Raise User DL Average Throughput  (N.ThpVol.DL / N.RLC.ThpTime.DL.Cell)

v9 parked the DL-tput levers. This sheet puts them back, in the only order that
serves both KPIs: recover the SRS that pairing AND weights need (P0), land the
four beam bits Ph1 never activated (P1), then spend LAYER_16 to lift pairing
above the 9 Sep baseline (P2).

A night that raises one KPI by lowering the other is a fail.

Run:  python3 build_v10_dual_plan.py
"""
import os

from mimo_excel_style import *
import build_combined_sheet as bcs
from build_combined_sheet import href_sheet
import build_v7_findings as v7
import build_v8_rollback as v8
import build_v9_pairing_plan as v9

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "MIMO_Deployment_v10.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v10.0.zip")

D_SHEET = bcs.DUAL_PLAN_SHEET
PID = "{NrDuCellId}"
COLS = v7.COLS
W = [8, 16, 28, 32, 38, 22, 22, 24, 22, 16, 18]
RED_HDR = v7.RED_HDR

ML = "MMIMO_MULTILAYER_ENHANCE_SW-1"
PREF = "MU_MIMO_PAIRING_PREFERRED_SW-1"
RANK = "MU_RANK_BOOSTING_SW-1"
SRSACC = "SRS_MEAS_ACCELERATING_SW-1"

PHASES = [
    ("P0", "NEXT\n(tonight)",
     "Recover pairing. Undo the Ph1 bits that reduced MuPairing. Add nothing that is supposed to raise tput yet — "
     "SRS_WEIGHT_ESTIMATE on top of a bad SRS estimate would bake the damage into DL weights.",
     "RB-1: SRS_IC_SW-0 alone. If 24 h later recover-gate missed: RB-2 PRECISE-0 (keep ANTI_INTRF), then RB-3 ANTI_INTRF-0. "
     "Full MML on sheet 18.",
     "LST NRDUCELLSRS (tight MUX still 1, IC still 1). Neighbour 32T control identified. Pull current DL tput in Mbit/s.",
     "PAIRING: all-layer ≥ 14.7 and ≥4L ≥ 8.8.  TPUT: no drop vs neighbour 32T control.  "
     "SRS NI back to 5–9 Sep. IBLER / drop unchanged.",
     "Restore that night’s bit only (sheet 18). Do not touch HighPrecisionBeamSwitch.",
     "NEXT"),
    ("P1", "P0 gate met\nthen 2 nights",
     "Land the DL throughput levers Ph1 never activated. This is the first real test of the DL-tput hypothesis, "
     "and SRS-based weights also feed a cleaner spatial estimate into pairing.",
     "Night 1: SUPER_COVERAGE_SW-1, then SRS_WEIGHT_ESTIMATE_SW-1, BEAM_SELECT_OPT_SW-1 — one MML each, LST in between. "
     "Night 2: BEAM_TRACKING_SW-1 alone (needs SUPER_COVERAGE already committed), then INTELLIGENT_BEAM_SELECTION "
     "with tracking listed, then AggLvlComprCceUsageThld=60.",
     "P0 pairing recover-gate already met. BeamPerceiveMode=DISTRIBUTED_MODE. HighPrecisionBeamSwitch=ON. "
     "LST NRDUCELLCHNCOVALGO before night 2.",
     "TPUT: MCS upshift and User DL tput ≥ control (fill the Mbit/s from MAE).  "
     "PAIRING: must not fall vs post-P0.  Drop/HOSR unchanged after SUPER_COVERAGE and tracking.",
     "Reverse night 2 first (intelligent + tracking off, then select-opt, then weights, then SUPER_COVERAGE back). "
     "Never touch HighPrecisionBeamSwitch.",
     "PLANNED"),
    ("P2", "P1 gate met\nthen 3–4 nights",
     "Raise pairing ABOVE the 9 Sep baseline by spending the live LAYER_16 quota. This is also a DL cell-capacity "
     "lever — more paired layers should lift tput if IBLER holds.",
     "Night 1: HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1 only. Then one child per night: "
     "PAIRING_PREFERRED, RANK_BOOSTING, SRS_MEAS_ACCELERATING. Always list every HighLayer bit that must stay ON.",
     "P1 green. LST LICENSE NR0S0DLEPU00. DL_MU_MIMO_SW-1. MaxMimoLayerNum already LAYER_16 — never downgrade it.",
     "PAIRING: all-layer > 16.4 and ≥4L > 9.8, and 8–16 layer RB leaves zero.  "
     "TPUT: ≥ post-P1.  Transmission.Layer.Max rises. IBLER inside band.",
     "Children back to 0 first (newest first), master last.",
     "PLANNED"),
    ("P3", "P2 gate met\n+ 24 h",
     "Hold healthy MU pairs through the busy hour (fewer SU fallbacks) without giving tput back.",
     "MOD NRDUCELLDLMIMO DlMuBackToSuSeThld=0. Live value is 5. One parameter, one night.",
     "P2 both-KPI gate green. Confirm live value is 5 with LST.",
     "PAIRING: Pair.Layer.Avg holds higher through BH.  TPUT: ≥ post-P2.  IBLER inside band.",
     "MOD NRDUCELLDLMIMO: DlMuBackToSuSeThld=5;",
     "PLANNED"),
    ("P4", "Only if a KPI is still short",
     "One optional night, not a bundle. Choose the lever for the KPI that is still short — never both in one window.",
     "If PAIRING still short in the loaded BH: AHR_CAPC_UPGRADE_PHASE2_SW-1 (sheet 19 P4). "
     "If TPUT still short at cell edge and Tilt ≠ 255: SSB_BEAM_ADAPT_SW-1 then VERTICAL_COV_IMP the next night.",
     "P3 green. For SSB: no Tilt=255 cell in the trial set. For AHR CU: Phase1+Turbo already ON, licence NR0S00ACT200.",
     "The short KPI moves; the other KPI does not fall. Drop/HOSR unchanged for SSB. CCE not blocking for AHR CU.",
     "That night’s switch only.",
     "OPTIONAL"),
    ("P5", "Both KPIs green\n+ 7 days",
     "Cluster rollout of the proven dual set only.",
     "Same MML sequence per remaining 32T cell: P0.1, then P1, then P2, then P3. Fix Tilt=255 before a cell joins. "
     "2T2R indoor stays out. Do not put SRS_IC or the two MU quality gates back.",
     "Per-cell LST of SrsDetectionAlgoSwitch, DlCoverageAlgoSwitch, WeightAlgoSwitch, HighLayerMuMimoSw, MaxMimoLayerNum.",
     "Each gNB: MuPairing ≥ its own pre-week AND ≥ trial post-P2; DL tput ≥ its own pre-week AND ≥ trial post-P1.",
     "Per-gNB reverse of that gNB’s CME file only.",
     "PLANNED"),
]

SEQ = [
    ("P0.1", "Tonight\nalone",
     "SrsDetectionAlgoSwitch=SRS_IC_SW-0",
     f"LST NRDUCELLSRS: NrDuCellId={PID};\n"
     f"MOD NRDUCELLSRS: NrDuCellId={PID}, SrsDetectionAlgoSwitch=SRS_IC_SW-0;\n"
     f"LST NRDUCELLSRS: NrDuCellId={PID};",
     "Rollback. Keep SRS_TIGHT_MULTIPLEXING_SW-1. Do not start P1 tonight.",
     "Tight MUX and SRS_SINR_MEAS_OPT still 1. RETCODE = 0.",
     "MuPairing.kLayer.RB, N.SRS.NI.Avg, User DL tput",
     "After 24 h: pairing ≥ 14.7 / 8.8 AND tput not down vs control. If YES → P1. If pairing still short → P0.2."),
    ("P0.2", "Only if P0.1\nmissed pairing",
     "DL_MU_PRECISE_SCH_SW-0, keep ANTI_INTRF",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
     "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-1;",
     "Bit-field: list ANTI_INTRF as -1. Watch IBLER — this can raise pairing and IBLER together.",
     "PRECISE gone, ANTI_INTRF still 1.",
     "MuPairing.kLayer.RB, DL IBLER, User DL tput",
     "Same pairing recover-gate. Tput must not drop. IBLER jump → restore PRECISE."),
    ("P0.3", "Only if P0.2\nmissed pairing",
     "DL_MU_ANTI_INTRF_SCH_SW-0",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
     "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-0;",
     "List both bits as -0 so PRECISE is not restored by omission.",
     "Neither PRECISE nor ANTI_INTRF selected.",
     "MuPairing.kLayer.RB, DL IBLER, User DL tput",
     "Then go to P1 even if pairing is still a little short — remaining lift is multilayer, not more rollback."),
    ("P1.1", "P0 green\nnight 1 first",
     "DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1",
     f"LST NRDUCELLCHNCOVALGO: NrDuCellId={PID};\n"
     f"MOD NRDUCELLCHNCOVALGO: NrDuCellId={PID}, DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1;\n"
     f"LST NRDUCELLCHNCOVALGO: NrDuCellId={PID};",
     "THE missing pre-requisite of BEAM_TRACKING_SW. Without this line P1.4 returns RETCODE 2147616329.",
     "DlCoverageAlgoSwitch shows SUPER_COVERAGE_SW-1. RETCODE = 0.",
     "Drop, HOSR, User DL tput, MuPairing",
     "Wait 30 min for drop/HO before P1.2. Coverage change must not raise drop."),
    ("P1.2", "After P1.1\nsame night OK",
     "WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1",
     f"LST NRDUCELLBEAMALGO: NrDuCellId={PID};   // BeamPerceiveMode=DISTRIBUTED_MODE\n"
     f"MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1;\n"
     f"LST NRDUCELLBEAMALGO: NrDuCellId={PID};",
     "MAIN DL-tput lever, and it feeds SRS-based spatial isolation into pairing. Own MML line — Ph1 lost it in a bundle.",
     "WeightAlgoSwitch shows SRS_WEIGHT_ESTIMATE_SW-1. RETCODE = 0.",
     "N.ChMeas.PDSCH.MCS.k, User DL tput, MuPairing, DL IBLER",
     "MCS should upshift. Pairing must not fall (better SRS weights usually help it). IBLER gate."),
    ("P1.3", "After P1.2\nsame night OK",
     "ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1",
     f"MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1;",
     "HighPrecisionBeamSwitch=ON (already). Own line. If this fails, do not send P1.4.",
     "ChannelOptAlgoSwitch shows BEAM_SELECT_OPT_SW-1.",
     "User DL tput, MCS, MuPairing",
     "Tput holds or rises. Pairing holds."),
    ("P1.4", "Night 2\nALONE first",
     "BeamOptAlgoSwitch=BEAM_TRACKING_SW-1",
     f"LST NRDUCELLCHNCOVALGO: NrDuCellId={PID};   // SUPER_COVERAGE must already be 1\n"
     f"MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, BeamOptAlgoSwitch=BEAM_TRACKING_SW-1;\n"
     f"LST NRDUCELLBEAMALGO: NrDuCellId={PID};",
     "The line that failed on 10 Sep. Send ALONE. SUPER_COVERAGE must already be committed (P1.1).",
     "BeamOptAlgoSwitch shows BEAM_TRACKING_SW-1. RETCODE = 0 — not 2147616329.",
     "MCS on mobility samples, User DL tput, HOSR, drop, MuPairing",
     "This is the mobility tput lever. Pairing should hold. Drop/HO rise → tracking off that night."),
    ("P1.5", "After P1.4\nRETCODE = 0",
     "BEAM_TRACKING_SW-1&INTELLIGENT_BEAM_SELECTION_SW-1",
     f"MOD NRDUCELLBEAMALGO: NrDuCellId={PID}, "
     "BeamOptAlgoSwitch=BEAM_TRACKING_SW-1&INTELLIGENT_BEAM_SELECTION_SW-1;",
     "Bit-field: list tracking as -1 so it stays ON while adding intelligent selection.",
     "Both options present in BeamOptAlgoSwitch.",
     "User DL tput, MCS, MuPairing, DL IBLER",
     "P1 dual gate after 24 h: tput ≥ control with MCS upshift, pairing ≥ post-P0."),
    ("P1.6", "End of night 2",
     "AggLvlComprCceUsageThld=60",
     f"MOD NRDUCELLPDCCHALGO: NrDuCellId={PID}, AggLvlComprCceUsageThld=60;",
     "PDCCH_AGG_LVL_COMPR_SW already 1 from Ph1; the threshold was never sent. Frees CCE for PDSCH — tput helper.",
     "AggLvlComprCceUsageThld = 60.",
     "N.CCE.Used.Avg, User DL tput, MuPairing",
     "CCE used falls or holds. Blocking up = fail (hurts both KPIs)."),
    ("P2.1", "P1 green\nnight 1",
     f"HighLayerMuMimoSw={ML}",
     f"LST LICENSE;\n"
     f"LST NRDUCELLPDSCH: NrDuCellId={PID};   // MaxMimoLayerNum = LAYER_16\n"
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML};\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "Master. Lets the cell spend LAYER_16. Without it high-layer pairing cannot rise. One bit this night.",
     "HighLayerMuMimoSw shows MMIMO_MULTILAYER_ENHANCE_SW-1. RETCODE = 0.",
     "Transmission.Layer.Max, MuPairing.8–16Layer.RB, ≥4L RB, User DL tput, DL IBLER",
     "Layer.Max up and 8–16 layer RB leaves 0. Tput ≥ post-P1. If Layer.Max stays flat, check licence — do not send P2.2."),
    ("P2.2", "P2.1 green\nnight 2",
     f"{ML}&{PREF}",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML}&{PREF};",
     "Largest pairing lever of the children. Master must stay listed as -1 or it drops.",
     "Both MMIMO_MULTILAYER_ENHANCE and MU_MIMO_PAIRING_PREFERRED selected.",
     "MuPairing.kLayer.RB, Pair.PRB, User DL tput, DL IBLER",
     "Largest expected pairing step-up. Tput must hold or rise. IBLER is the hard gate."),
    ("P2.3", "P2.2 green\nnight 3",
     f"{ML}&{PREF}&{RANK}",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML}&{PREF}&{RANK};",
     "More layers inside an already-paired group. List every bit that must stay ON.",
     "All three HighLayer bits selected.",
     "MuPairing.4–8Layer.RB, N.PDSCH.InitTbDl.Rank*, User DL tput, DL IBLER",
     "4–8 layer RB and rank mix up. Tput ≥ post-P2.2."),
    ("P2.4", "P2.3 green\nnight 4",
     f"{ML}&{PREF}&{RANK}&{SRSACC}",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML}&{PREF}&{RANK}&{SRSACC};",
     "Faster SRS so pairing decisions stay fresh. Watch SRS NI.",
     "Four HighLayer bits selected.",
     "MuPairing, N.UL.SRS.PreSINR, N.SRS.NI.Avg, User DL tput",
     "P2 dual gate: pairing > 9 Sep AND tput ≥ post-P1. NI must not jump."),
    ("P3.1", "P2 green\n1 night",
     "DlMuBackToSuSeThld=0  (live is 5)",
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};\n"
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, DlMuBackToSuSeThld=0;",
     "FPD multilayer sample. Lowers the SE test that kicks a MU pair back to SU.",
     "DlMuBackToSuSeThld = 0.",
     "Pair.Layer.Avg through BH, User DL tput, DL IBLER",
     "Pairs last longer. Tput holds. IBLER jump → set back to 5."),
]

LEVERS = [
    ("RECOVER pairing", "SRS_IC_SW-0", "P0.1",
     "Removes the SRS-detection conflict. Restores the SRS that BOTH pairing and SRS-weights need."),
    ("RECOVER pairing", "DL_MU_PRECISE_SCH / ANTI_INTRF → 0", "P0.2 / P0.3",
     "Quality gates Ph1 added. They cut paired RB by design. Only if P0.1 missed."),
    ("RAISE tput (+ pairing)", "SRS_WEIGHT_ESTIMATE_SW", "P1.2",
     "Main DL MCS / user-tput lever. Also a cleaner spatial estimate for MU pairing. Never landed in Ph1."),
    ("RAISE tput", "BEAM_TRACKING_SW  (after SUPER_COVERAGE)", "P1.1 then P1.4",
     "Mobility MCS / tput. SUPER_COVERAGE first or RETCODE 2147616329 repeats."),
    ("RAISE tput", "BEAM_SELECT_OPT / INTELLIGENT_BEAM_SELECTION / AggLvl=60", "P1.3 / P1.5 / P1.6",
     "Serving-beam pick + CCE headroom. Tput helpers, pairing-neutral if IBLER holds."),
    ("RAISE pairing (+ tput)", "MMIMO_MULTILAYER_ENHANCE_SW", "P2.1",
     "Master. Spends LAYER_16. High-layer MuPairing cannot rise without it. More layers also lift cell tput."),
    ("RAISE pairing", "MU_MIMO_PAIRING_PREFERRED_SW", "P2.2",
     "Largest pairing child. Prefers MU over SU when a pair has gain."),
    ("RAISE pairing", "MU_RANK_BOOSTING_SW / SRS_MEAS_ACCELERATING / BackToSu=0", "P2.3 / P2.4 / P3",
     "Deeper pairs, fresher SRS, fewer SU fallbacks."),
    ("REDUCE pairing", "SRS_IC + tight MUX together; PRECISE / ANTI_INTRF", "Ph1 (done)",
     "What dropped pairing. Stay off."),
]

PARK = [
    ("SRS_IC_SW-1",
     "Most likely cause of the Ph1 pairing collapse next to tight multiplexing. Do not put it back to chase SRS capacity.",
     "Do not re-enable"),
    ("DL_MU_PRECISE_SCH / DL_MU_ANTI_INTRF",
     "Quality gates. They reduce paired RB by design and can also cap tput of paired UEs. Keep OFF until both KPIs "
     "have been above target for 7 days, then trial one of them — not both — only if IBLER needs help.",
     "Off until both KPIs green + 7 d"),
    ("iBeam 2.0 / 3.0  (PRECISE_MUMIMO_EVAL, FAR_UE_RANK_OPT, DL_ROBUST_WEIGHT, DL_SMART_AMC)",
     "More pairing-quality / weight-robustness gates. Same direction as PRECISE_SCH. Not this campaign.",
     "Hold"),
    ("UL Boosting 2.0 / PMI_WEIGHT_OPT / OPEN_LOOP_WEIGHT_OPT",
     "UL KPI or an alternative weight source. Do not mix weight sources with the P1 SRS-weight baseline.",
     "Hold"),
    ("AHR Capacity Upgrade 2.0 / SSB_BEAM_ADAPT",
     "Optional P4 only if one KPI is still short after P3. Not in the default path.",
     "P4 optional only"),
]


def build_dual_plan(wb):
    ws = wb.create_sheet(D_SHEET, 20)
    setup_sheet(ws, D_SHEET)
    set_widths(ws, W)
    ws.oddHeader.left.text = "5G MIMO — Action Plan  (pairing + DL user throughput)"
    ws.oddFooter.left.text = FOOTER_LEFT
    ws.sheet_properties.tabColor = "1F4E79"

    r = banner(ws, 1, COLS,
               "  Action Plan  —  two equal targets:  N.ChMeas.MIMO.DL.MuPairing   and   User DL Average Throughput",
               fill_hex=NAVY)
    r = note_bar(ws, r, COLS,
                 "v9 parked the DL-tput levers. This plan puts them back, after pairing is recovered, because "
                 "SRS_WEIGHT_ESTIMATE and BEAM_TRACKING never reached the air in Ph1 (RETCODE 2147616329) — that is "
                 "why DL tput was flat. Sequence: P0 recover pairing → P1 land the rejected tput bits → P2 spend "
                 "LAYER_16 to raise pairing above 9 Sep. A night that raises one KPI by lowering the other is a fail. "
                 "Replace {NrDuCellId} with 101 / 102 / 103. One switch per MML line.",
                 fill_hex="FFF2CC")
    r = blank(ws, r, 6)

    r = section(ws, r, COLS, "Section 1.  Dual numeric target")
    r = v7._hdr(ws, r, ["#", "KPI", "Ph1 result", "P0 recover gate", "P1 success (tput levers)",
                        "P2 success (multilayer)", "How we get there", "How (cont.)",
                        "Dual-fail rule", "Safety gate", "Safety (cont.)"],
                fill_hex=TEAL, height=36)
    dual = [
        (1, "N.ChMeas.MIMO.DL.MuPairing  all-layer RB",
         "16.4 → 6.4 (−61 %) on 11 Sep",
         "≥ 14.7  (90 % of 9 Sep)",
         "must not fall vs post-P0",
         "> 16.4  and 8–16L leaves 0",
         "P0 rollback, then P2.1 master + P2.2 PAIRING_PREFERRED",
         "Pairing↑ + tput↓ = fail that night (over-pairing / IBLER)",
         "PRB load comparable; DL IBLER inside band vs control"),
        (2, "N.ChMeas.MIMO.DL.MuPairing  ≥4-layer RB",
         "9.8 → 3.3 (−66 %) on 11 Sep",
         "≥ 8.8",
         "must not fall vs post-P0",
         "> 9.8",
         "Same as row 1 — this is the part that collapsed",
         "Same dual-fail",
         "Transmission.Layer.Max must rise with P2"),
        (3, "User DL Average Throughput (DU)",
         "Reported flat — expected: weights/tracking never landed",
         "no drop vs neighbour 32T control",
         "MCS upshift and tput ≥ control (fill Mbit/s from MAE)",
         "tput ≥ post-P1",
         "P1.2 SRS_WEIGHT_ESTIMATE then P1.4 BEAM_TRACKING (SUPER_COVERAGE first)",
         "Tput↑ + pairing↓ = fail that night (SU-biased / weights not helping MU)",
         "Same BH, similar PRB load, IBLER / drop unchanged"),
    ]
    for rec in dual:
        i = rec[0]
        base = alt_fill(i)
        kpi_f = "D6EAF8" if i < 3 else "FFF2CC"
        put(ws, r, 1, rec[0], size=9, bold=True, fill_hex=base, h="center", v="center", border=True)
        put(ws, r, 2, rec[1], size=8, bold=True, fill_hex=kpi_f, h="left", v="top", border=True)
        put(ws, r, 3, rec[2], size=8, fill_hex="F8CBAD", h="left", v="top", border=True)
        put(ws, r, 4, rec[3], size=8, bold=True, fill_hex="FFF2CC", h="left", v="top", border=True)
        put(ws, r, 5, rec[4], size=8, bold=True, fill_hex=PALE_GREEN, h="left", v="top", border=True)
        put(ws, r, 6, rec[5], size=8, bold=True, fill_hex=PALE_GREEN, h="left", v="top", border=True)
        merge(ws, r, 7, r, 8)
        put(ws, r, 7, rec[6], size=8, fill_hex="D6EAF8", h="left", v="top", border=True)
        ws.cell(r, 8).border = thin
        put(ws, r, 9, rec[7], size=8, fill_hex="FCE4D6", h="left", v="top", border=True)
        merge(ws, r, 10, r, 11)
        put(ws, r, 10, rec[8], size=8, fill_hex="FFF2CC", h="left", v="top", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(W[1], rec[1]), (W[2], rec[2]), (v7._w(W, 7, 8), rec[6]),
             (W[8], rec[7]), (v7._w(W, 10, 11), rec[8])], cap=88)
        r += 1
    r = note_bar(ws, r, COLS,
                 "Pairing figures are approximate (read off the Ph1 pivot). Confirm from MAE. "
                 "DL tput in Mbit/s was not supplied — pull 5 / 9 / 10 / 11 / 12 Sep trial vs control "
                 "before P0 is signed off, otherwise P1 has nothing to be measured against.")
    r = blank(ws, r, 8)

    r = section(ws, r, COLS, "Section 2.  What raises which KPI")
    r = v7._hdr(ws, r, ["#", "Serves", "Switch / parameter", "Switch (cont.)", "Phase",
                        "Why it is in this plan", "Why (cont.)", "Why (cont.)", "Why (cont.)",
                        "Send?", "Send?"], fill_hex=NAVY, height=28)
    serve_fill = {
        "RECOVER pairing": "D6EAF8",
        "RAISE tput (+ pairing)": "FFF2CC",
        "RAISE tput": "FFF2CC",
        "RAISE pairing (+ tput)": "D5F5E3",
        "RAISE pairing": "D5F5E3",
        "REDUCE pairing": "F8CBAD",
    }
    for i, (d, sw, ph, fx) in enumerate(LEVERS, 1):
        df = serve_fill[d]
        put(ws, r, 1, i, size=9, bold=True, fill_hex=df, h="center", v="center", border=True)
        put(ws, r, 2, d, size=8, bold=True, fill_hex=df, h="left", v="center", border=True)
        merge(ws, r, 3, r, 4)
        put(ws, r, 3, sw, size=8, bold=True, fill_hex=df, h="left", v="top", border=True)
        ws.cell(r, 4).border = thin
        put(ws, r, 5, ph, size=8, bold=True, fill_hex=df, h="center", v="center", border=True)
        merge(ws, r, 6, r, 9)
        put(ws, r, 6, fx, size=8, fill_hex=WHITE, h="left", v="top", border=True)
        for c in (7, 8, 9):
            ws.cell(r, c).border = thin
        merge(ws, r, 10, r, 11)
        send = "NO — keep off" if d.startswith("REDUCE") else "YES, in sequence"
        put(ws, r, 10, send, size=8, bold=True,
            fill_hex="F8CBAD" if d.startswith("REDUCE") else PALE_GREEN,
            h="center", v="center", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(W[1], d), (v7._w(W, 3, 4), sw), (v7._w(W, 6, 9), fx)], cap=78)
        r += 1
    r = blank(ws, r, 8)

    r = section(ws, r, COLS, "Section 3.  Phase overview  (replaces the old Ph1–Ph7 order for this campaign)",
                fill_hex=GREEN)
    r = v7._hdr(ws, r, ["Phase", "Window", "Objective", "Content  (one switch per MML line)",
                        "Pre-requisite BEFORE the work order", "Pre-req (cont.)",
                        "Exit gate  (BOTH KPIs)", "Exit (cont.)",
                        "Rollback", "Status", "Status"], fill_hex=GREEN, height=40)
    st_fill = {"NEXT": "FFC000", "PLANNED": "EAECEE", "OPTIONAL": "D6EAF8"}
    ph_fill = {"NEXT": RED_HDR, "PLANNED": GREEN, "OPTIONAL": "5B9BD5"}
    for ph, win, obj, content, pre, gate, back, status in PHASES:
        base = st_fill[status]
        put(ws, r, 1, ph, size=12, bold=True, color=WHITE,
            fill_hex=ph_fill[status], h="center", v="center", border=True)
        put(ws, r, 2, win, size=8, bold=True, fill_hex=base, h="center", v="center", border=True)
        put(ws, r, 3, obj, size=8, bold=True, fill_hex=base, h="left", v="top", border=True)
        put(ws, r, 4, content, size=8, fill_hex=WHITE, h="left", v="top", border=True)
        merge(ws, r, 5, r, 6)
        put(ws, r, 5, pre, size=8, fill_hex="FCE4E4", h="left", v="top", border=True)
        ws.cell(r, 6).border = thin
        merge(ws, r, 7, r, 8)
        put(ws, r, 7, gate, size=8, fill_hex="FFF2CC", h="left", v="top", border=True)
        ws.cell(r, 8).border = thin
        put(ws, r, 9, back, size=8, fill_hex=PALE_BLUE, h="left", v="top", border=True)
        merge(ws, r, 10, r, 11)
        put(ws, r, 10, status, size=10, bold=True, fill_hex=base, h="center", v="center", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(W[2], obj), (W[3], content), (v7._w(W, 5, 6), pre),
             (v7._w(W, 7, 8), gate), (W[8], back)], cap=180)
        r += 1
    r = blank(ws, r, 8)

    r = section(ws, r, COLS, "Section 4.  Sequenced MML  —  copy one block per night", fill_hex=RED_HDR)
    r = note_bar(ws, r, COLS,
                 "P0.1 is the only change tonight. P1.4 is the line that failed on 10 Sep — SUPER_COVERAGE "
                 "must already be committed. HighLayerMuMimoSw is a bit-field: when adding PAIRING_PREFERRED "
                 "you must list MMIMO_MULTILAYER_ENHANCE_SW-1 in the same MOD or the master drops.",
                 fill_hex="F8CBAD")
    r = v7._hdr(ws, r, ["Seq", "When", "Switch / parameter", "MML  (one switch, LST → MOD → LST)",
                        "MML (cont.)", "Pre-requisite / bit-field rule", "Verify",
                        "Counter  (pairing + tput)", "Counter (cont.)",
                        "Stop / continue gate", "Stop / continue (cont.)"],
                fill_hex=RED_HDR, height=40)
    for i, (seq, when, sw, mml, pre, verify, ctr, gate) in enumerate(SEQ, 1):
        base = alt_fill(i)
        night = "FFC000" if seq == "P0.1" else base
        put(ws, r, 1, seq, size=10, bold=True, color=WHITE, fill_hex=RED_HDR,
            h="center", v="center", border=True)
        put(ws, r, 2, when, size=8, bold=True, fill_hex=night, h="center", v="top", border=True)
        put(ws, r, 3, sw, size=8, bold=True, fill_hex="FFF2CC", h="left", v="top", border=True)
        merge(ws, r, 4, r, 5)
        put(ws, r, 4, mml, size=8, bold=True, fill_hex="FCE4E4", h="left", v="top", border=True)
        ws.cell(r, 5).border = thin
        put(ws, r, 6, pre, size=8, fill_hex="FCE4E4", h="left", v="top", border=True)
        put(ws, r, 7, verify, size=8, fill_hex=PALE_BLUE, h="left", v="top", border=True)
        merge(ws, r, 8, r, 9)
        put(ws, r, 8, ctr, size=8, fill_hex="D5F5E3", h="left", v="top", border=True)
        ws.cell(r, 9).border = thin
        merge(ws, r, 10, r, 11)
        put(ws, r, 10, gate, size=8, fill_hex=base, h="left", v="top", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(W[2], sw), (v7._w(W, 4, 5), mml), (W[5], pre), (W[6], verify),
             (v7._w(W, 8, 9), ctr), (v7._w(W, 10, 11), gate)], cap=210)
        r += 1
    r = blank(ws, r, 8)

    r = section(ws, r, COLS, "Section 5.  Parked  —  not in the default path", fill_hex=GRAY)
    r = v7._hdr(ws, r, ["#", "Parked item", "Parked item (cont.)", "Parked item (cont.)",
                        "Why it is not in the default path", "Why (cont.)", "Why (cont.)", "Why (cont.)",
                        "When it may return", "When (cont.)", "When (cont.)"],
                fill_hex=GRAY, height=28)
    for i, (item, why, when) in enumerate(PARK, 1):
        base = alt_fill(i)
        put(ws, r, 1, i, size=9, bold=True, fill_hex=base, h="center", v="center", border=True)
        merge(ws, r, 2, r, 4)
        put(ws, r, 2, item, size=8, bold=True, fill_hex="FFF2CC", h="left", v="top", border=True)
        for c in (3, 4):
            ws.cell(r, c).border = thin
        merge(ws, r, 5, r, 8)
        put(ws, r, 5, why, size=8, fill_hex=WHITE, h="left", v="top", border=True)
        for c in (6, 7, 8):
            ws.cell(r, c).border = thin
        merge(ws, r, 9, r, 11)
        put(ws, r, 9, when, size=8, bold=True, fill_hex="FCE4D6", h="left", v="center", border=True)
        for c in (10, 11):
            ws.cell(r, c).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(v7._w(W, 2, 4), item), (v7._w(W, 5, 8), why)], cap=86)
        r += 1
    r = blank(ws, r, 8)

    r = subsection(ws, r, COLS, "Standing rules for both KPIs", fill_hex=NAVY2)
    r = bullets(ws, r, COLS, [
        "Success = BOTH N.ChMeas.MIMO.DL.MuPairing (all-layer and ≥4L) AND User DL Average Throughput. "
        "One up and the other down is a fail of that night — restore that night’s switch.",
        "P0 tonight is SRS_IC_SW-0 only. Do not start P1 (weights / tracking) on the same night — "
        "weights on a bad SRS estimate would bake the pairing damage into DL BF.",
        "P1.4 BEAM_TRACKING_SW needs P1.1 SUPER_COVERAGE_SW already committed. That is the 10-Sep failure.",
        "One switch per MML line. HighLayerMuMimoSw, BeamOptAlgoSwitch and DLMuMimoSchSupplementSw are bit-fields "
        "— list every bit that must remain ON.",
        "Never downgrade MaxMimoLayerNum=LAYER_16. Never roll HighPrecisionBeamSwitch while children are 1.",
        "Neighbour 32T control, same busy hour, PRB load as a normaliser. IBLER, drop and HOSR are hard gates.",
    ], fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 128
    return ws


def add_cross_links(wb):
    ws = wb[v7.S14]
    r = ws.max_row + 2
    r = blank(ws, r, 10)
    r = section(ws, r, bcs.COLS, "Trial result of this proposal  —  Ph1 executed 10–11 Sep 2026")
    r = note_bar(ws, r, bcs.COLS,
                 "DL MU pairing fell and DL user throughput stayed flat. Sheet 20 is the current action plan "
                 "for both KPIs. P0 tonight = sheet 18 RB-1. Sheet 19 is the pairing-only variant. "
                 "Sheet 17 is the older tput-first plan and is superseded.",
                 fill_hex="FFF2CC")
    for label, target in (
        (f"→  {D_SHEET}   (START HERE — pairing + DL tput action plan)", D_SHEET),
        (f"→  {v8.R_SHEET}   (P0 rollback MML)", v8.R_SHEET),
        (f"→  {v9.P_SHEET}   (pairing-only variant, if tput is parked)", v9.P_SHEET),
        (f"→  {v7.F_SHEET}   (what happened in Ph1)", v7.F_SHEET),
        (f"→  {v7.A_SHEET}   (older Ph1–Ph7 plan — superseded)", v7.A_SHEET),
    ):
        merge(ws, r, 1, r, bcs.COLS)
        href_sheet(ws.cell(r, 1), target, label)
        for c in range(1, bcs.COLS + 1):
            ws.cell(r, c).fill = fill(PALE_GOLD)
            ws.cell(r, c).border = thin
        ws.row_dimensions[r].height = 22
        r += 1


def extra(wb):
    v7.build_findings(wb)
    v7.build_action(wb)
    v8.build_rollback(wb)
    v9.build_pairing_plan(wb)
    build_dual_plan(wb)
    add_cross_links(wb)


def main():
    bcs.main(out=OUT, zip_out=ZIP_OUT, with_prereq=True, extra=extra, version="v10.0")


if __name__ == "__main__":
    main()
