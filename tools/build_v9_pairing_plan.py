"""
MIMO_Deployment v9.0 — Pairing Lift Plan.

North-star KPI is N.ChMeas.MIMO.DL.MuPairing (kLayer.RB / Pair.Layer.Avg / Pair.PRB),
not DL user throughput. The previous Ph1–Ph7 plan mixed DL-tput levers (beam tracking,
SRS weights) with pairing. This sheet is pairing only: recover the Ph1 loss, then turn
on the multilayer bits that actually spend the live LAYER_16 quota.

Run:  python3 build_v9_pairing_plan.py
"""
import os

from mimo_excel_style import *
import build_combined_sheet as bcs
from build_combined_sheet import href_sheet
import build_v7_findings as v7
import build_v8_rollback as v8

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "MIMO_Deployment_v9.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v9.0.zip")

P_SHEET = bcs.PAIRING_PLAN_SHEET
PID = "{NrDuCellId}"
COLS = v7.COLS
W = [8, 16, 28, 32, 38, 22, 22, 24, 22, 16, 18]
RED_HDR = v7.RED_HDR

# HighLayerMuMimoSw is a bit-field — always list every bit that must remain ON.
ML = "MMIMO_MULTILAYER_ENHANCE_SW-1"
PREF = "MU_MIMO_PAIRING_PREFERRED_SW-1"
RANK = "MU_RANK_BOOSTING_SW-1"
SRSACC = "SRS_MEAS_ACCELERATING_SW-1"

PHASES = [
    ("P0", "NEXT\n(tonight)",
     "Recover pairing. Undo the Ph1 bits that reduced N.ChMeas.MIMO.DL.MuPairing. Add nothing new.",
     "RB-1: SRS_IC_SW-0 alone. If 24 h later the recover-gate is missed: RB-2 PRECISE-0 (keep ANTI_INTRF), "
     "then RB-3 ANTI_INTRF-0. Full MML on sheet 18.",
     "LST NRDUCELLSRS (tight MUX still 1, IC still 1 before the change). Neighbour 32T control identified.",
     "All-layer MuPairing.RB ≥ 14.7 (90 % of 9 Sep ≈ 16.4) AND ≥4L ≥ 8.8. SRS NI back to 5–9 Sep. "
     "DL IBLER / drop unchanged vs control.",
     "Restore that night’s bit only (sheet 18 restore block). Do not touch HighPrecisionBeamSwitch.",
     "NEXT"),
    ("P1", "P0 gate met\n+ 24 h",
     "Spend the LAYER_16 quota. This is the first switch that can RAISE high-layer pairing rather than only recover it.",
     f"LST LICENSE (NR0S0DLEPU00). Then HighLayerMuMimoSw={ML} as the ONLY HighLayer bit this night.",
     "P0 gate already met. MuMimoSwitch=DL_MU_MIMO_SW-1. MaxMimoLayerNum already LAYER_16 — never downgrade it. "
     "FR1 32T only, not 2T2R indoor.",
     "N.ChMeas.MIMO.DL.Transmission.Layer.Max rises. 8–16 layer MuPairing.RB leaves zero. "
     "≥4L paired RB ≥ 9.8 (the 9 Sep baseline). IBLER inside band.",
     f"MOD NRDUCELLDLMIMO: HighLayerMuMimoSw={ML.replace('-1', '-0')};",
     "PLANNED"),
    ("P2", "P1 gate met\nthen 3 nights",
     "Prefer MU pairing, then raise MU rank. These children only work while the P1 master stays 1.",
     "Night 1: add PAIRING_PREFERRED (largest pairing lever of the three). Night 2: RANK_BOOSTING. "
     "Night 3: SRS_MEAS_ACCELERATING. Always list the master bit as -1 in the same MOD.",
     f"LST NRDUCELLDLMIMO must already show {ML}. P1 gate green.",
     "All-layer MuPairing.RB > 16.4 (beat 9 Sep). ≥4L > 9.8. Pair.Layer.Avg up. IBLER inside band.",
     "Children back to 0 first (newest first), master last.",
     "PLANNED"),
    ("P3", "P2 gate met\n+ 24 h",
     "Stop healthy MU pairs falling back to SU. FPD multilayer sample uses DlMuBackToSuSeThld=0; live is 5.",
     "MOD NRDUCELLDLMIMO DlMuBackToSuSeThld=0. Parameter, not a switch — one line, one night.",
     "P2 gate green. Confirm live value is 5 with LST before changing it.",
     "Pair.Layer.Avg holds higher through the busy hour (fewer SU fallbacks). IBLER still inside band.",
     "MOD NRDUCELLDLMIMO: DlMuBackToSuSeThld=5;   // back to the live value",
     "PLANNED"),
    ("P4", "P3 green\nloaded BH only",
     "Loaded-hour pairing resolution. AHR Capacity Upgrade 2.0 is the AHR wave that upgrades MU pairing, not Turbo.",
     "Night 1: AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1. Then one child per night: MuMimoOptSwith=ON, "
     "DlMuGatherOptSw=ON, PDCCH_MULTI_DIM_JOINT_SCH_SW-1. Do not send DlSrsMuMimoPreSinrThld=0 unless IBLER has room.",
     "AHR_PHASE1 and AHR_EXP_TURBO_PHASE2 already ON. Licence NR0S00ACT200. P0–P3 green. Loaded 32T only.",
     "All-layer and ≥4L pairing above the post-P2 level in the same BH. CCE not blocking. IBLER inside band.",
     "Children, then AHR_CAPC_UPGRADE_PHASE2_SW-0. Turbo and Phase1 stay ON.",
     "PLANNED"),
    ("P5", "P4 + 7 days",
     "Cluster rollout of the proven pairing set only.",
     "Same MML sequence per remaining 32T cell. Fix Tilt=255 before a cell joins. 2T2R indoor stays out. "
     "Do not put SRS_IC, PRECISE_SCH or ANTI_INTRF back until pairing has been above target for 7 days.",
     "Per-cell LST of HighLayerMuMimoSw, SrsDetectionAlgoSwitch, DLMuMimoSchSupplementSw, MaxMimoLayerNum.",
     "Each gNB ≥ its own pre-week MuPairing and ≥ the trial-cluster post-P2 level. No IBLER/drop regression.",
     "Per-gNB reverse of that gNB’s CME file only.",
     "PLANNED"),
]

# Sequenced MML for P0–P4. P0 points at sheet 18; P1–P4 are the lift.
SEQ = [
    ("P0.1", "Tonight\nalone",
     "SrsDetectionAlgoSwitch=SRS_IC_SW-0",
     f"LST NRDUCELLSRS: NrDuCellId={PID};\n"
     f"MOD NRDUCELLSRS: NrDuCellId={PID}, SrsDetectionAlgoSwitch=SRS_IC_SW-0;\n"
     f"LST NRDUCELLSRS: NrDuCellId={PID};",
     "None — rollback. Keep SRS_TIGHT_MULTIPLEXING_SW-1.",
     "Sheet 18 RB-1. Tight MUX and SRS_SINR_MEAS_OPT still 1. RETCODE = 0.",
     "MuPairing.kLayer.RB, N.SRS.NI.Avg",
     "Recover-gate after 24 h: all-layer ≥ 14.7 and ≥4L ≥ 8.8. If YES → P1. If NO → P0.2."),
    ("P0.2", "Only if P0.1\nmissed",
     "DL_MU_PRECISE_SCH_SW-0, keep ANTI_INTRF",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
     "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-1;",
     "Bit-field: list ANTI_INTRF as -1. Never the same night as P0.3.",
     "Sheet 18 RB-2. PRECISE gone, ANTI_INTRF still 1.",
     "MuPairing.kLayer.RB, DL IBLER",
     "Same recover-gate. IBLER jump → restore PRECISE."),
    ("P0.3", "Only if P0.2\nmissed",
     "DL_MU_ANTI_INTRF_SCH_SW-0",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
     "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-0;",
     "List both bits as -0 so PRECISE is not restored by omission.",
     "Sheet 18 RB-3.",
     "MuPairing.kLayer.RB, DL IBLER",
     "Same recover-gate. Then go to P1 even if short — remaining lift is multilayer, not more rollback."),
    ("P1.1", "P0 green\n1 night",
     f"HighLayerMuMimoSw={ML}",
     f"LST LICENSE;\n"
     f"LST NRDUCELLPDSCH: NrDuCellId={PID};   // MaxMimoLayerNum must be LAYER_16\n"
     f"LST NRDUCELLALGOSWITCH: NrDuCellId={PID};   // DL_MU_MIMO_SW must be 1\n"
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML};\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "DL_MU_MIMO_SW-1 + LAYER_16 + NR0S0DLEPU00. Master of every later HighLayer bit.",
     "HighLayerMuMimoSw shows MMIMO_MULTILAYER_ENHANCE_SW-1. RETCODE = 0.",
     "Transmission.Layer.Max, MuPairing.8–16Layer.RB, ≥4L RB, DL IBLER",
     "Layer.Max up and 8–16 layer RB leaves 0. If Layer.Max stays flat, do not send P2 — check licence."),
    ("P2.1", "P1 green\nnight 1",
     f"{ML}&{PREF}",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML}&{PREF};\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "THE pairing lever among the children: prefer MU when a pair has gain. Master must stay listed as -1.",
     "Both MMIMO_MULTILAYER_ENHANCE and MU_MIMO_PAIRING_PREFERRED selected.",
     "MuPairing.kLayer.RB, Pair.PRB, Pair.Layer.Avg, DL IBLER",
     "Largest expected step-up in paired RB. IBLER is the hard gate."),
    ("P2.2", "P2.1 green\nnight 2",
     f"{ML}&{PREF}&{RANK}",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML}&{PREF}&{RANK};\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "Rank boosting raises layers inside an already-paired group. List every bit that must stay ON.",
     "All three HighLayer bits selected.",
     "MuPairing.4–8Layer.RB, N.PDSCH.InitTbDl.Rank*, DL IBLER",
     "4–8 layer RB should rise. Rank mix shifts up."),
    ("P2.3", "P2.2 green\nnight 3",
     f"{ML}&{PREF}&{RANK}&{SRSACC}",
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, HighLayerMuMimoSw={ML}&{PREF}&{RANK}&{SRSACC};\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "Faster SRS so pairing decisions stay fresh. Watch SRS NI — this is not itself a pairing gate.",
     "Four HighLayer bits selected.",
     "MuPairing.kLayer.RB, N.UL.SRS.PreSINR, N.SRS.NI.Avg",
     "NI must not jump. Pairing holds or rises."),
    ("P3.1", "P2 green\n1 night",
     "DlMuBackToSuSeThld=0  (live is 5)",
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};   // confirm 5\n"
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, DlMuBackToSuSeThld=0;\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "FPD multilayer sample. Lowers the SE test that kicks a MU pair back to SU.",
     "DlMuBackToSuSeThld = 0.",
     "Pair.Layer.Avg through BH, DL IBLER",
     "Pairs last longer in BH. IBLER jump → set back to 5."),
    ("P4.1", "Loaded BH\nnight 1",
     "AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1",
     f"LST LICENSE;   // NR0S00ACT200\n"
     f"LST NRDUCELLFEATURESW: NrDuCellId={PID};   // Phase1+Turbo already 1\n"
     f"MOD NRDUCELLFEATURESW: NrDuCellId={PID}, AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1;",
     "AHR Phase1 + Turbo already ON. This is the AHR wave that upgrades pairing resolution under load.",
     "AhrSwitch shows AHR_CAPC_UPGRADE_PHASE2.",
     "MuPairing in the loaded BH, CQI, DL IBLER, CCE",
     "Pairing in the loaded hour above post-P2. Then children one per night (sheet 14 AHR CU box)."),
    ("P4.2", "P4.1 green\nnight 2",
     "MuMimoOptSwith=ON  (FPD spelling)",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoOptSwith=ON;",
     "AHR CU child. MU-MIMO optimisation.",
     "MuMimoOptSwith = ON.",
     "MuPairing.kLayer.RB, DL IBLER",
     "Loaded-hour pairing holds or rises."),
    ("P4.3", "P4.2 green\nnight 3",
     "DlMuGatherOptSw=ON",
     f"MOD NRDUCELLPDSCH: NrDuCellId={PID}, DlMuGatherOptSw=ON;",
     "Gathering MU — more UEs considered for the same pair group.",
     "DlMuGatherOptSw = ON.",
     "MuPairing.kLayer.RB, Pair.PRB, DL IBLER",
     "Paired PRB up. IBLER gate."),
    ("P4.4", "P4.3 green\nnight 4",
     "PDCCH_MULTI_DIM_JOINT_SCH_SW-1",
     f"MOD NRDUCELLALGOSWITCH: NrDuCellId={PID}, MuMimoSwitch=PDCCH_MULTI_DIM_JOINT_SCH_SW-1;",
     "Joint PDCCH/PDSCH sch so pairing is not CCE-blocked. Keep existing DL_MU / UL_MU / PDCCH_MU bits ON.",
     "PDCCH_MULTI_DIM_JOINT_SCH selected and DL_MU_MIMO still 1.",
     "MuPairing, N.CCE.Used.Avg, N.CCE.DL.AllocReq.Num",
     "CCE used must not rise into blocking. Pairing holds."),
]

PARK = [
    ("SRS_WEIGHT_ESTIMATE_SW / BEAM_TRACKING_SW / INTELLIGENT_BEAM_SELECTION_SW / BEAM_SELECT_OPT_SW",
     "These are DL MCS / DL user-throughput levers. They never landed in Ph1, and they are not what moves "
     "N.ChMeas.MIMO.DL.MuPairing. Re-run from old sheet 17 only AFTER P2 pairing target is met.",
     "Park until pairing > 9 Sep"),
    ("SRS_IC_SW-1",
     "This is the switch that most likely caused the Ph1 pairing collapse next to tight multiplexing. "
     "Do not put it back to chase SRS capacity.",
     "Do not re-enable"),
    ("DL_MU_PRECISE_SCH_SW / DL_MU_ANTI_INTRF_SCH_SW",
     "Quality gates. They reduce paired RB by design. Keep them OFF until pairing has been above the "
     "9 Sep baseline for 7 days, then trial one of them — not both — if IBLER needs help.",
     "Off until pairing is green + 7 d"),
    ("SRS_BLIND_IS_SW (multilayer child) / iBeam 2.0 PRECISE_MUMIMO_EVAL",
     "More interference-source / pairing-quality gates. Same direction as PRECISE_SCH — they can cut pairing. "
     "Not in this plan.",
     "Hold"),
    ("SSB_BEAM_ADAPT / UL Boosting 2.0 / iBeam 2.0–3.0",
     "Different KPIs (SSB users, UL tput, DL MCS). They do not lift DL MuPairing and they cost windows.",
     "Hold"),
    ("DlSrsMuMimoSpaceIsoThld / DlMuMimoSrsPreSinrThld relaxation",
     "Live isolation 50 and PreSINR −50 already match the FPD MU sample. Loosening them raises pairing "
     "and IBLER together. Only as a last resort after P4, 5-cell trial, IBLER as hard gate.",
     "Last resort only"),
]

LEVERS = [
    ("RAISE", "MMIMO_MULTILAYER_ENHANCE_SW", "P1",
     "Master. Lets the cell spend LAYER_16. Without it high-layer MuPairing cannot rise."),
    ("RAISE", "MU_MIMO_PAIRING_PREFERRED_SW", "P2.1",
     "Largest child lever. Prefers MU over SU when a pair has gain — directly MuPairing.PRB."),
    ("RAISE", "MU_RANK_BOOSTING_SW", "P2.2",
     "More layers inside an already-paired group. Moves 4–8 layer RB."),
    ("RAISE", "DlMuBackToSuSeThld 5 → 0", "P3",
     "Stops healthy pairs falling back to SU on the SE test."),
    ("RAISE", "AHR_CAPC_UPGRADE_PHASE2 + MuMimoOpt / Gather / joint sch", "P4",
     "Loaded-hour pairing resolution. Turbo does not do this — CU does."),
    ("RECOVER", "SRS_IC_SW-0", "P0.1",
     "Removes the SRS-detection conflict with tight multiplexing. Restores the SRS that pairing needs."),
    ("RECOVER", "DL_MU_PRECISE_SCH_SW-0 then ANTI_INTRF-0", "P0.2 / P0.3",
     "Removes the two quality gates Ph1 added. Only if P0.1 missed the recover-gate."),
    ("REDUCE", "SRS_IC_SW-1 + tight MUX together", "Ph1 (done)",
     "What dropped pairing. Do not repeat."),
    ("REDUCE", "DL_MU_PRECISE_SCH / ANTI_INTRF", "Ph1 (done)",
     "Quality gates. Fewer pairs by design. Stay off until the lift is proven."),
]


def build_pairing_plan(wb):
    ws = wb.create_sheet(P_SHEET, 19)
    setup_sheet(ws, P_SHEET)
    set_widths(ws, W)
    ws.oddHeader.left.text = "5G MIMO — Pairing Lift Plan  (north-star N.ChMeas.MIMO.DL.MuPairing)"
    ws.oddFooter.left.text = FOOTER_LEFT
    ws.sheet_properties.tabColor = "C00000"

    r = banner(ws, 1, COLS,
               "  Pairing Lift Plan  —  north-star KPI  N.ChMeas.MIMO.DL.MuPairing",
               fill_hex=RED)
    r = note_bar(ws, r, COLS,
                 "Target is to RAISE DL MU pairing, not DL user throughput. Recover the Ph1 loss first (P0), "
                 "then enable the multilayer bits that spend the live LAYER_16 quota (P1–P2). Beam tracking / "
                 "SRS weights stay parked — they move MCS, not paired RB. Replace {NrDuCellId} with 101 / 102 / 103. "
                 "One switch per night. HighLayerMuMimoSw is a bit-field: list every bit that must remain ON.",
                 fill_hex="F8CBAD")
    r = blank(ws, r, 6)

    # ----- Section 1: target numbers
    r = section(ws, r, COLS, "Section 1.  Numeric target  —  N.ChMeas.MIMO.DL.MuPairing.kLayer.RB")
    r = v7._hdr(ws, r, ["#", "Counter", "5 Sep", "9 Sep\n(baseline)", "11 Sep\n(Ph1)", "12 Sep",
                        "P0 recover gate\n(90 % of 9 Sep)", "P1–P2 success\n(beat 9 Sep)",
                        "How we get there", "Safety gate", "Safety (cont.)"],
                fill_hex=TEAL, height=36)
    targets = [
        (1, "All-layer paired RB  Σ k=1…16", "16.4", "16.4", "6.4 (−61 %)", "9.6",
         "≥ 14.7", "> 16.4", "P0 recover, then P1–P2 lift",
         "PRB load comparable; DL IBLER inside band vs control"),
        (2, "High-layer paired RB  Σ k≥4", "9.1", "9.8", "3.3 (−66 %)", "4.8",
         "≥ 8.8", "> 9.8", "P1 master + P2.1 PAIRING_PREFERRED",
         "Transmission.Layer.Max must rise with it"),
        (3, "8–16 layer paired RB", "~0", "~0", "0", "0",
         "still ~0 (ok)", "first non-zero", "P1 MMIMO_MULTILAYER_ENHANCE_SW",
         "If still 0 after P1, licence NR0S0DLEPU00 is the blocker"),
        (4, "Pair.Layer.Avg / Pair.PRB", "—", "baseline", "down", "partial",
         "back toward baseline", "above baseline", "Sister of kLayer.RB — must agree",
         "Must not move against kLayer.RB"),
    ]
    for rec in targets:
        i = rec[0]
        base = alt_fill(i)
        fills = [base, base, PALE_GREEN, PALE_GREEN, "F8CBAD", "FCE4D6",
                 "FFF2CC", PALE_GREEN, "D6EAF8", "FFF2CC", "FFF2CC"]
        vals = list(rec) + [""] * (COLS - len(rec))
        r = v7._row(ws, r, vals[:COLS], fills, bolds={2, 7, 8},
                    center={1, 3, 4, 5, 6, 7, 8},
                    height=v7._height([(W[1], rec[1]), (W[8], rec[8]), (W[9], rec[9])], cap=70))
    r = note_bar(ws, r, COLS,
                 "Pairing figures are approximate (read off the Ph1 pivot). Confirm from MAE before signing "
                 "a gate. Baseline = 5 Sep and 9 Sep. 10 Sep is execution day 1, not a baseline day.")
    r = blank(ws, r, 8)

    # ----- Section 2: levers
    r = section(ws, r, COLS, "Section 2.  What raises pairing vs what reduces it")
    r = v7._hdr(ws, r, ["#", "Direction", "Switch / parameter", "Switch (cont.)", "Phase",
                        "Effect on N.ChMeas.MIMO.DL.MuPairing", "Effect (cont.)", "Effect (cont.)",
                        "Effect (cont.)", "Send?", "Send?"], fill_hex=NAVY, height=28)
    dir_fill = {"RAISE": "D5F5E3", "RECOVER": "D6EAF8", "REDUCE": "F8CBAD"}
    for i, (d, sw, ph, fx) in enumerate(LEVERS, 1):
        df = dir_fill[d]
        put(ws, r, 1, i, size=9, bold=True, fill_hex=df, h="center", v="center", border=True)
        put(ws, r, 2, d, size=9, bold=True, fill_hex=df, h="center", v="center", border=True)
        merge(ws, r, 3, r, 4)
        put(ws, r, 3, sw, size=8, bold=True, fill_hex=df, h="left", v="top", border=True)
        ws.cell(r, 4).border = thin
        put(ws, r, 5, ph, size=8, bold=True, fill_hex=df, h="center", v="center", border=True)
        merge(ws, r, 6, r, 9)
        put(ws, r, 6, fx, size=8, fill_hex=WHITE, h="left", v="top", border=True)
        for c in (7, 8, 9):
            ws.cell(r, c).border = thin
        merge(ws, r, 10, r, 11)
        send = "YES, in sequence" if d != "REDUCE" else "NO — keep off"
        put(ws, r, 10, send, size=8, bold=True,
            fill_hex=PALE_GREEN if d != "REDUCE" else "F8CBAD", h="center", v="center", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(v7._w(W, 3, 4), sw), (v7._w(W, 6, 9), fx)], cap=72)
        r += 1
    r = blank(ws, r, 8)

    # ----- Section 3: phase overview
    r = section(ws, r, COLS, "Section 3.  Phase overview  (pairing-only — replaces the old Ph1–Ph7 order)")
    r = v7._hdr(ws, r, ["Phase", "Window", "Objective", "Content  (one switch per night)",
                        "Pre-requisite BEFORE the work order", "Pre-req (cont.)",
                        "Exit gate  (N.ChMeas.MIMO.DL.MuPairing)", "Exit (cont.)",
                        "Rollback", "Status", "Status"], fill_hex=GREEN, height=40)
    st_fill = {"NEXT": "FFC000", "PLANNED": "EAECEE"}
    for ph, win, obj, content, pre, gate, back, status in PHASES:
        base = st_fill[status]
        put(ws, r, 1, ph, size=12, bold=True, color=WHITE,
            fill_hex=RED_HDR if status == "NEXT" else GREEN, h="center", v="center", border=True)
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
             (v7._w(W, 7, 8), gate), (W[8], back)], cap=160)
        r += 1
    r = blank(ws, r, 8)

    # ----- Section 4: sequenced MML
    r = section(ws, r, COLS, "Section 4.  Sequenced MML  —  copy one block per night", fill_hex=RED_HDR)
    r = note_bar(ws, r, COLS,
                 "HighLayerMuMimoSw is a bit-field, same class of mistake as the 10-Sep BEAM_TRACKING rejection. "
                 "When adding PAIRING_PREFERRED you must list MMIMO_MULTILAYER_ENHANCE_SW-1 in the same MOD or the "
                 "master drops. P0.1 is the only change tonight.",
                 fill_hex="F8CBAD")
    r = v7._hdr(ws, r, ["Seq", "When", "Switch / parameter", "MML  (one switch, LST → MOD → LST)",
                        "MML (cont.)", "Pre-requisite / bit-field rule", "Verify",
                        "Counter  (MuPairing first)", "Counter (cont.)",
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
             (v7._w(W, 8, 9), ctr), (v7._w(W, 10, 11), gate)], cap=200)
        r += 1
    r = blank(ws, r, 8)

    # ----- Section 5: park
    r = section(ws, r, COLS, "Section 5.  Parked  —  do not send these until pairing is above 9 Sep",
                fill_hex=GRAY)
    r = v7._hdr(ws, r, ["#", "Parked item", "Parked item (cont.)", "Parked item (cont.)",
                        "Why it is not in this plan", "Why (cont.)", "Why (cont.)", "Why (cont.)",
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
            [(v7._w(W, 2, 4), item), (v7._w(W, 5, 8), why)], cap=80)
        r += 1
    r = blank(ws, r, 8)

    r = subsection(ws, r, COLS, "Standing rules for this KPI", fill_hex=NAVY2)
    r = bullets(ws, r, COLS, [
        "North-star is N.ChMeas.MIMO.DL.MuPairing.kLayer.RB (all-layer and ≥4L). DL user throughput is a "
        "safety KPI, not the success KPI.",
        "One switch per night. HighLayerMuMimoSw and DLMuMimoSchSupplementSw are bit-fields — list every "
        "bit that must remain ON.",
        "P0 tonight is SRS_IC_SW-0 only. Do not start P1 on the same night.",
        "Never downgrade MaxMimoLayerNum=LAYER_16. Never roll HighPrecisionBeamSwitch while children are 1.",
        "Neighbour 32T control, same busy hour, PRB load as a normaliser. A pairing rise on a load jump is not a pass.",
        "IBLER, drop and HOSR are hard gates on every night. Pairing bought with IBLER is a fail.",
    ], fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 110
    return ws


def add_cross_links(wb):
    ws = wb[v7.S14]
    r = ws.max_row + 2
    r = blank(ws, r, 10)
    r = section(ws, r, bcs.COLS, "Trial result of this proposal  —  Ph1 executed 10–11 Sep 2026")
    r = note_bar(ws, r, bcs.COLS,
                 "DL MU pairing fell. The current north-star is N.ChMeas.MIMO.DL.MuPairing. "
                 "Sheet 19 is the pairing lift plan. Sheet 18 is the P0 rollback MML. "
                 "Sheet 17 is the older DL-throughput plan and is parked until pairing is above 9 Sep.",
                 fill_hex="F8CBAD")
    for label, target in (
        (f"→  {P_SHEET}   (pairing lift plan — START HERE)", P_SHEET),
        (f"→  {v8.R_SHEET}   (P0 rollback MML)", v8.R_SHEET),
        (f"→  {v7.F_SHEET}   (what happened in Ph1)", v7.F_SHEET),
        (f"→  {v7.A_SHEET}   (older DL-tput plan — parked)", v7.A_SHEET),
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
    build_pairing_plan(wb)
    add_cross_links(wb)


def main():
    bcs.main(out=OUT, zip_out=ZIP_OUT, with_prereq=True, extra=extra, version="v9.0")


if __name__ == "__main__":
    main()
