"""
MIMO_Deployment v8.0 — pairing rollback pack on top of v7.0.

The Ph1 trial dropped DL MU pairing (all-layer −61 % on 11 Sep, high-layer worse).
This sheet is the executable rollback: one switch per night, in the order most likely
to restore pairing, with a hard STOP if the gate is met.

Run:  python3 build_v8_rollback.py
"""
import os

from mimo_excel_style import *
import build_combined_sheet as bcs
from build_combined_sheet import href_sheet
import build_v7_findings as v7

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "MIMO_Deployment_v8.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v8.0.zip")

R_SHEET = bcs.ROLLBACK_SHEET
PID = "{NrDuCellId}"
COLS = v7.COLS
W = [8, 14, 28, 32, 36, 22, 22, 24, 22, 16, 18]
RED_HDR = v7.RED_HDR

# Why this order: SRS input first, then the two pairing gates one at a time.
# DLMuMimoSchSupplementSw is a bit-field — always list the bits that must remain ON.
STEPS = [
    ("RB-1", "P0  ·  tonight  ·  ALONE",
     "NRDUCellSrs / SrsDetectionAlgoSwitch",
     "SRS_IC_SW-0",
     f"LST NRDUCELLSRS: NrDuCellId={PID};\n"
     f"MOD NRDUCELLSRS: NrDuCellId={PID}, SrsDetectionAlgoSwitch=SRS_IC_SW-0;\n"
     f"LST NRDUCELLSRS: NrDuCellId={PID};",
     "None — this is a rollback. It removes the switch that conflicts with "
     "SRS_TIGHT_MULTIPLEXING_SW-1. Keep tight multiplexing ON.",
     "LST: SrsDetectionAlgoSwitch no longer shows SRS_IC_SW-1. "
     "SrsAlgoSwitch still shows SRS_TIGHT_MULTIPLEXING_SW-1 and SRS_SINR_MEAS_OPT_SW-1. "
     "RETCODE = 0.",
     "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB, N.SRS.NI.Avg, N.UL.SRS.PreSINR.Index*",
     "After 24 h busy hour: all-layer paired RB ≥ 14.7 (90 % of 9 Sep ≈ 16.3) AND "
     "high-layer (≥4L) ≥ 8.7. If YES → STOP, do not send RB-2. If NO → RB-2 next night."),
    ("RB-2", "only if RB-1 missed the gate",
     "NRDUCellDlMimo / DLMuMimoSchSupplementSw",
     "DL_MU_PRECISE_SCH_SW-0  (keep ANTI_INTRF)",
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};\n"
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
     "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-1;\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "Bit-field: list ANTI_INTRF as -1 so it stays selected while PRECISE comes off. "
     "Never send PRECISE-0 and ANTI_INTRF-0 in the same line.",
     "LST: DLMuMimoSchSupplementSw shows DL_MU_ANTI_INTRF_SCH_SW-1 and does NOT show "
     "DL_MU_PRECISE_SCH_SW-1. RETCODE = 0.",
     "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB, DL IBLER, N.PRB.DL.Used.Avg",
     "Same pairing gate after 24 h. Watch IBLER — if it jumps vs control, restore "
     "DL_MU_PRECISE_SCH_SW-1 that night. If pairing still short and IBLER holds → RB-3."),
    ("RB-3", "only if RB-2 missed the gate",
     "NRDUCellDlMimo / DLMuMimoSchSupplementSw",
     "DL_MU_ANTI_INTRF_SCH_SW-0",
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};\n"
     f"MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
     "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-0;\n"
     f"LST NRDUCELLDLMIMO: NrDuCellId={PID};",
     "Both MU-sch bits now 0. PRECISE is already 0 from RB-2 — list it as -0 so the "
     "bit-field is not restored by omission.",
     "LST: neither PRECISE nor ANTI_INTRF remains selected. RETCODE = 0.",
     "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB, DL IBLER",
     "Same pairing gate after 24 h. If IBLER jumped, restore ANTI_INTRF only. "
     "If pairing still short and IBLER holds → RB-4."),
    ("RB-4", "last-resort pairing gate",
     "NRDUCellSrsMeas / SrsMeasOptSwitch",
     "SRS_BLIND_IS_MEAS_SW-0",
     f"LST NRDUCELLSRSMEAS: NrDuCellId={PID};\n"
     f"MOD NRDUCELLSRSMEAS: NrDuCellId={PID}, SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-0;\n"
     f"LST NRDUCELLSRSMEAS: NrDuCellId={PID};",
     "iBeam 1.0 child. HighPrecisionBeamSwitch stays ON. Do not roll the master.",
     "LST: SrsMeasOptSwitch no longer shows SRS_BLIND_IS_MEAS_SW-1. RETCODE = 0.",
     "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB, N.UL.SRS.PreSINR.Index*, DL IBLER",
     "Same pairing gate. If still short after RB-4, the remaining cause is load, RF "
     "(Tilt=255) or something not in the Ph1 list — stop rolling switches."),
    ("RB-5", "only if SRS NI still high after RB-1",
     "NRDUCellUlPcConfig / UlPwrCtrlAlgoSwitch",
     "SRS_JOINT_PC_SW-0",
     f"LST NRDUCELLULPCCONFIG: NrDuCellId={PID};\n"
     f"MOD NRDUCELLULPCCONFIG: NrDuCellId={PID}, UlPwrCtrlAlgoSwitch=SRS_JOINT_PC_SW-0;\n"
     f"LST NRDUCELLULPCCONFIG: NrDuCellId={PID};",
     "Skip this if SRS NI recovered after RB-1. Joint PC was tuned to sit next to SRS_IC; "
     "it is not a pairing gate.",
     "LST: UlPwrCtrlAlgoSwitch no longer shows SRS_JOINT_PC_SW-1. RETCODE = 0.",
     "N.SRS.NI.Avg, N.UL.SRS.PreSINR.Index*, N.ThpVol.UL, UL IBLER",
     "SRS NI back to 5–9 Sep. UL tput / UL IBLER must not regress. Independent of RB-2/3/4."),
]

KEEP = [
    ("SRS_TIGHT_MULTIPLEXING_SW-1", "NRDUCellSrs / SrsAlgoSwitch",
     "This is the SRS bit CR01 chose. The conflict is SRS_IC, not tight multiplexing. "
     "Rolling this back would shrink SRS capacity on a loaded 32T cell."),
    ("SRS_SINR_MEAS_OPT_SW-1", "NRDUCellSrs / SrsAlgoSwitch",
     "Measurement-accuracy bit. Not a pairing gate. Confirm it is still 1 after RB-1."),
    ("HighPrecisionBeamSwitch=ON", "NRDUCellFeatureSw",
     "iBeam 1.0 master. Rolling it would silently disable every iBeam child still at 1. "
     "Roll children first, master last — and we are not rolling the master."),
    ("UL_LOW_NOISE_SW-1 + UL_MU_GRP_PAIR / DIFF_WAVEFORM / UL_CORR_ACCELERATION",
     "NRDUCellFeatureSw / NRDUCellUlMimo",
     "UL Boosting. The reported regression is DL pairing, not UL. Leave the UL pack alone."),
    ("TAIL_PKT_MCS_OPT / RES_BASED_DL_ADAPT_SCH / DL_RLC_STAT_RPT_MERGE / PDCCH_AGG_LVL_COMPR",
     "NRDUCellDlSch / NRDUCellPdcch",
     "iBeam 1.0 schedulers. Neutral on pairing. Keep ON."),
    ("UL_RANK_FAST_DECREASE / PUSCH_CE_SINR_LEVEL_ENH",
     "NRDUCellUlRank / NRDUCellPusch",
     "UL rank / CE. Not on the DL pairing path."),
    ("MaxMimoLayerNum=LAYER_16, SU/MU masters, AHR Phase1+Turbo",
     "NRDUCellPdsch / NRDUCellAlgoSwitch / NRDUCellFeatureSw",
     "Already commercial. Never downgrade these to ‘fix’ pairing."),
    ("The four rejected beam bits (SRS_WEIGHT_ESTIMATE / BEAM_SELECT_OPT / "
     "BEAM_TRACKING / INTELLIGENT_BEAM_SELECTION)",
     "NRDUCellBeamAlgo — never landed",
     "They were never ON, so there is nothing to roll back. Re-run them later from "
     "sheet 17 Ph2 night 2, after pairing has recovered."),
]

RANK = [
    (1, "SRS_IC_SW-1", "SrsDetectionAlgoSwitch",
     "Enabled in the same window as SRS_TIGHT_MULTIPLEXING_SW-1. Two SRS-detection changes "
     "at once degrade the channel estimate that MU pairing uses. Highest probability, "
     "lowest collateral.",
     "RB-1  P0"),
    (2, "DL_MU_PRECISE_SCH_SW-1", "DLMuMimoSchSupplementSw",
     "Quality gate: pairs only UEs whose channel estimate is clean. Fed with degraded SRS "
     "it rejects most candidates. Explains why 4/6-layer RB fell hardest.",
     "RB-2  if RB-1 missed"),
    (3, "DL_MU_ANTI_INTRF_SCH_SW-1", "DLMuMimoSchSupplementSw",
     "Same mechanism as PRECISE — fewer but cleaner pairs. Second gate, so second in line, "
     "never the same night.",
     "RB-3  if RB-2 missed"),
    (4, "SRS_BLIND_IS_MEAS_SW-1", "SrsMeasOptSwitch",
     "Sharper interference-source measurement. Also a pairing gate, but a weaker one. "
     "Last-resort after the two MU-sch bits.",
     "RB-4  last resort"),
    (5, "SRS_JOINT_PC_SW-1", "UlPwrCtrlAlgoSwitch",
     "Not a pairing gate. Only if SRS NI is still high after RB-1. Independent path.",
     "RB-5  SRS-NI path"),
]


def build_rollback(wb):
    ws = wb.create_sheet(R_SHEET, 18)
    setup_sheet(ws, R_SHEET)
    set_widths(ws, W)
    ws.oddHeader.left.text = "5G MIMO — Pairing rollback pack (one switch per night)"
    ws.oddFooter.left.text = FOOTER_LEFT
    ws.sheet_properties.tabColor = "C00000"

    r = banner(ws, 1, COLS,
               "  Pairing rollback  —  recover DL MU pairing after the Ph1 drop,  one switch per night",
               fill_hex=RED)
    r = note_bar(ws, r, COLS,
                 "Replace {NrDuCellId} with 101 / 102 / 103 on every trial cell. Send RB-1 tonight as "
                 "the ONLY change. After 24 h of busy-hour counters: if pairing is back to ≥90 % of "
                 "the 9 Sep baseline, STOP. Do not bundle RB-1 with RB-2 — that is how Ph1 became "
                 "unattributable. Findings behind this order are on sheet 16 (F-02, F-03).",
                 fill_hex="F8CBAD")
    r = blank(ws, r, 6)

    r = section(ws, r, COLS, "Section 1.  Which Ph1 switches dropped pairing  —  rollback order")
    r = v7._hdr(ws, r, ["Rank", "Switch currently ON", "MO / Parameter ID",
                        "Why it dropped pairing", "(cont.)", "(cont.)",
                        "Rollback step", "When to send it", "When (cont.)",
                        "Do tonight?", "Do tonight?"], fill_hex=RED_HDR, height=30)
    for rank, sw, mo, why, step in RANK:
        base = "F8CBAD" if rank == 1 else ("FCE4D6" if rank <= 3 else alt_fill(rank))
        put(ws, r, 1, rank, size=12, bold=True, color=WHITE, fill_hex=RED_HDR,
            h="center", v="center", border=True)
        put(ws, r, 2, sw, size=8, bold=True, fill_hex=base, h="left", v="top", border=True)
        put(ws, r, 3, mo, size=8, fill_hex=base, h="left", v="top", border=True)
        merge(ws, r, 4, r, 6)
        put(ws, r, 4, why, size=8, fill_hex=WHITE, h="left", v="top", border=True)
        for c in (5, 6):
            ws.cell(r, c).border = thin
        merge(ws, r, 7, r, 9)
        put(ws, r, 7, step, size=9, bold=True, fill_hex=base, h="left", v="center", border=True)
        for c in (8, 9):
            ws.cell(r, c).border = thin
        merge(ws, r, 10, r, 11)
        tonight = "YES  —  only this" if rank == 1 else "NO  —  wait 24 h"
        put(ws, r, 10, tonight, size=9, bold=True,
            fill_hex="FFC000" if rank == 1 else "EAECEE", h="center", v="center", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(W[1], sw), (v7._w(W, 4, 6), why), (v7._w(W, 7, 9), step)], cap=90)
        r += 1

    r = blank(ws, r, 8)

    r = section(ws, r, COLS, "Section 2.  Sequenced rollback MML  —  copy one block per night",
                fill_hex=RED_HDR)
    r = note_bar(ws, r, COLS,
                 "DLMuMimoSchSupplementSw is a bit-field. RB-2 lists ANTI_INTRF as -1 so it stays "
                 "selected. RB-3 lists both bits as -0 so PRECISE is not restored by omission. "
                 "Confirm every line with LST before the next night.")
    r = v7._hdr(ws, r, ["Seq", "When", "MO / Parameter", "Target value",
                        "MML  (LST → MOD → LST, one switch)",
                        "MML (cont.)", "Pre-requisite / bit-field rule",
                        "Verify  (LST + RETCODE)", "Counter to watch",
                        "Stop / continue gate", "Stop / continue (cont.)"],
                fill_hex=RED_HDR, height=42)
    for i, (seq, when, mo, val, mml, pre, verify, ctr, gate) in enumerate(STEPS, 1):
        base = alt_fill(i)
        night_fill = "FFC000" if seq == "RB-1" else base
        put(ws, r, 1, seq, size=11, bold=True, color=WHITE, fill_hex=RED_HDR,
            h="center", v="center", border=True)
        put(ws, r, 2, when, size=8, bold=True, fill_hex=night_fill, h="center", v="top", border=True)
        put(ws, r, 3, mo, size=8, fill_hex=base, h="left", v="top", border=True)
        put(ws, r, 4, val, size=8, bold=True, fill_hex="FFF2CC", h="left", v="top", border=True)
        merge(ws, r, 5, r, 6)
        put(ws, r, 5, mml, size=8, bold=True, fill_hex="FCE4E4", h="left", v="top", border=True)
        ws.cell(r, 6).border = thin
        put(ws, r, 7, pre, size=8, fill_hex="FCE4E4", h="left", v="top", border=True)
        put(ws, r, 8, verify, size=8, fill_hex=PALE_BLUE, h="left", v="top", border=True)
        put(ws, r, 9, ctr, size=8, fill_hex="D5F5E3", h="left", v="top", border=True)
        merge(ws, r, 10, r, 11)
        put(ws, r, 10, gate, size=8, fill_hex=base, h="left", v="top", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(W[1], when), (W[3], val), (v7._w(W, 5, 6), mml), (W[6], pre),
             (W[7], verify), (W[8], ctr), (v7._w(W, 10, 11), gate)], cap=210)
        r += 1

    r = blank(ws, r, 6)
    copy_block = (
        "RB-1 tonight (repeat for every trial cell 101 / 102 / 103):\n"
        f"  LST NRDUCELLSRS: NrDuCellId={PID};\n"
        f"  MOD NRDUCELLSRS: NrDuCellId={PID}, SrsDetectionAlgoSwitch=SRS_IC_SW-0;\n"
        f"  LST NRDUCELLSRS: NrDuCellId={PID};\n"
        "Then wait 24 h of busy-hour counters. If the pairing gate is met, you are done.\n\n"
        "RB-2 only if the gate was missed:\n"
        f"  MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
        "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-1;\n\n"
        "RB-3 only if RB-2 was missed:\n"
        f"  MOD NRDUCELLDLMIMO: NrDuCellId={PID}, "
        "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-0&DL_MU_ANTI_INTRF_SCH_SW-0;"
    )
    r = body(ws, r, COLS, copy_block, fill_hex="FFF2CC", size=9,
             height=v7._height([(sum(W), copy_block)], size=9, cap=200))
    r = blank(ws, r, 8)

    r = section(ws, r, COLS, "Section 3.  Do NOT roll these back  —  they did not drop pairing",
                fill_hex=GREEN)
    r = v7._hdr(ws, r, ["SN", "Keep ON", "Keep ON (cont.)", "MO", "MO (cont.)", "Why keep it",
                        "Why (cont.)", "Why (cont.)", "Why (cont.)", "Risk if rolled",
                        "Risk if rolled (cont.)"], fill_hex=GREEN, height=28)
    risks = [
        "SRS capacity on a loaded 32T cell falls.",
        "SRS SINR estimate gets worse, which is the opposite of what pairing needs.",
        "Every remaining iBeam child silently dies with the master.",
        "UL tput and UL MU pairing would move for no DL gain.",
        "CCE / latency / small-packet behaviour change with no pairing benefit.",
        "UL IBLER on moving UEs would likely rise.",
        "You would be undoing the commercial baseline, not the trial.",
        "Nothing to undo — they never reached the air. Re-run later from sheet 17.",
    ]
    for i, ((keep, mo, why), risk) in enumerate(zip(KEEP, risks), 1):
        base = alt_fill(i)
        put(ws, r, 1, i, size=9, bold=True, fill_hex=base, h="center", v="center", border=True)
        merge(ws, r, 2, r, 3)
        put(ws, r, 2, keep, size=8, bold=True, fill_hex=PALE_GREEN, h="left", v="top", border=True)
        ws.cell(r, 3).border = thin
        merge(ws, r, 4, r, 5)
        put(ws, r, 4, mo, size=8, fill_hex=base, h="left", v="top", border=True)
        ws.cell(r, 5).border = thin
        merge(ws, r, 6, r, 9)
        put(ws, r, 6, why, size=8, fill_hex=WHITE, h="left", v="top", border=True)
        for c in (7, 8, 9):
            ws.cell(r, c).border = thin
        merge(ws, r, 10, r, 11)
        put(ws, r, 10, risk, size=8, fill_hex="FFF2CC", h="left", v="top", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = v7._height(
            [(v7._w(W, 2, 3), keep), (v7._w(W, 6, 9), why), (v7._w(W, 10, 11), risk)], cap=90)
        r += 1

    r = blank(ws, r, 8)
    r = subsection(ws, r, COLS, "Restore MML  —  if a night’s rollback hurts IBLER or drop",
                   fill_hex=NAVY2)
    restore = (
        "RB-1 undo:   MOD NRDUCELLSRS: NrDuCellId={NrDuCellId}, SrsDetectionAlgoSwitch=SRS_IC_SW-1;\n"
        "RB-2 undo:   MOD NRDUCELLDLMIMO: NrDuCellId={NrDuCellId}, "
        "DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-1&DL_MU_ANTI_INTRF_SCH_SW-1;\n"
        "RB-3 undo:   MOD NRDUCELLDLMIMO: NrDuCellId={NrDuCellId}, "
        "DLMuMimoSchSupplementSw=DL_MU_ANTI_INTRF_SCH_SW-1;\n"
        "RB-4 undo:   MOD NRDUCELLSRSMEAS: NrDuCellId={NrDuCellId}, "
        "SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-1;\n"
        "RB-5 undo:   MOD NRDUCELLULPCCONFIG: NrDuCellId={NrDuCellId}, "
        "UlPwrCtrlAlgoSwitch=SRS_JOINT_PC_SW-1;"
    )
    r = body(ws, r, COLS, restore, fill_hex=PALE_BLUE, size=9,
             height=v7._height([(sum(W), restore)], size=9, cap=140))
    return ws


def add_cross_links(wb):
    ws = wb[v7.S14]
    r = ws.max_row + 2
    r = blank(ws, r, 10)
    r = section(ws, r, bcs.COLS, "Trial result of this proposal  —  Ph1 executed 10–11 Sep 2026")
    r = note_bar(ws, r, bcs.COLS,
                 "18 switches were accepted; 4 were rejected with RETCODE 2147616329. DL user "
                 "throughput stayed flat and DL MU paired RB fell. Use sheet 16 for why, sheet 17 "
                 "for the full plan, and sheet 18 for the pairing rollback MML to send tonight.",
                 fill_hex="F8CBAD")
    for label, target in (
        (f"→  {v7.F_SHEET}   (what happened and why)", v7.F_SHEET),
        (f"→  {v7.A_SHEET}   (Ph1…Ph7 plan)", v7.A_SHEET),
        (f"→  {R_SHEET}   (pairing rollback MML, one switch per night)", R_SHEET),
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
    build_rollback(wb)
    add_cross_links(wb)


def main():
    bcs.main(out=OUT, zip_out=ZIP_OUT, with_prereq=True, extra=extra, version="v8.0")


if __name__ == "__main__":
    main()
