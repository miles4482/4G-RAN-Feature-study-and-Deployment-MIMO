"""
MIMO_Deployment v7.0 — three changes on top of v6.0:

  1. Sheet 14 Section 1: every MML row gains "Pre-requisite switch / parameter (must be ON
     first)" and "Pre-requisite MML (run BEFORE this line)", so the enable order is on the row.
     This is the correction for the 10-Sep-2026 execution error (RETCODE 2147616329).
  2. New sheet "16. Master Findings"  — what Ph1 actually did on 10–11 Sep 2026, what the
     gNodeB rejected, why DL user throughput stayed flat, why DL MU pairing fell, and a
     keep / roll-back / re-run verdict for every switch in the WO.
  3. New sheet "17. Action Plan Ph1-Ph7" — Ph1 complete, Ph2 corrective night with the MML in
     sequence, then Ph3…Ph7 each with its own pre-requisites and KPI exit gate.

Run:  python3 build_v7_findings.py
"""
import math
import os

from mimo_excel_style import *
from mimo_common import add_kpi_header, add_kpi
import build_combined_sheet as bcs
from build_combined_sheet import href_sheet, href_row

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "MIMO_Deployment_v7.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v7.0.zip")

F_SHEET = bcs.FINDINGS_SHEET      # "16. Master Findings"
A_SHEET = bcs.ACTION_SHEET        # "17. Action Plan Ph1-Ph7"
S14 = bcs.SHEET_NAME

COLS = 11
F_WIDTHS = [6, 21, 26, 29, 10, 14, 15, 33, 34, 14, 41]
A_WIDTHS = [9, 14, 30, 30, 34, 20, 20, 22, 20, 16, 18]

RED_HDR = "A93226"
BLUE_HDR = "5B9BD5"
YELLOW_HDR = "FFC000"

VERDICT_FILL = {
    "Keep": "D5F5E3",
    "Keep (master)": "D5F5E3",
    "Keep, watch": "FFF2CC",
    "Keep, re-check": "FFF2CC",
    "Keep — this is the one to keep": "D5F5E3",
    "ROLL BACK  (P0)": "F8CBAD",
    "Hold — rollback candidate": "FCE4D6",
    "RE-RUN in Ph2": "D6EAF8",
}

# ---------------------------------------------------------------- Ph1 evidence
# Executed 10–11 Sep 2026 on the trial cluster (from the WO "MO Name / Parameter ID /
# proposed Parameter Value / WO execution = Done" list).
PH1_DONE = [
    ("NRDUCellSrs", "SrsAlgoSwitch", "SRS_SINR_MEAS_OPT_SW-1", "AHR / SRS",
     "More accurate SRS SINR estimate. Feeds UL MCS and the DL weight calculation.",
     "No harm visible. Cannot be isolated — it landed together with 17 other bits.",
     "Keep",
     "Keep ON. Measurement-accuracy bit, not a pairing bit. No action in Ph2."),
    ("NRDUCellUlRank", "UlRankAlgoSw", "UL_RANK_FAST_DECREASE_SW-1", "UL rank",
     "Faster UL rank down-switch when SINR falls. Lower UL IBLER on moving UEs.",
     "UL was not reported in the trial evidence, so this is unmeasured.",
     "Keep",
     "Keep ON. Pull N.ChMeas.PUSCH.MCS.k and UL IBLER for 9–12 Sep to close the gap."),
    ("NRDUCellPusch", "PuschPerformanceSwitch", "PUSCH_CE_SINR_LEVEL_ENH_SW-1", "UL CE",
     "Better PUSCH channel estimation at low SINR. Cell-edge UL gain.",
     "UL was not reported — unmeasured.",
     "Keep",
     "Keep ON. Same UL counter pull as above."),
    ("NRDUCellSrs", "SrsDetectionAlgoSwitch", "SRS_IC_SW-1", "AHR / SRS",
     "SRS interference cancellation. Intended only for cells that are SRS-capacity limited.",
     "PRIME SUSPECT for the pairing loss. It was enabled in the SAME window as "
     "SRS_TIGHT_MULTIPLEXING_SW-1, which sheet 14 explicitly forbids. Two SRS detection "
     "changes at once degrade the SRS channel estimate that MU pairing depends on.",
     "ROLL BACK  (P0)",
     "Ph2 night 1, as the ONLY change: SrsDetectionAlgoSwitch=SRS_IC_SW-0. Then watch "
     "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB and N.SRS.NI.Avg for 24 h before anything else."),
    ("NRDUCellUlPcConfig", "UlPwrCtrlAlgoSwitch", "SRS_JOINT_PC_SW-1", "AHR / SRS",
     "Joint SRS power control. Lifts SRS SINR of far UEs.",
     "Neutral to positive on its own, but it was tuned to work alongside SRS_IC.",
     "Keep, re-check",
     "Keep ON through Ph2. If SRS NI is still high after the SRS_IC rollback, test "
     "SRS_JOINT_PC_SW-0 as a single change in Ph3."),
    ("NRDUCellFeatureSw", "HighPrecisionBeamSwitch", "ON", "iBeam 1.0",
     "Master switch. Every iBeam child bit below is inactive without it.",
     "Required. Nothing to conclude on its own.",
     "Keep (master)",
     "Never roll this back while any iBeam child bit is still 1 — roll the children back first."),
    ("NRDUCellSrsMeas", "SrsMeasOptSwitch", "SRS_BLIND_IS_MEAS_SW-1", "iBeam 1.0",
     "Blind interference-source SRS measurement. Sharper MU pairing decisions.",
     "Part of the pairing-gate group, so it contributed to fewer and stricter pairs.",
     "Keep, watch",
     "Keep in Ph2. Rollback candidate #3 only if pairing is still short of the gate after "
     "SRS_IC and the two DL MU gates have been dealt with."),
    ("NRDUCellSrs", "SrsAlgoSwitch", "SRS_TIGHT_MULTIPLEXING_SW-1", "AHR / SRS",
     "Tighter SRS multiplexing. More SRS capacity, which a loaded 32T cell needs.",
     "Conflicted with SRS_IC_SW-1 in the same window. The feature itself is the one CR01 "
     "chose and the one to keep.",
     "Keep",
     "Keep ON. SRS_IC is the bit that must go, not this one."),
    ("NRDUCellDlMimo", "DLMuMimoSchSupplementSw", "DL_MU_PRECISE_SCH_SW-1", "iBeam 1.0 / MU",
     "Precise MU scheduling. Pairs only UEs whose channel estimate is clean enough.",
     "By design this reduces paired RB when the SRS input is poor. Together with finding "
     "F-02 it explains most of the −61 % paired RB.",
     "Hold — rollback candidate",
     "Keep through Ph2. If 4/5/6-layer paired RB is still below 90 % of the 5–9 Sep baseline, "
     "set it to 0 on its own night in Ph3 — never together with DL_MU_ANTI_INTRF_SCH_SW."),
    ("NRDUCellDlMimo", "DLMuMimoSchSupplementSw", "DL_MU_ANTI_INTRF_SCH_SW-1", "iBeam 1.0 / MU",
     "Anti-interference MU scheduling. Rejects pairs that would interfere with each other.",
     "Same mechanism as the row above: fewer but cleaner pairs.",
     "Hold — rollback candidate",
     "Same treatment as DL_MU_PRECISE_SCH_SW. One switch per night, 24 h of counters between."),
    ("NRDUCellDlSch", "DlSchAlgoSwitch", "TAIL_PKT_MCS_OPT_SW-1", "iBeam 1.0",
     "Tail-packet MCS optimisation. Helps small-packet latency.",
     "Neutral on throughput. Not a pairing bit.",
     "Keep",
     "Keep ON. Watch N.Traffic.DL.RlcFirstPktDelay.Time."),
    ("NRDUCellDlSch", "DlSchAlgoSwitch", "RES_BASED_DL_ADAPT_SCH_SW-1", "iBeam 1.0",
     "Resource-based DL adaptive scheduling.",
     "Neutral.",
     "Keep",
     "Keep ON."),
    ("NRDUCellDlSch", "DlSchAlgoSwitch", "DL_RLC_STAT_RPT_MERGE_SCH_SW-1", "iBeam 1.0",
     "Merges RLC status reports. Less air-interface overhead.",
     "Neutral.",
     "Keep",
     "Keep ON."),
    ("NRDUCellPdcch", "PdcchAlgoSwitch", "PDCCH_AGG_LVL_COMPR_SW-1", "iBeam 1.0",
     "PDCCH aggregation-level compression. Frees CCE for PDSCH scheduling.",
     "Not measurable — no CCE counter was supplied. Its companion threshold was also not set.",
     "Keep, watch",
     "Keep ON but add the missing line AggLvlComprCceUsageThld=60 in Ph2, and pull "
     "N.CCE.Used.Avg / N.CCE.DL.AggLvl*."),
    ("NRDUCellFeatureSw", "MimoFeatureSwitch", "UL_LOW_NOISE_SW-1", "UL Boosting 1.0",
     "UL low-noise receiver. UL SINR gain. Master for the three UL MU bits below.",
     "UL was not reported — unmeasured.",
     "Keep",
     "Keep ON. It is the pre-requisite of the next three rows."),
    ("NRDUCellUlMimo", "UlMuMimoAlgoSwitch", "UL_MU_GRP_PAIR_SW-1", "UL Boosting 1.0",
     "UL MU group pairing.",
     "UL was not reported — unmeasured.",
     "Keep",
     "Keep ON. Pull N.ChMeas.MIMO.UL.Pair.Layer / Pair.PRB."),
    ("NRDUCellUlMimo", "UlMuMimoAlgoSwitch", "DIFF_WAVEFORM_PAIR_SW-1", "UL Boosting 1.0",
     "Pairs UEs that use different waveforms (CP-OFDM against DFT-s-OFDM).",
     "UL was not reported — unmeasured.",
     "Keep",
     "Keep ON."),
    ("NRDUCellUlMimo", "UlMuMimoAlgoSwitch", "UL_CORR_ACCELERATION_SW-1", "UL Boosting 1.0",
     "Faster UL correlation-matrix calculation.",
     "UL was not reported — unmeasured.",
     "Keep",
     "Keep ON."),
]

# Same CR, but the gNodeB refused the command — these four never reached the air.
PH1_REJECTED = [
    ("NRDUCellBeamAlgo", "WeightAlgoSwitch", "SRS_WEIGHT_ESTIMATE_SW-1", "Beam / weights",
     "THE main DL throughput lever: build the DL beamforming weights from SRS instead of "
     "from the PMI codebook. This is what was supposed to lift DL user throughput.",
     "Never active. It was sent inside the same MOD as BEAM_TRACKING_SW, and MOD is atomic, "
     "so the rejection killed this bit too.",
     "RE-RUN in Ph2",
     "Re-send on its OWN MML line in Ph2. Nothing else in that line."),
    ("NRDUCellBeamAlgo", "ChannelOptAlgoSwitch", "BEAM_SELECT_OPT_SW-1", "Beam / weights",
     "Beam-selection optimisation. Picks a better serving beam per UE.",
     "Never active — collateral damage from the same rejected command.",
     "RE-RUN in Ph2",
     "Re-send on its own MML line in Ph2."),
    ("NRDUCellBeamAlgo", "BeamOptAlgoSwitch", "BEAM_TRACKING_SW-1", "Beam / weights",
     "Beam tracking. Keeps the beam on moving UEs, which is where MCS gain comes from.",
     "REJECTED — RETCODE 2147616329. BEAM_TRACKING_SW may only be selected while "
     "DlCoverageAlgoSwitch = SUPER_COVERAGE_SW-1 in the matching NRDUCellChnCovAlgo MO. "
     "That pre-requisite was never sent, so the gNodeB refused the whole line.",
     "RE-RUN in Ph2",
     "Ph2 step 2.2: MOD NRDUCELLCHNCOVALGO DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1 FIRST, "
     "then BEAM_TRACKING_SW-1 on its own line."),
    ("NRDUCellBeamAlgo", "BeamOptAlgoSwitch", "INTELLIGENT_BEAM_SELECTION_SW-1", "Beam / weights",
     "Intelligent beam selection. Works together with beam tracking.",
     "Never active — same rejected command.",
     "RE-RUN in Ph2",
     "Re-send after BEAM_TRACKING_SW has been accepted, on its own line."),
]

# N.ChMeas.MIMO.DL.MuPairing.{k}Layer.RB — read off the trial pivot chart.
PAIRING = [
    ("1 layer", 0.70, 0.50, 0.42, 0.30, 0.25),
    ("2 layers", 3.20, 3.00, 2.50, 1.20, 1.80),
    ("3 layers", 3.40, 3.10, 3.00, 1.60, 2.80),
    ("4 layers", 7.60, 8.30, 5.90, 2.90, 4.30),
    ("5 layers", 0.60, 0.50, 0.35, 0.20, 0.25),
    ("6 layers", 0.80, 0.90, 0.50, 0.15, 0.20),
    ("7 layers", 0.05, 0.05, 0.02, 0.01, 0.01),
    ("8–16 layers", 0.02, 0.03, 0.01, 0.00, 0.00),
]

FINDINGS = [
    ("F-01", "CRITICAL",
     "The DL throughput levers never reached the air, so “no DL gain” is the expected result",
     "MML log 2026-09-10 15:27:14 returns RETCODE 2147616329 for the NRDUCELLBEAMALGO MOD on "
     "NrDuCellId=102. None of SRS_WEIGHT_ESTIMATE_SW, BEAM_SELECT_OPT_SW, BEAM_TRACKING_SW or "
     "INTELLIGENT_BEAM_SELECTION_SW appears in the executed “Done” list.",
     "SRS-based weights and beam tracking are the two bits that actually move DL MCS and DL user "
     "throughput. What did land was the interference / pairing-gate package, which does not add "
     "throughput on its own. The trial therefore did not test the DL hypothesis at all.",
     "Ph2: send DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1 first, then the four beam bits one per "
     "MML line. Only then re-measure DL user throughput."),
    ("F-02", "CRITICAL",
     "SRS_IC_SW and SRS_TIGHT_MULTIPLEXING_SW were enabled in the same window",
     "Executed list rows 4 and 8: SrsDetectionAlgoSwitch=SRS_IC_SW-1 and "
     "SrsAlgoSwitch=SRS_TIGHT_MULTIPLEXING_SW-1, both stamped Done in the same WO. Sheet 14 "
     "carries the explicit rule “do not add SRS_IC the same night as SRS_TIGHT_MULTIPLEXING”.",
     "Both switches change how SRS is detected and multiplexed. Together the SRS channel estimate "
     "degrades, and everything downstream of SRS — DL weights, MU pairing eligibility, UL rank — "
     "degrades with it. This is the most likely direct cause of the pairing collapse.",
     "Ph2 night 1, single change: SRS_IC_SW-0. Keep tight multiplexing. Confirm with "
     "N.SRS.NI.Avg and N.UL.SRS.PreSINR.Index* over 9–12 Sep."),
    ("F-03", "MAJOR",
     "MU pairing gates were tightened at the same time as the SRS input was degraded",
     "DL_MU_PRECISE_SCH_SW-1 and DL_MU_ANTI_INTRF_SCH_SW-1 both went in. Paired RB fell hardest "
     "on the high layers: 4-layer 8.30 → 2.90 RB (−65 %), 6-layer 0.90 → 0.15 RB (−83 %).",
     "Both switches are quality gates — they reject pairs whose channel estimate is not clean. "
     "Fed with degraded SRS they reject almost everything, so paired RB falls. The pairing drop is "
     "the two findings above acting together, not a defect in either feature.",
     "Do not roll these back first. Fix SRS (F-02), re-measure, and only then decide — one switch "
     "per night, 24 h apart."),
    ("F-04", "MAJOR",
     "Nothing in Ph1 could increase high-layer pairing, so the direction of travel was one-way down",
     "The executed list contains no MMIMO_MULTILAYER_ENHANCE_SW, no MU_RANK_BOOSTING_SW, no "
     "MU_MIMO_PAIRING_PREFERRED_SW. The cluster already carries the LAYER_16 quota.",
     "Ph1 contained only gates that can reduce pairing and no feature that can add layers. Even "
     "with healthy SRS, N.ChMeas.MIMO.DL.MuPairing.4…16Layer.RB had no mechanism to rise.",
     "Ph4: enable HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1 as the master, then its children "
     "one per night. That is the phase that is supposed to spend the LAYER_16 quota."),
    ("F-05", "MAJOR",
     "Several switches were bundled into one MML line, and MOD is atomic",
     "The failed command carried four bits at once: WeightAlgoSwitch, ChannelOptAlgoSwitch and two "
     "BeamOptAlgoSwitch options in a single MOD NRDUCELLBEAMALGO.",
     "One invalid option rejects the entire command, so three healthy switches were lost with the "
     "one that had a missing pre-requisite. It also makes the result impossible to attribute: when "
     "a KPI moves, nobody can say which bit moved it.",
     "Execution rule from v7.0 onward: one switch per MML line, in the sequence given by the new "
     "pre-requisite columns on sheet 14. Sheet 17 Ph2 is written that way."),
    ("F-06", "INFO",
     "The real execution window is 10–11 Sep 2026, not 11 Sep only",
     "The rejected MML is stamped 2026-09-10 15:27:14, while the WO reports 11 Sep. The pairing "
     "counters already sag on 10 Sep (all-layer paired RB ≈ 12.7 against ≈ 16.3 on 9 Sep).",
     "Using 11 Sep as the only execution date makes 10 Sep look like a valid baseline day, which "
     "would understate the regression by roughly a third.",
     "Use 5 Sep and 9 Sep as the baseline, 10–11 Sep as the execution window and 12 Sep as D+1 in "
     "every comparison from now on."),
    ("F-07", "INFO",
     "12 Sep recovers part of the loss but stays well below baseline — it is not self-healing",
     "All-layer paired RB ≈ 16.3 on 9 Sep, ≈ 6.35 on 11 Sep (−61 %), ≈ 9.6 on 12 Sep (−41 %).",
     "Part of the 11 Sep dip is load or transient, but 12 Sep is still 41 % down. Waiting will not "
     "close the gap, so a configuration change is required.",
     "Do not wait for further recovery. Run Ph2 at the next maintenance window."),
    ("F-08", "INFO",
     "Evidence gap: the counters that would prove the root cause were not supplied",
     "The trial evidence is the DL MU pairing pivot plus a statement that DL user throughput is "
     "flat. There is no SRS NI, no SRS pre-SINR, no DL IBLER, no CCE and no PRB load series.",
     "Without SRS NI the SRS_IC conflict stays a strong hypothesis. Without PRB load the pairing "
     "drop cannot be separated from a simple traffic change. Without IBLER it is unknown whether "
     "the stricter gates bought quality for the RB they gave up.",
     "Pull N.SRS.NI.Avg, N.UL.SRS.PreSINR.Index*, DL IBLER, N.CCE.Used.Avg and "
     "N.PRB.DL.Used.Avg for 5–12 Sep, trial and neighbour control, before Ph2 sign-off."),
]

# ------------------------------------------------------------------ action plan
# (phase, window, objective, content, prereq to verify, exit gate, rollback, status)
PHASES = [
    ("Ph1", "10–11 Sep 2026\n(done)",
     "Land the SRS / UL Boosting / iBeam 1.0 interference package and the beam-weight levers.",
     "18 switches accepted (see sheet 16 Section 2). 4 switches rejected: SRS_WEIGHT_ESTIMATE_SW, "
     "BEAM_SELECT_OPT_SW, BEAM_TRACKING_SW, INTELLIGENT_BEAM_SELECTION_SW.",
     "Not checked before execution — that is what caused the rejection.",
     "MISSED. DL user throughput flat; all-layer DL MU paired RB −61 % on 11 Sep, −41 % on 12 Sep.",
     "Not required as a whole. Ph2 rolls back one switch (SRS_IC_SW) only.",
     "COMPLETE"),
    ("Ph2", "Next window\n(night 1 + night 2)",
     "Undo the one switch that broke SRS, then finally land the four rejected DL levers in "
     "sequence. This is the phase that tests the original DL throughput hypothesis.",
     "Night 1: SRS_IC_SW-0 alone. Night 2 (only if night 1 is green): SUPER_COVERAGE_SW-1, then "
     "SRS_WEIGHT_ESTIMATE_SW-1, BEAM_SELECT_OPT_SW-1, BEAM_TRACKING_SW-1, "
     "INTELLIGENT_BEAM_SELECTION_SW-1 — one switch per MML line. Plus the missed "
     "AggLvlComprCceUsageThld=60.",
     "LST NRDUCELLSRS (tight MUX = 1, IC = 1 before the change); "
     "LST NRDUCELLCHNCOVALGO (DlCoverageAlgoSwitch); "
     "LST NRDUCELLBEAMALGO (BeamPerceiveMode = DISTRIBUTED_MODE).",
     "All-layer DL MU paired RB ≥ 14.7 RB (90 % of the 16.3 RB baseline) AND high-layer (≥4L) "
     "≥ 8.7 RB AND DL user throughput at or above the control trend AND SRS NI flat.",
     "Night 1: SRS_IC_SW-1 back. Night 2: the four beam bits back to 0, SUPER_COVERAGE_SW back to "
     "its pre-change value. Never touch HighPrecisionBeamSwitch.",
     "NEXT"),
    ("Ph3", "Ph2 + 7 days",
     "Decide the fate of the two DL MU quality gates, then add SSB beam adaptation.",
     "Only if Ph2 missed its pairing gate: DL_MU_PRECISE_SCH_SW-0 on night 1, then (if still "
     "short) DL_MU_ANTI_INTRF_SCH_SW-0 on night 2. Then SSB_BEAM_ADAPT_SW-1 and "
     "SSB_BEAM_VERTICAL_COV_IMP_SW-1.",
     "Tilt and azimuth must match the RF design — no Tilt=255 cell in the trial set. "
     "DL_INITIAL_BEAM_SELECT_SW must already be 1.",
     "Pairing gate met with the smallest possible rollback; N.User.OptimalSSBBeam.Avg rises; "
     "drop, HOSR and RRC success unchanged against control.",
     "Re-enable the rolled-back MU gate, or SSB_BEAM_ADAPT_SW-0 — whichever that night touched.",
     "PLANNED"),
    ("Ph4", "Ph3 + 7 days",
     "Spend the LAYER_16 quota. This is the phase that is actually supposed to raise high-layer "
     "MU pairing and DL cell capacity.",
     "Night 1: HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1 (master). Then one child per night: "
     "MU_RANK_BOOSTING_SW, MU_MIMO_PAIRING_PREFERRED_SW, SRS_MEAS_ACCELERATING_SW, SRS_BLIND_IS_SW.",
     "LST LICENSE for NR0S0DLEPU00 layer-capacity units; DL_MU_MIMO_SW = 1; "
     "MaxMimoLayerNum already LAYER_16 (never downgrade it).",
     "N.ChMeas.MIMO.DL.Transmission.Layer.Max rises; high-layer paired RB above the 9 Sep "
     "baseline; DL IBLER inside the planned band.",
     "Children back to 0 first, master last.",
     "PLANNED"),
    ("Ph5", "Ph4 + 7 days",
     "Capacity upgrade layer: AHR Capacity Upgrade 2.0 and iBeam 2.0.",
     "AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1, then PDCCH_MULTI_DIM_JOINT_SCH_SW-1. "
     "Then HighPrecisionBeamPhase2Sw=ON and its children one per night.",
     "AHR_PHASE1_SW and AHR_EXP_TURBO_PHASE2_SW must both already be 1; iBeam 1.0 must be "
     "KPI-green from Ph2/Ph3; licences checked with LST LICENSE.",
     "DL user throughput above the post-Ph2 level; CQI and IBLER stable; CCE not blocking.",
     "iBeam 2.0 children, then Phase2Sw=OFF, then the AHR CU bit.",
     "PLANNED"),
    ("Ph6", "Ph5 + 7 days",
     "UL side and the remaining weight variants.",
     "UL_LOW_NOISE_PHASE2_SW-1, then PUSCH_COORD_PWR_CTRL_SW-1. Separately, trial "
     "PMI_WEIGHT_OPT_SW / OPEN_LOOP_WEIGHT_OPT_SW against the now-proven SRS weight baseline.",
     "DSP CLKTST must report inter-gNB time sync before the coordinated power-control bit. "
     "BeamPerceiveMode = DISTRIBUTED_MODE for the weight variants.",
     "UL user throughput and UL IBLER improve or hold; DL untouched.",
     "One switch back per night, newest first.",
     "PLANNED"),
    ("Ph7", "Ph6 + 14 days",
     "Cluster rollout of the proven set.",
     "Apply only the green set from Ph2–Ph6 to the remaining 32T cells, in the same MML sequence. "
     "Fix Tilt=255 cells before they join. 2T2R indoor cells stay out.",
     "Per-cell pre-check with the sheet 14 pre-requisite columns; RF tilt/azimuth verified; "
     "no VSWR alarm.",
     "Cluster DL and UL user throughput at or above the trial-cluster result; no KPI regression on "
     "any gNB against its own pre-week.",
     "Per-gNB rollback of that gNB's CME file only.",
     "PLANNED"),
]

# Ph2 corrective sequence: (seq, mml, prereq switch, prereq mml, verify, counter, gate)
PH2_SEQ = [
    ("2.1", "MOD NRDUCELLSRS: NrDuCellId=102, SrsDetectionAlgoSwitch=SRS_IC_SW-0;",
     "None. This is a rollback — it removes the switch that conflicts with "
     "SRS_TIGHT_MULTIPLEXING_SW-1.",
     "— (run first, alone, nothing else this night)",
     "LST NRDUCELLSRS: NrDuCellId=102;  →  SrsAlgoSwitch must still show "
     "SRS_TIGHT_MULTIPLEXING_SW-1 and SrsDetectionAlgoSwitch must now show SRS_IC_SW-0",
     "N.SRS.NI.Avg, N.UL.SRS.PreSINR.Index*, N.ChMeas.MIMO.DL.MuPairing.kLayer.RB",
     "Night 1 gate after 24 h: all-layer paired RB ≥ 14.7 RB and SRS NI back to the 5–9 Sep level. "
     "Stop here for a full day — do not continue to 2.2 on the same night."),
    ("2.2", "MOD NRDUCELLCHNCOVALGO: NrDuCellId=102, DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1;",
     "None — this switch is itself the missing pre-requisite. It is the correction for "
     "RETCODE 2147616329.",
     "— (this is the pre-requisite line; it must be step 1 of night 2)",
     "LST NRDUCELLCHNCOVALGO: NrDuCellId=102;  →  DlCoverageAlgoSwitch must show "
     "SUPER_COVERAGE_SW-1 BEFORE step 2.5 is sent",
     "N.ThpVol.DL, N.ChMeas.PDSCH.MCS.k, N.UECntx.AbnormRel",
     "Command returns RETCODE = 0. Super coverage changes DL coverage behaviour, so watch drop and "
     "HOSR for 30 minutes before continuing."),
    ("2.3", "MOD NRDUCELLBEAMALGO: NrDuCellId=102, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1;",
     "BeamPerceiveMode=DISTRIBUTED_MODE and AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-1 "
     "(both already live on this cluster).",
     "LST NRDUCELLBEAMALGO: NrDuCellId=102;   // confirm BeamPerceiveMode=DISTRIBUTED_MODE",
     "LST NRDUCELLBEAMALGO: NrDuCellId=102;  →  WeightAlgoSwitch shows SRS_WEIGHT_ESTIMATE_SW-1",
     "N.ChMeas.PDSCH.MCS.k, N.PDSCH.InitTbDl.Rank*, N.ThpVol.DL, DL IBLER",
     "This is the main DL lever. Expect the DL MCS distribution to upshift. It is sent on its own "
     "line precisely because Ph1 lost it inside a bundled command."),
    ("2.4", "MOD NRDUCELLBEAMALGO: NrDuCellId=102, ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1;",
     "HighPrecisionBeamSwitch=ON (already ON from Ph1).",
     "LST NRDUCELLFEATURESW: NrDuCellId=102;   // HighPrecisionBeamSwitch must be ON",
     "LST NRDUCELLBEAMALGO: NrDuCellId=102;  →  ChannelOptAlgoSwitch shows BEAM_SELECT_OPT_SW-1",
     "N.User.OptimalSSBBeam.Avg, N.ThpVol.DL, DL IBLER",
     "Own MML line. If it fails, steps 2.5 and 2.6 must not be sent."),
    ("2.5", "MOD NRDUCELLBEAMALGO: NrDuCellId=102, BeamOptAlgoSwitch=BEAM_TRACKING_SW-1;",
     "MANDATORY: step 2.2 must already be committed — NRDUCellChnCovAlgo.DlCoverageAlgoSwitch = "
     "SUPER_COVERAGE_SW-1. Without it this exact line returns RETCODE 2147616329, which is what "
     "happened on 10 Sep 2026.",
     "MOD NRDUCELLCHNCOVALGO: NrDuCellId=102, DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1;   "
     "// step 2.2, must be committed before this line",
     "LST NRDUCELLBEAMALGO: NrDuCellId=102;  →  BeamOptAlgoSwitch shows BEAM_TRACKING_SW-1 and "
     "RETCODE = 0",
     "N.ChMeas.PDSCH.MCS.k on mobility samples, N.ThpVol.DL, HOSR, N.UECntx.AbnormRel",
     "The line that failed in Ph1. Send it ALONE — never together with step 2.6."),
    ("2.6", "MOD NRDUCELLBEAMALGO: NrDuCellId=102, "
            "BeamOptAlgoSwitch=BEAM_TRACKING_SW-1&INTELLIGENT_BEAM_SELECTION_SW-1;",
     "Step 2.5 accepted (so BEAM_TRACKING_SW is already 1, and SUPER_COVERAGE_SW behind it).",
     "Steps 2.2 and 2.5 in that order.",
     "LST NRDUCELLBEAMALGO: NrDuCellId=102;  →  both options present in BeamOptAlgoSwitch",
     "N.ThpVol.DL, N.User.OptimalSSBBeam.Avg, DL IBLER",
     "BeamOptAlgoSwitch is a bit-field, so both options must be listed to keep tracking ON while "
     "adding intelligent selection. Send it only after 2.5 returned RETCODE = 0."),
    ("2.7", "MOD NRDUCELLPDCCHALGO: NrDuCellId=102, AggLvlComprCceUsageThld=60;",
     "PdcchAlgoSwitch=PDCCH_AGG_LVL_COMPR_SW-1 (already set in Ph1, but its threshold was "
     "never sent).",
     "LST NRDUCELLPDCCHALGO: NrDuCellId=102;   // PDCCH_AGG_LVL_COMPR_SW must be 1",
     "LST NRDUCELLPDCCHALGO: NrDuCellId=102;  →  AggLvlComprCceUsageThld = 60",
     "N.CCE.Used.Avg, N.CCE.DL.AggLvl*, N.CCE.DL.AllocReq.Num",
     "Closes a gap left by the Ph1 work order. Equivalent CCE used should fall or hold; PDCCH "
     "blocking must not rise."),
]

GOLDEN_RULES = [
    "One switch per MML line. MOD is atomic — in Ph1 four bits went in one command and one bad "
    "bit rejected all four.",
    "Read the pre-requisite columns on sheet 14 before writing the work order, and send the "
    "pre-requisite line first. BEAM_TRACKING_SW needs DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1.",
    "Never two suspect switches in one window. SRS_IC_SW and SRS_TIGHT_MULTIPLEXING_SW in the "
    "same night is exactly what made the Ph1 pairing result unattributable.",
    "Confirm every line with an LST before the next line is sent, and keep the RETCODE output. "
    "A Done in a work order is not proof that the switch reached the air.",
    "Roll back children before masters: iBeam child bits before HighPrecisionBeamSwitch, "
    "multilayer children before MMIMO_MULTILAYER_ENHANCE_SW.",
    "Every phase keeps a neighbour 32T control cluster and compares the same busy hour. "
    "Baseline for this campaign is 5 Sep and 9 Sep 2026.",
    "Never downgrade what is already good: MaxMimoLayerNum stays LAYER_16, and SU / MU / AHR "
    "Phase1 + Turbo stay ON.",
]


def _hdr(ws, r, titles, fill_hex=NAVY, height=30):
    for i, t in enumerate(titles, 1):
        put(ws, r, i, t, size=9, bold=True, color=WHITE, fill_hex=fill_hex,
            h="center", v="center", border=True)
    ws.row_dimensions[r].height = height
    return r + 1


def _row(ws, r, vals, fills, bolds=(), center=(), height=44, size=8):
    for i, v in enumerate(vals, 1):
        put(ws, r, i, v, size=size, bold=(i in bolds), fill_hex=fills[i - 1],
            h="center" if i in center else "left", v="top", border=True)
    ws.row_dimensions[r].height = height
    return r + 1


def _rows_needed(text, width_units, size=8):
    """Wrapped line count for `text` in a column block `width_units` wide at `size` pt."""
    cpl = max(8.0, width_units * 11.0 / size * 0.92)
    return sum(max(1, math.ceil(len(p) / cpl)) for p in str(text or "").split("\n"))


def _height(blocks, size=8, pad=9, cap=210, floor=26):
    """Row height in points from (width_units, text) blocks — the tallest block wins."""
    lines = max((_rows_needed(t, w, size) for w, t in blocks), default=1)
    return min(cap, max(floor, lines * (size + 3.4) + pad))


def _w(widths, c1, c2=None):
    return sum(widths[c1 - 1:(c2 or c1)])


def build_findings(wb):
    ws = wb.create_sheet(F_SHEET, 16)
    setup_sheet(ws, F_SHEET)
    set_widths(ws, F_WIDTHS)
    ws.oddHeader.left.text = "5G MIMO — Master Findings, Ph1 trial 10–11 Sep 2026"
    ws.oddFooter.left.text = FOOTER_LEFT

    r = banner(ws, 1, COLS,
               "  Master Findings  —  Ph1 MIMO trial on the 32T cluster,  executed 10–11 Sep 2026")
    r = note_bar(ws, r, COLS,
                 "Input: the executed work order (18 switches, all marked Done), the "
                 "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB pivot for 5 / 9 / 10 / 11 / 12 Sep, and the MML "
                 "log of 2026-09-10 15:27:14 that returned RETCODE 2147616329. "
                 "Reported result: no visible improvement in DL user throughput, and high-layer "
                 "DL MU pairing down.")
    r = blank(ws, r, 6)

    # ---------- Section 1: the answer in three lines
    r = section(ws, r, COLS, "Section 1.  Bottom line")
    bottom = (
             "The trial did not fail — it never ran the experiment it was designed to run.\n\n"
             "The four switches that produce DL throughput gain (SRS_WEIGHT_ESTIMATE_SW, "
             "BEAM_SELECT_OPT_SW, BEAM_TRACKING_SW, INTELLIGENT_BEAM_SELECTION_SW) were all sent in "
             "ONE MOD NRDUCELLBEAMALGO command. The gNodeB rejected that command with "
             "RETCODE 2147616329 because BEAM_TRACKING_SW requires "
             "DlCoverageAlgoSwitch = SUPER_COVERAGE_SW-1 in NRDUCellChnCovAlgo, which was never "
             "sent. MOD is atomic, so all four bits were lost — and none of them appears in the "
             "executed list. Flat DL user throughput is the expected outcome of that.\n\n"
             "What did land was the interference and pairing-gate package, and inside it two SRS "
             "detection changes in the same window: SRS_IC_SW-1 together with "
             "SRS_TIGHT_MULTIPLEXING_SW-1. Sheet 14 forbids that combination. Degraded SRS plus the "
             "stricter DL MU gates (DL_MU_PRECISE_SCH_SW, DL_MU_ANTI_INTRF_SCH_SW) is what pushed "
             "high-layer MU pairing down — 4-layer paired RB fell about 65 % and 6-layer about 83 % "
             "against the 9 Sep baseline.\n\n"
             "So: one rollback (SRS_IC_SW-0), one missing pre-requisite (SUPER_COVERAGE_SW-1), and "
             "a re-run of the four beam bits one per MML line. That is Ph2 on sheet 17.")
    r = body(ws, r, COLS, bottom, fill_hex=PALE_GOLD, size=11,
             height=_height([(sum(F_WIDTHS), bottom)], size=11, cap=320))
    r = blank(ws, r, 8)

    # ---------- Section 2: what was executed
    r = section(ws, r, COLS, "Section 2.  What was executed and accepted  (18 switches, WO = Done)")
    r = note_bar(ws, r, COLS,
                 "Verdict column: Keep = leave ON, no action. Keep, watch = leave ON but a counter is "
                 "still open. Hold — rollback candidate = only after the SRS fix has been measured. "
                 "ROLL BACK (P0) = the one change Ph2 night 1 must make, on its own.")
    r = _hdr(ws, r, ["SN", "MO Name", "Parameter ID", "Parameter value sent", "WO", "Landed on air?",
                     "Feature family", "What it was expected to do",
                     "What the trial actually shows", "Verdict", "Action to take now"])
    for i, (mo, param, val, fam, exp, obs, verdict, act) in enumerate(PH1_DONE, 1):
        vf = VERDICT_FILL.get(verdict, WHITE)
        base = alt_fill(i)
        fills = [base, base, base, base, "D5F5E3", "D5F5E3", base, base,
                 "F8CBAD" if verdict.startswith("ROLL") else base, vf, vf]
        r = _row(ws, r, [i, mo, param, val, "Done", "Yes", fam, exp, obs, verdict, act],
                 fills, bolds={3, 4, 10}, center={1, 5, 6, 10},
                 height=_height([(F_WIDTHS[7], exp), (F_WIDTHS[8], obs),
                                 (F_WIDTHS[10], act), (F_WIDTHS[3], val)]))

    r = blank(ws, r, 8)
    r = subsection(ws, r, COLS,
                   "Sent in the same CR but REJECTED by the gNodeB — these four never reached the air",
                   fill_hex=RED_HDR)
    r = note_bar(ws, r, COLS,
                 "MOD NRDUCELLBEAMALGO: NRDUCELLID=102, WEIGHTALGOSWITCH=SRS_WEIGHT_ESTIMATE_SW-1, "
                 "CHANNELOPTALGOSWITCH=BEAM_SELECT_OPT_SW-1, "
                 "BEAMOPTALGOSWITCH=INTELLIGENT_BEAM_SELECTION_SW-1&BEAM_TRACKING_SW-1;   →   "
                 "RETCODE = 2147616329   (2026-09-10 15:27:14, DHGUL66).  One invalid option, four "
                 "switches lost.", fill_hex="F8CBAD")
    r = _hdr(ws, r, ["SN", "MO Name", "Parameter ID", "Parameter value sent", "WO", "Landed on air?",
                     "Feature family", "What it was expected to do",
                     "Why it did not land", "Verdict", "Action to take now"], fill_hex=RED_HDR)
    for i, (mo, param, val, fam, exp, obs, verdict, act) in enumerate(PH1_REJECTED, 1):
        base = "FDEDEC" if i % 2 else "FCE4E4"
        fills = [base] * 4 + ["FFF2CC", "F8CBAD"] + [base, base, "F8CBAD", "D6EAF8", "D6EAF8"]
        r = _row(ws, r, [f"R{i}", mo, param, val, "Done\n(in the WO)", "NO — rejected",
                         fam, exp, obs, verdict, act],
                 fills, bolds={3, 6, 10}, center={1, 5, 6, 10},
                 height=_height([(F_WIDTHS[7], exp), (F_WIDTHS[8], obs),
                                 (F_WIDTHS[10], act)]))

    r = blank(ws, r, 8)

    # ---------- Section 3: KPI evidence
    r = section(ws, r, COLS, "Section 3.  KPI evidence  —  N.ChMeas.MIMO.DL.MuPairing.kLayer.RB")
    r = note_bar(ws, r, COLS,
                 "Values read off the trial pivot chart, so treat them as approximate and confirm the "
                 "exact figures from a MAE export. Baseline = 5 Sep and 9 Sep. Execution = 10–11 Sep "
                 "(the MML log is stamped 10 Sep, so 10 Sep is NOT a clean baseline day). D+1 = 12 Sep. "
                 "Δ columns are against 9 Sep.")
    r = _hdr(ws, r, ["SN", "MU layer depth", "5 Sep\n(baseline)", "9 Sep\n(baseline)",
                     "10 Sep\n(exec day 1)", "11 Sep\n(exec day 2)", "12 Sep\n(D+1)",
                     "Δ 11 Sep vs 9 Sep", "Δ 12 Sep vs 9 Sep", "Verdict",
                     "Reading"], fill_hex=TEAL, height=34)

    def pct(now, base):
        if not base:
            return "n/a"
        return f"{(now - base) / base * 100:+.0f} %"

    readings = {
        "1 layer": "Single-layer pairing is small and noisy — not the story.",
        "2 layers": "Down 60 %. Even the easy pairs are being rejected.",
        "3 layers": "Down 48 %, recovers most of the way by 12 Sep.",
        "4 layers": "THE headline. 4-layer pairing carries most of the paired RB on this cluster "
                    "and it lost about two thirds.",
        "5 layers": "Down 60 %. Same mechanism as 4-layer.",
        "6 layers": "Worst hit at −83 %. The deeper the pairing, the cleaner the SRS has to be — "
                    "which is exactly what F-02 damaged.",
        "7 layers": "Effectively zero before and after. Nothing to conclude.",
        "8–16 layers": "Never used on this cluster. Ph4 is the phase that would change that.",
    }
    for i, (lay, d5, d9, d10, d11, d12) in enumerate(PAIRING, 1):
        base = alt_fill(i)
        big = d9 >= 0.5
        vd = "Regression" if (big and d11 < d9 * 0.9) else "Flat / noise"
        vf = "F8CBAD" if vd == "Regression" else "EAECEE"
        fills = [base, base, PALE_GREEN, PALE_GREEN, "FFF2CC", "F8CBAD", "FCE4D6",
                 "F8CBAD", "FCE4D6", vf, base]
        r = _row(ws, r,
                 [i, lay, f"{d5:.2f}", f"{d9:.2f}", f"{d10:.2f}", f"{d11:.2f}", f"{d12:.2f}",
                  pct(d11, d9), pct(d12, d9), vd, readings[lay]],
                 fills, bolds={2, 8, 10}, center={1, 3, 4, 5, 6, 7, 8, 9, 10}, height=32)

    tot = [sum(x[i] for x in PAIRING) for i in range(1, 6)]
    hi = [sum(x[i] for x in PAIRING if x[0] not in ("1 layer", "2 layers", "3 layers"))
          for i in range(1, 6)]
    for lbl, vals, txt in (
        ("TOTAL  all layers", tot,
         "All-layer paired RB. Ph2 night-1 gate is 90 % of the 9 Sep figure."),
        ("TOTAL  high layers (≥4L)", hi,
         "High-layer paired RB — the part the user flagged. Ph2 gate is 90 % of the 9 Sep figure."),
    ):
        r = _row(ws, r,
                 ["Σ", lbl, f"{vals[0]:.2f}", f"{vals[1]:.2f}", f"{vals[2]:.2f}",
                  f"{vals[3]:.2f}", f"{vals[4]:.2f}", pct(vals[3], vals[1]),
                  pct(vals[4], vals[1]), "Regression", txt],
                 ["FFC000"] * 11, bolds=set(range(1, 12)),
                 center={1, 3, 4, 5, 6, 7, 8, 9, 10}, height=32)

    r = blank(ws, r, 6)
    r = body(ws, r, COLS,
             "DL user throughput: reported flat, no figure supplied. That is the expected result of "
             "finding F-01 — the two switches that move DL throughput (SRS-based weights and beam "
             "tracking) were never active during the trial. Do not conclude anything about the DL "
             "features from this window; the DL measurement only becomes valid after Ph2 step 2.6.",
             height=60, fill_hex="FFF2CC", size=10)
    r = blank(ws, r, 8)

    # ---------- Section 4: findings
    r = section(ws, r, COLS, "Section 4.  Findings and root cause")
    r = _hdr(ws, r, ["ID", "Severity", "Finding", "Finding (cont.)", "Evidence",
                     "Evidence (cont.)", "Evidence (cont.)",
                     "Why it explains the reported result",
                     "Why it explains the reported result (cont.)",
                     "Correction  —  and the phase it goes to",
                     "Correction (cont.)"], fill_hex=NAVY2, height=34)
    sev_fill = {"CRITICAL": "F8CBAD", "MAJOR": "FCE4D6", "INFO": "EAECEE"}
    goes = {"F-01": "Ph2 steps 2.2 → 2.6", "F-02": "Ph2 step 2.1", "F-03": "Ph3 nights 1–2",
            "F-04": "Ph4", "F-05": "Sheet 14 pre-requisite columns + sheet 17 Section 4",
            "F-06": "Every future baseline", "F-07": "Ph2 scheduling",
            "F-08": "Counter pull before Ph2 sign-off"}
    for fid, sev, title, ev, why, fix in FINDINGS:
        base = sev_fill[sev]
        fix_txt = f"{fix}\n→  Goes to:  {goes[fid]}"
        put(ws, r, 1, fid, size=10, bold=True, fill_hex=base, h="center", v="center", border=True)
        put(ws, r, 2, sev, size=9, bold=True, fill_hex=base, h="center", v="center", border=True)
        for c1, c2, txt, fh, sz, bold in (
            (3, 4, title, base, 9, True),
            (5, 7, ev, WHITE, 8, False),
            (8, 9, why, "FFF2CC", 8, False),
            (10, 11, fix_txt, PALE_GREEN, 8, False),
        ):
            merge(ws, r, c1, r, c2)
            put(ws, r, c1, txt, size=sz, bold=bold, fill_hex=fh, h="left", v="top", border=True)
            for c in range(c1 + 1, c2 + 1):
                ws.cell(r, c).border = thin
        ws.row_dimensions[r].height = _height(
            [(_w(F_WIDTHS, 3, 4), title), (_w(F_WIDTHS, 5, 7), ev),
             (_w(F_WIDTHS, 8, 9), why), (_w(F_WIDTHS, 10, 11), fix_txt)])
        r += 1

    r = blank(ws, r, 8)

    # ---------- Section 5: rules that were broken
    r = section(ws, r, COLS, "Section 5.  Execution rules that were broken  (and the fix in v7.0)")
    r = _hdr(ws, r, ["SN", "Rule in the workbook", "Rule (cont.)",
                     "What the Ph1 work order did instead", "(cont.)", "Consequence",
                     "Consequence (cont.)", "Consequence (cont.)", "Fix now in place",
                     "Fix (cont.)", "Where to find it"],
             fill_hex="C65911", height=34)
    breaks = [
        ("Send one parameter / one switch per MML line.",
         "Four switches were bundled into a single MOD NRDUCELLBEAMALGO.",
         "MOD is atomic, so one invalid option rejected all four — including the main DL lever.",
         "Sheet 17 Ph2 is written as seven separately numbered MML lines, one switch each.",
         "Sheet 17 Section 2"),
        ("Check the dependent switch before writing the work order.",
         "DlCoverageAlgoSwitch was never read, so the SUPER_COVERAGE_SW pre-requisite was missed.",
         "RETCODE 2147616329 at 2026-09-10 15:27:14 and a wasted maintenance window.",
         "Every MML row on sheet 14 now carries its pre-requisite switch and the exact pre-requisite "
         "MML to run before it.",
         "Sheet 14, the two red columns"),
        ("Do not enable SRS_IC the same night as SRS_TIGHT_MULTIPLEXING.",
         "Both were enabled in the same window and both are marked Done.",
         "The SRS estimate that DL weights and MU pairing depend on degraded, and the result cannot "
         "be attributed to either switch.",
         "SRS_IC_SW-0 is Ph2 step 2.1 and it is the ONLY change that night.",
         "Sheet 17 Ph2 step 2.1"),
        ("Verify each line with an LST and keep the RETCODE.",
         "The work order recorded Done for switches the gNodeB had refused.",
         "Four switches were believed active for two days while they were not, so the whole DL "
         "conclusion was drawn from an experiment that never ran.",
         "Every Ph2 step now has its own Verify (LST) column and must return RETCODE = 0 before the "
         "next step is sent.",
         "Sheet 17 Ph2, Verify column"),
        ("One root cause per maintenance night.",
         "Eighteen switches across five feature families went in together.",
         "No switch can be credited or blamed, so the only safe next move is to unwind one bit at a "
         "time — which costs more windows than doing it in order would have.",
         "Ph2 to Ph7 are one family per phase and one switch per night, each with its own exit gate.",
         "Sheet 17 Section 1"),
    ]
    for i, (rule, did, cons, fix, where) in enumerate(breaks, 1):
        base = alt_fill(i)
        put(ws, r, 1, i, size=9, bold=True, fill_hex=base, h="center", v="center", border=True)
        spans = ((2, 3, rule, PALE_GREEN, True), (4, 5, did, "F8CBAD", False),
                 (6, 8, cons, "FFF2CC", False), (9, 10, fix, "D6EAF8", False),
                 (11, 11, where, base, True))
        for c1, c2, txt, fh, bold in spans:
            merge(ws, r, c1, r, c2) if c2 > c1 else None
            put(ws, r, c1, txt, size=8, bold=bold, fill_hex=fh, h="left", v="top", border=True)
            for c in range(c1 + 1, c2 + 1):
                ws.cell(r, c).border = thin
        ws.row_dimensions[r].height = _height(
            [(_w(F_WIDTHS, c1, c2), txt) for c1, c2, txt, _, _ in spans])
        r += 1

    r = blank(ws, r, 8)

    # ---------- Section 6: evidence still needed
    r = section(ws, r, COLS, "Section 6.  Evidence still missing before Ph2 is signed off")
    r = bullets(ws, r, COLS, [
        "N.SRS.NI.Avg and N.UL.SRS.PreSINR.Index*, 5–12 Sep, trial and control — the only data that "
        "turns finding F-02 from a strong hypothesis into proof.",
        "DL IBLER (Σ N.DL.SCH.*.ErrTB.Ibler / Σ N.DL.SCH.*.TB) — tells whether the stricter MU gates "
        "bought link quality for the paired RB they gave up. If IBLER improved, the gates are working "
        "as designed and only the SRS input was bad.",
        "N.PRB.DL.Used.Avg and N.PRB.UL.Used.Avg — paired RB scales with offered load, so the 11 Sep "
        "drop must be shown to be a feature effect and not a traffic change.",
        "N.CCE.Used.Avg and N.CCE.DL.AggLvl* — PDCCH_AGG_LVL_COMPR_SW went in without its threshold, "
        "so its effect is currently unknown either way.",
        "The exact MAE figures behind the pairing pivot, plus DL user throughput in Mbit/s for the "
        "same five days — the numbers in Section 3 were read off a chart.",
        "The full MML execution log for 10 and 11 Sep with every RETCODE, not only the one failure — "
        "there may be further silent rejections in the same work order.",
    ], fill_hex=PALE_BLUE)
    ws.row_dimensions[r - 1].height = 118
    return ws


def build_action(wb):
    ws = wb.create_sheet(A_SHEET, 17)
    setup_sheet(ws, A_SHEET)
    set_widths(ws, A_WIDTHS)
    ws.oddHeader.left.text = "5G MIMO — Action Plan Ph1…Ph7"
    ws.oddFooter.left.text = FOOTER_LEFT

    r = banner(ws, 1, COLS, "  Action Plan  —  Ph1 complete,  Ph2 corrective,  Ph3…Ph7 staged",
               fill_hex=GREEN)
    r = note_bar(ws, r, COLS,
                 "One feature family per phase, one switch per night, one exit gate per phase. A phase "
                 "that misses its gate is rolled back that night and the next phase does not start. "
                 "Findings behind this plan are on sheet 16.")
    r = blank(ws, r, 6)

    # ---------- Section 1: phase overview
    r = section(ws, r, COLS, "Section 1.  Phase overview", fill_hex=GREEN)
    r = _hdr(ws, r, ["Phase", "Window", "Objective", "Content  (one switch per MML line)",
                     "Pre-requisite to verify BEFORE the work order is written",
                     "Pre-requisite to verify (continued)", "Exit KPI gate",
                     "Exit KPI gate (continued)", "Rollback", "Status", "Status (continued)"],
             fill_hex=GREEN, height=42)
    st_fill = {"COMPLETE": "D5F5E3", "NEXT": "FFC000", "PLANNED": "EAECEE"}
    for ph, win, obj, content, pre, gate, back, status in PHASES:
        base = st_fill[status]
        put(ws, r, 1, ph, size=12, bold=True, color=WHITE,
            fill_hex=GREEN if status == "COMPLETE" else (RED_HDR if status == "NEXT" else GRAY),
            h="center", v="center", border=True)
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
        ws.row_dimensions[r].height = _height(
            [(A_WIDTHS[2], obj), (A_WIDTHS[3], content), (_w(A_WIDTHS, 5, 6), pre),
             (_w(A_WIDTHS, 7, 8), gate), (A_WIDTHS[8], back)], cap=250)
        r += 1

    r = blank(ws, r, 8)

    # ---------- Section 2: Ph2 sequence
    r = section(ws, r, COLS,
                "Section 2.  Ph2 corrective sequence  —  the exact MML order that fixes the "
                "10-Sep execution error", fill_hex=RED_HDR)
    r = note_bar(ws, r, COLS,
                 "Step 2.1 runs ALONE on night 1 and nothing else is touched for 24 h. Steps 2.2 to 2.7 "
                 "run on night 2 in this order, one MML line each, and each line must return "
                 "RETCODE = 0 and pass its LST before the next one is sent. NrDuCellId=102 is the cell "
                 "from the failed log — repeat per cell (101 / 102 / 103).",
                 fill_hex="F8CBAD")
    r = _hdr(ws, r, ["Seq", "Night", "MML command  (one switch per line)",
                     "Pre-requisite switch / parameter  (must be ON first)",
                     "Pre-requisite MML  (run BEFORE this line)", "Verify  (LST + RETCODE)",
                     "Verify (continued)", "Counter to watch", "Counter (continued)",
                     "Gate / note", "Gate / note (continued)"], fill_hex=RED_HDR, height=42)
    for i, (seq, mml, pre, pre_mml, verify, ctr, gate) in enumerate(PH2_SEQ, 1):
        base = alt_fill(i)
        night = "Night 1\n(alone)" if seq == "2.1" else "Night 2"
        put(ws, r, 1, seq, size=11, bold=True, color=WHITE, fill_hex=RED_HDR,
            h="center", v="center", border=True)
        put(ws, r, 2, night, size=8, bold=True, fill_hex="FFC000" if seq == "2.1" else base,
            h="center", v="center", border=True)
        put(ws, r, 3, mml, size=8, bold=True, fill_hex="FFF2CC", h="left", v="top", border=True)
        put(ws, r, 4, pre, size=8, fill_hex="FCE4E4", h="left", v="top", border=True)
        put(ws, r, 5, pre_mml, size=8, bold=True, fill_hex="FCE4E4", h="left", v="top", border=True)
        merge(ws, r, 6, r, 7)
        put(ws, r, 6, verify, size=8, fill_hex=PALE_BLUE, h="left", v="top", border=True)
        ws.cell(r, 7).border = thin
        merge(ws, r, 8, r, 9)
        put(ws, r, 8, ctr, size=8, fill_hex="D5F5E3", h="left", v="top", border=True)
        ws.cell(r, 9).border = thin
        merge(ws, r, 10, r, 11)
        put(ws, r, 10, gate, size=8, fill_hex=base, h="left", v="top", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = _height(
            [(A_WIDTHS[2], mml), (A_WIDTHS[3], pre), (A_WIDTHS[4], pre_mml),
             (_w(A_WIDTHS, 6, 7), verify), (_w(A_WIDTHS, 8, 9), ctr),
             (_w(A_WIDTHS, 10, 11), gate)], cap=230)
        r += 1

    r = blank(ws, r, 6)
    r = body(ws, r, COLS,
             "Rollback for the whole of Ph2 night 2, in reverse order:  "
             "MOD NRDUCELLBEAMALGO BeamOptAlgoSwitch=BEAM_TRACKING_SW-0&INTELLIGENT_BEAM_SELECTION_SW-0;  "
             "then ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-0;  then WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-0;  "
             "then MOD NRDUCELLCHNCOVALGO DlCoverageAlgoSwitch back to its pre-change value.  "
             "Do not touch HighPrecisionBeamSwitch, MaxMimoLayerNum, or the SU / MU / AHR Phase1 + Turbo bits.",
             height=54, fill_hex=PALE_BLUE, size=9)
    r = blank(ws, r, 8)

    # ---------- Section 3: per-phase MML with pre-requisites
    r = section(ws, r, COLS,
                "Section 3.  Ph3 to Ph7  —  switch list with its pre-requisite, in send order",
                fill_hex=GREEN)
    r = note_bar(ws, r, COLS,
                 "Same discipline as Ph2: one switch per MML line, master before children, and the "
                 "pre-requisite line committed and verified first. Pre-requisite text is generated from "
                 "the same table that feeds the two new columns on sheet 14, so the two sheets cannot "
                 "drift apart.")
    r = _hdr(ws, r, ["Phase", "Seq", "Switch / parameter to enable",
                     "Pre-requisite switch / parameter  (must be ON first)",
                     "Pre-requisite MML  (run BEFORE this line)", "Night", "Counter to watch",
                     "Counter (continued)", "Note", "Note (continued)", "Note (continued)"],
             fill_hex=GREEN, height=42)

    later = [
        ("Ph3", "DL_MU_PRECISE_SCH_SW-0", "N1",
         "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB, DL IBLER",
         "Only if Ph2 missed the pairing gate. Rolling this back should return paired RB; if it does "
         "not, the cause was never the gate."),
        ("Ph3", "DL_MU_ANTI_INTRF_SCH_SW-0", "N2",
         "N.ChMeas.MIMO.DL.MuPairing.kLayer.RB, DL IBLER",
         "Only if N1 was still short of the gate. Never the same night as the row above."),
        ("Ph3", "SSB_BEAM_ADAPT_SW-1", "N3",
         "N.User.OptimalSSBBeam.Avg, HOSR, N.UECntx.AbnormRel",
         "Tilt must not be 255 on any trial cell — fix RF first or leave that cell out."),
        ("Ph3", "SSB_BEAM_VERTICAL_COV_IMP_SW-1", "N4",
         "N.User.OptimalSSBBeam.Avg, N.MAC.ThpVol.DL.OptimalSSB",
         "Child of SSB_BEAM_ADAPT_SW. Send only after N3 is green."),
        ("Ph4", "HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1", "N1",
         "N.ChMeas.MIMO.DL.Transmission.Layer.Max, paired RB per layer",
         "The master that lets the cluster spend its LAYER_16 quota. This is the phase that is "
         "supposed to raise high-layer pairing."),
        ("Ph4", "MU_RANK_BOOSTING_SW-1", "N2",
         "N.ChMeas.MIMO.DL.Pair.Layer.Avg, DL IBLER",
         "Child. One per night so each one can be credited or rolled back on its own."),
        ("Ph4", "MU_MIMO_PAIRING_PREFERRED_SW-1", "N3",
         "Paired RB per layer, DL IBLER", "Child."),
        ("Ph4", "SRS_MEAS_ACCELERATING_SW-1", "N4",
         "N.UL.SRS.PreSINR.Index*, paired RB", "Child. Watch SRS quality, not only pairing."),
        ("Ph5", "AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1", "N1",
         "N.ChMeas.CQI.SingleCW.k, DL IBLER, N.ThpVol.DL",
         "Was deliberately held out of CR01. Only after Ph4 is green."),
        ("Ph5", "PDCCH_MULTI_DIM_JOINT_SCH_SW-1", "N2",
         "N.CCE.Used.Avg, N.CCE.DL.AggLvl*", "Child of AHR Capacity Upgrade 2.0."),
        ("Ph5", "HighPrecisionBeamPhase2Sw=ON", "N3",
         "N.ThpVol.DL, DL IBLER, N.UECntx.AbnormRel",
         "iBeam 2.0 master. Its children (PRECISE_MUMIMO_EVAL_SW, FAR_UE_RANK_OPT_SW, "
         "DL_ROBUST_WEIGHT_SW) follow one per night."),
        ("Ph6", "UL_LOW_NOISE_PHASE2_SW-1", "N1",
         "N.ThpVol.UL, UL IBLER, N.UL.SRS.PreSINR.Index*",
         "Needs inter-gNB time sync — DSP CLKTST must confirm it first."),
        ("Ph6", "PUSCH_COORD_PWR_CTRL_SW-1", "N2",
         "N.ThpVol.UL, UL IBLER, N.SRS.NI.Avg", "Child of UL Boosting 2.0."),
        ("Ph6", "PMI_WEIGHT_OPT_SW-1", "N3",
         "N.ChMeas.PDSCH.MCS.k, N.ThpVol.DL",
         "Alternative weight source. Only meaningful once the SRS weight baseline from Ph2 is proven, "
         "and only as an A/B against it."),
        ("Ph7", "Green set from Ph2–Ph6, per gNB", "rolling",
         "Full pack: DL and UL user throughput, paired RB, IBLER, CCE, drop, HOSR",
         "Same MML sequence per cell. Fix Tilt=255 cells before they join. 2T2R indoor cells stay out."),
    ]
    for i, (ph, sw, night, ctr, note) in enumerate(later, 1):
        pre, pre_mml = bcs.prereq_for(sw.replace("-0", "-1"))
        base = alt_fill(i)
        put(ws, r, 1, ph, size=10, bold=True, color=WHITE, fill_hex=GREEN,
            h="center", v="center", border=True)
        put(ws, r, 2, f"{ph}.{night}", size=8, bold=True, fill_hex=base,
            h="center", v="center", border=True)
        put(ws, r, 3, sw, size=8, bold=True, fill_hex="FFF2CC", h="left", v="top", border=True)
        put(ws, r, 4, pre, size=8, fill_hex="FCE4E4", h="left", v="top", border=True)
        put(ws, r, 5, pre_mml, size=8, fill_hex="FCE4E4", h="left", v="top", border=True)
        put(ws, r, 6, night, size=8, bold=True, fill_hex=base, h="center", v="center", border=True)
        merge(ws, r, 7, r, 8)
        put(ws, r, 7, ctr, size=8, fill_hex="D5F5E3", h="left", v="top", border=True)
        ws.cell(r, 8).border = thin
        merge(ws, r, 9, r, 11)
        put(ws, r, 9, note, size=8, fill_hex=base, h="left", v="top", border=True)
        for c in (10, 11):
            ws.cell(r, c).border = thin
        ws.row_dimensions[r].height = _height(
            [(A_WIDTHS[2], sw), (A_WIDTHS[3], pre), (A_WIDTHS[4], pre_mml),
             (_w(A_WIDTHS, 7, 8), ctr), (_w(A_WIDTHS, 9, 11), note)], cap=170)
        r += 1

    r = blank(ws, r, 8)

    # ---------- Section 4: golden rules
    r = section(ws, r, COLS, "Section 4.  Execution rules for every phase from here on",
                fill_hex=RED_HDR)
    r = bullets(ws, r, COLS, GOLDEN_RULES, fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 132
    r = blank(ws, r, 8)

    r = subsection(ws, r, COLS, "Immediate next actions", fill_hex=NAVY2)
    r = _hdr(ws, r, ["#", "Action", "Action (continued)", "Action (continued)", "Why",
                     "Why (continued)", "Owner", "Before", "Before (continued)", "Blocks",
                     "Blocks (continued)"], fill_hex=NAVY2, height=26)
    todo = [
        ("Pull N.SRS.NI.Avg, N.UL.SRS.PreSINR.Index*, DL IBLER, N.CCE.Used.Avg and "
         "N.PRB.DL.Used.Avg for 5–12 Sep, trial and neighbour control.",
         "Finding F-02 is the whole basis of Ph2 step 2.1 and it is currently a hypothesis, not "
         "proof. Load data also rules out a pure traffic explanation.",
         "Performance", "Ph2 work order is written", "Ph2"),
        ("Run LST NRDUCELLCHNCOVALGO and LST NRDUCELLBEAMALGO on every trial cell "
         "(101 / 102 / 103 on each gNB).",
         "Confirms the DlCoverageAlgoSwitch value per cell and whether the 10-Sep rejection hit "
         "every cell or only NrDuCellId=102.",
         "O&M", "Ph2 night 2", "Ph2 steps 2.2–2.6"),
        ("Re-issue the Ph2 work order as seven numbered single-switch MML lines with the "
         "pre-requisite column filled in from sheet 14.",
         "The bundled command is what lost four switches in Ph1. One line per switch also makes "
         "each result attributable.",
         "RF / Optimisation", "Next maintenance window", "Ph2"),
        ("Retrieve the complete MML log with RETCODE for 10 and 11 Sep.",
         "There may be further silent rejections in the same work order that nobody has noticed "
         "yet — the one that surfaced was found by chance.",
         "O&M", "Ph2 night 1", "Confidence in the whole Ph1 baseline"),
        ("Confirm the DL user throughput figure in Mbit/s for 5, 9, 10, 11 and 12 Sep, trial and "
         "control.",
         "“No visible improvement” needs a number before and after, otherwise Ph2 has nothing to "
         "be measured against.",
         "Performance", "Ph2 night 1", "Ph2 exit gate"),
    ]
    for i, (act, why, owner, before, blocks) in enumerate(todo, 1):
        base = alt_fill(i)
        put(ws, r, 1, i, size=9, bold=True, fill_hex=base, h="center", v="center", border=True)
        merge(ws, r, 2, r, 4)
        put(ws, r, 2, act, size=8, bold=True, fill_hex="FFF2CC", h="left", v="top", border=True)
        for c in (3, 4):
            ws.cell(r, c).border = thin
        merge(ws, r, 5, r, 6)
        put(ws, r, 5, why, size=8, fill_hex=base, h="left", v="top", border=True)
        ws.cell(r, 6).border = thin
        put(ws, r, 7, owner, size=8, bold=True, fill_hex=PALE_BLUE, h="center", v="center", border=True)
        merge(ws, r, 8, r, 9)
        put(ws, r, 8, before, size=8, fill_hex=PALE_GREEN, h="left", v="center", border=True)
        ws.cell(r, 9).border = thin
        merge(ws, r, 10, r, 11)
        put(ws, r, 10, blocks, size=8, bold=True, fill_hex="F8CBAD", h="left", v="center", border=True)
        ws.cell(r, 11).border = thin
        ws.row_dimensions[r].height = _height(
            [(_w(A_WIDTHS, 2, 4), act), (_w(A_WIDTHS, 5, 6), why),
             (_w(A_WIDTHS, 8, 9), before), (_w(A_WIDTHS, 10, 11), blocks)], cap=120)
        r += 1
    return ws


def add_cross_links(wb):
    """Point sheet 14 and the cover at the two new sheets."""
    ws = wb[S14]
    r = ws.max_row + 2
    r = blank(ws, r, 10)
    r = section(ws, r, bcs.COLS, "Trial result of this proposal  —  Ph1 executed 10–11 Sep 2026")
    r = note_bar(ws, r, bcs.COLS,
                 "The Section 2 list below was executed on the trial cluster. 18 switches were "
                 "accepted; 4 were rejected with RETCODE 2147616329 because "
                 "DlCoverageAlgoSwitch=SUPER_COVERAGE_SW-1 was missing. DL user throughput stayed flat "
                 "and DL MU paired RB fell. Findings and the corrective sequence are on the two sheets "
                 "linked here.", fill_hex="F8CBAD")
    for label, target in ((f"→  {F_SHEET}   (what happened and why)", F_SHEET),
                          (f"→  {A_SHEET}   (what to do next, in MML order)", A_SHEET)):
        merge(ws, r, 1, r, bcs.COLS)
        href_sheet(ws.cell(r, 1), target, label)
        for c in range(1, bcs.COLS + 1):
            ws.cell(r, c).fill = fill(PALE_GOLD)
            ws.cell(r, c).border = thin
        ws.row_dimensions[r].height = 22
        r += 1


def extra(wb):
    build_findings(wb)
    build_action(wb)
    add_cross_links(wb)


def main():
    bcs.main(out=OUT, zip_out=ZIP_OUT, with_prereq=True, extra=extra, version="v7.0")


if __name__ == "__main__":
    main()
