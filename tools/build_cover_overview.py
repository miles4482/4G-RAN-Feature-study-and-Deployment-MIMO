# Cover, Overview, Types — combined MIMO family (same document-style as CA workbook).
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from mimo_excel_style import *
from mimo_common import COLS, fig, STEP_WIDTHS

COVER_WIDTHS = [18, 22, 22, 22, 22, 22, 18, 18, 18, 18]


def build_cover(wb):
    ws = wb.create_sheet("0. Cover & Index", 0)
    setup_sheet(ws, "Cover")
    set_widths(ws, COVER_WIDTHS)
    r = 1
    r = banner(ws, r, COLS, "  5G MIMO (all features together)  —  Deployment Workbook", size=22, height=36)
    r = body(ws, r, COLS,
             "Prepared from the eight Huawei 5G RAN10.1 Feature Parameter Descriptions in this repository. "
             "Layout matches the earlier Carrier Aggregation workbook: document-style (gridlines off), "
             "one combined Excel, numbered live-network sequence, and every MML row uses a Seq column. "
             "Values marked as document examples are taken verbatim from FPD MML samples (NrDuCellId=0/1/90). "
             "Where the FPD does not publish a factory default, Default is “See Parameter Reference”. "
             "Replace sample IDs with live-network NR DU cell / TRP IDs before execution.",
             fill_hex=PALE_GOLD)
    r = blank(ws, r)
    r = section(ws, r, COLS, "Document identity")
    r = headers(ws, r, ["Item", "Value"] + [""] * 8)
    for a, b in [
        ("Source documents (8 FPDs)", "MIMO (TDD); Beam Management; Massive MIMO AHR; iBeam; Uplink Boosting; Distributed Antenna Solutions; Fusion Cell; mmWave Beam Management"),
        ("Product / version", "Huawei 5G RAN 10.1  (Draft B 2026-01-30 for MIMO/Beam/AHR/iBeam/UL Boosting; Draft A 2025-12-31 for DAS/Fusion/mmWave)"),
        ("Applicable RAT", "NR TDD  —  FR1 low-frequency (410–7125 MHz) and FR2 mmWave (24.25–71 GHz)"),
        ("Huawei feature family", "FBFD-010003 / FOFD-010020 / FOFD-010010 / FBFD-010025 / FBFD-010015 / FOFD-010100 / FOFD-051301 / FOFD-061201 / FOFD-061202 / FOFD-081201 / FOFD-091200 / FOFD-100200 / FOFD-091201 / FOFD-100201 / FOFD-050202 / FOFD-071211 / FBFD-091101 / FOFD-030201"),
        ("Purpose of this workbook", "Deployment-oriented extract: principles, prerequisites, parameters, sequenced MML, counters, KPI, licenses — all MIMO together"),
        ("How to use", "Read sheets 1–2 for concepts. Deploy Step1 → Step11 in order. Do not skip: later Massive MIMO features list earlier MIMO / MU-MIMO / Beam Management as prerequisites. Sheet 13 is Inter-Cell Cable Sequence Detection (commissioning)."),
    ]:
        r = table_row(ws, r, [a, b] + [""] * 8, fills=[PALE_BLUE, WHITE] + [WHITE] * 8, bolds=[True] + [False] * 9, height=36)
        merge(ws, r - 1, 2, r - 1, COLS)
    r = blank(ws, r)
    r = section(ws, r, COLS, "Sheet map  (document sequence — MUST follow this live-network sequence)")
    r = headers(ws, r, ["#", "Sheet", "Maps to document", "What you will find"] + [""] * 6)
    rows = [
        ("0", "0. Cover & Index", "—", "Identity, how to use, mandatory sequence"),
        ("1", "1. Overview of MIMO", "MIMO TDD Ch.3 + family overviews", "Why MIMO, antenna modes, SU/MU/Massive/mmWave split"),
        ("2", "2. Type of MIMO Config", "All 8 FPDs §2.2 feature lists", "Feature IDs, FR1 vs FR2, NSA/SA, which step owns which switch"),
        ("3", "3. Step1 Basic MIMO", "MIMO TDD Ch.4  FBFD-010003", "Rx diversity, DL beamforming, weight algorithms, MML Seq 1.x"),
        ("4", "4. Step2 SU-MIMO", "MIMO TDD Ch.5  FOFD-010020", "Single-UE layers, MaxMimoLayer*, rank adapt, licenses NR0S0DLEPU00 / NR0S0ULEPU00"),
        ("5", "5. Step3 MU-MIMO", "MIMO TDD Ch.6  FOFD-010010", "PDSCH/PUSCH pairing, isolation thresholds, MuMimoSwitch"),
        ("6", "6. Step4 MM Multi-Layer", "MIMO TDD Ch.7", "High-layer MU, UL multilayer demod, after Step3 is green"),
        ("7", "7. Step5 Beam Management", "Beam Mgmt Ch.4–5  FBFD-010015 / FOFD-010100", "Broadcast/control beams then 3D coverage pattern"),
        ("8", "8. Step6 AHR", "AHR FPD Ch.4–6  FOFD-051301 / 061201 / 061202", "AHR → Experience Turbo 2.0 → Capacity Upgrade 2.0"),
        ("9", "9. Step7 iBeam", "iBeam FPD Ch.3–5  FOFD-081201 / 091200 / 100200", "iBeam → 2.0 → 3.0 (HighPrecisionBeam* switches)"),
        ("10", "10. Step8 UL Boosting", "UL Boosting FPD Ch.4–5  FOFD-091201 / 100201", "UL_LOW_NOISE_SW then PHASE2"),
        ("11", "11. Step9 DAS + Fusion", "DAS FPD + Fusion Cell FPD", "Distributed Massive MIMO, then Fusion Cell / Virtual 128T"),
        ("12", "12. Step10 mmWave", "mmWave Beam + MIMO TDD Ch.8–9", "FR2 beam management, mmWave MU-MIMO, multi-beam FDM"),
        ("13", "13. Step11 Cable Sequence", "MIMO TDD Ch.10  FBFD-010025", "Inter-cell cable sequence detection (STR ANTENNAPORTOPTDET)"),
    ]
    for i, rec in enumerate(rows):
        vals = list(rec) + [""] * 6
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=30)
        merge(ws, r - 1, 4, r - 1, COLS)
    r = blank(ws, r)
    r = section(ws, r, COLS, "Mandatory live-network sequence  (do not skip or reverse FR1 steps)")
    r = bullets(ws, r, COLS, [
        "Seq 0  Commissioning: confirm TDD NR cell is 2T2R/4T4R/8T8R/32T32R/64T64R as planned; polarization of physical vs logical antenna ports matches (MIMO TDD Ch.4.3.4). Run Step11 cable-sequence detection on 4T4R NORMAL_CELL sites before blaming MIMO KPI.",
        "Seq 1  Basic MIMO (FBFD-010003) — NR cells are multi-antenna by default. Enable weight/optimization bits only after MetaAAU beam sensing if those functions are used.",
        "Seq 2  SU-MIMO Multiple Layers (FOFD-010020) — purchase SU-MIMO license + layer-capacity licenses (one unit = 2 layers). Set MaxMimoLayerNum / MaxMimoLayerCnt.",
        "Seq 3  MU-MIMO Basic Pairing (FOFD-010010) — requires SU-MIMO. Turn MuMimoSwitch UL+DL. Verify N.ChMeas.MIMO.*.Layer.Max.",
        "Seq 4  Massive MIMO multi-layer enhancement — only after MU-MIMO is stable on 32T/64T. HighLayerMuMimoSw / UlHighLayerMuMimoSwitch.",
        "Seq 5  Beam Management: Basic (FBFD-010015) then 3D Coverage Pattern (FOFD-010100, NR0SBSC3DC00). Coverage scenario / tilt / azimuth on NRDUCELLTRPBEAM.",
        "Seq 6  AHR Introduction (AHR_PHASE1_SW) → Experience Turbo 2.0 → Capacity Upgrade 2.0. Do not jump to Capacity Upgrade without Phase1.",
        "Seq 7  iBeam → iBeam 2.0 → iBeam 3.0 (HighPrecisionBeamSwitch then Phase2Sw then Phase3Sw).",
        "Seq 8  Uplink Boosting (UL_LOW_NOISE_SW) then Boosting 2.0 (UL_LOW_NOISE_PHASE2_SW).",
        "Seq 9  Site architecture last on FR1: Distributed Massive MIMO (DAS) if multi-TRP indoor/macro split; Fusion Cell / Virtual 128T if two 64T TRPs in one cell.",
        "Seq 10 FR2 mmWave is a parallel track, not a substitute for FR1 Seq 1–9. Basic mmWave beams first, then 3D coverage / dense beam / MU-MIMO / multi-beam FDM.",
        "Huawei note: these FPDs are activation guidance. Feature gains depend on the live scenario; contact Huawei service for optimization. Check “Service Interrupted After Modification” in the Parameter Reference before each MOD.",
    ], fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 175
    return ws


def build_overview(wb):
    ws = wb.create_sheet("1. Overview of MIMO")
    setup_sheet(ws, "Overview")
    set_widths(ws, [16, 22, 18, 18, 16, 16, 16, 16, 16, 18])
    r = 1
    r = banner(ws, r, COLS, "  1.  Overview of MIMO  (all features together)")
    r = note_bar(ws, r, COLS, "Primary source: MIMO (TDD) FPD Ch.3 Overview. Family overviews from Beam / AHR / iBeam / UL Boosting / DAS / Fusion / mmWave FPDs.")

    r = section(ws, r, COLS, "1.1  Introduction — why MIMO exists")
    r = body(ws, r, COLS,
             "Extending system bandwidth raises capacity but not spectral efficiency. Raising modulation order helps only while signal quality allows it. "
             "MIMO uses multiple antennas at transmitter and receiver to multiply spectral efficiency. "
             "NR TDD supports 2T2R / 4T4R / 8T8R / 32T32R / 64T64R. Capability scales with antenna count, so 32T32R/64T64R (massive MIMO cells) is recommended for NR TDD. "
             "To use MIMO fully, Huawei combines receive diversity, downlink beamforming, spatial multiplexing (SU-MIMO and MU-MIMO), and (FR2) multi-beam FDM.")
    r = insert_figure(ws, r, fig("mimo_tdd", "p19_0.png"), COLS,
                      "Figure 4-1  Principles of uplink receive diversity  (MIMO TDD FPD Ch.4.1.1)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p20_0.png"), COLS,
                      "Downlink beamforming weight application  (MIMO TDD FPD Ch.4.1.2)")

    r = section(ws, r, COLS, "1.2  How MIMO works (building blocks)")
    r = bullets(ws, r, COLS, [
        "Uplink receive diversity: UE sends x; gNodeB applies weight vector W to M receive antennas and combines y = W(Hx + N). Improves uplink reception (Ch.4.1.1).",
        "Downlink beamforming: gNodeB forms a directional beam toward the UE using PMI-based and/or SRS-based weights (FR1). FR2 uses PMI-based weights only — no SRS-based weights (Ch.2.3 high/low frequency difference).",
        "SU-MIMO spatial multiplexing: multiple layers of one UE. Peak rate theoretically N times single-layer if the UE supports N layers (Ch.5).",
        "MU-MIMO spatial multiplexing: multiple UEs scheduled on the same time-frequency resource with spatial isolation (Ch.6 FR1, Ch.8 FR2).",
        "Massive MIMO multi-layer enhancement: more paired MU layers need stronger interference suppression, scheduling, and weight accuracy (Ch.7, FR1 only).",
        "Multi-beam FDM (FR2): frequency-division of beams to raise spectral efficiency (Ch.9).",
        "Antenna-port matching: physical vs logical polarization of RRU/AAU ports must match or MIMO collapses (Ch.4.3.4). Inter-cell cable sequence detection finds crossed feeders (Ch.10).",
    ])
    ws.row_dimensions[r - 1].height = 130

    r = section(ws, r, COLS, "1.3  Classification")
    r = headers(ws, r, ["Dimension", "Type", "Meaning"] + [""] * 7)
    classif = [
        ("Antenna scale", "Conventional MIMO (2/4/8T)", "Basic Rx diversity + beamforming + limited SU layers."),
        ("Antenna scale", "Massive MIMO (32T/64T)", "Recommended NR TDD. MU pairing, AHR, iBeam, UL boosting apply here."),
        ("User multiplexing", "SU-MIMO", "Layers of one UE. License FOFD-010020 + 2-layer capacity units."),
        ("User multiplexing", "MU-MIMO", "Layers of several UEs on the same PRB. License FOFD-010010."),
        ("Frequency", "FR1 low-frequency TDD", "PMI + SRS weights; AHR/iBeam/UL boosting/DAS/Fusion; cable sequence."),
        ("Frequency", "FR2 mmWave TDD", "PMI weights only; mmWave beam management, 3D coverage, MU-MIMO, multi-beam FDM."),
        ("Site architecture", "Single-TRP cell", "Default. One NRDUCELLTRP per NR DU cell."),
        ("Site architecture", "Distributed Massive MIMO", "Master+slave TRPs (DAS FPD). Indoor LampSite FOFD-050202 or macro FOFD-071211."),
        ("Site architecture", "Fusion Cell / Virtual 128T", "Two 64T TRPs in one cell (FBFD-091101). Cluster INTRA_CELL_MIMO."),
        ("Beam family", "Broadcast / control / traffic beams", "SSB/CSI-RS broadcast, PDCCH control beams, PDSCH traffic weights (Beam Management FPD)."),
    ]
    for i, rec in enumerate(classif):
        vals = list(rec) + [""] * 7
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=30)
        merge(ws, r - 1, 3, r - 1, COLS)

    r = blank(ws, r)
    r = section(ws, r, COLS, "1.4  FR1 vs FR2  (MIMO TDD Ch.2.3)")
    r = headers(ws, r, ["Function", "Low frequency (FR1)", "High frequency (FR2)", "Workbook step"] + [""] * 6)
    fr = [
        ("Basic MIMO weights", "PMI-based and SRS-based", "PMI-based only", "Step1 / Step10"),
        ("PUSCH beam/time-domain enh.", "Supported", "Not supported", "Step1"),
        ("SU-MIMO UL UE SINR opt.", "Not this FPD’s FR1 item", "Supported only in high frequency", "Step2 / Step10"),
        ("Massive MIMO multi-layer", "Supported", "Not supported", "Step4"),
        ("Multi-beam FDM", "Not supported", "Supported", "Step10"),
        ("Inter-cell cable sequence", "Supported (4T4R)", "Not supported", "Step11"),
        ("AHR / iBeam / UL Boosting", "FR1 Massive MIMO", "Out of those FPDs", "Step6–8"),
        ("3D coverage pattern", "FOFD-010100 NR0SBSC3DC00", "FOFD-030201 NR0SMMW3DCP0", "Step5 / Step10"),
    ]
    for i, rec in enumerate(fr):
        vals = list(rec) + [""] * 6
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=26)
        merge(ws, r - 1, 4, r - 1, COLS)

    r = blank(ws, r)
    r = section(ws, r, COLS, "1.5  Enhancement families (what each later FPD adds)")
    r = bullets(ws, r, COLS, [
        "Beam Management: broadcast-beam coverage (32T/64T vs 2/4/8T), control-beam robustness, then 3D coverage scenarios (including MetaAAU and customized beams).",
        "AHR: adaptive CSI-RS multi-beam + adaptive multi-stream perception (Phase1); Experience Turbo 2.0 (SRS interference coordination + precise AMC); Capacity Upgrade 2.0 (MU pairing / gathering).",
        "iBeam: high-precision beam weights and interference-aware MU scheduling, then 2.0 robust weights / dynamic cluster, then 3.0 self-fusion weights and smart AMC.",
        "Uplink Boosting: reduce PUSCH/PUCCH/SRS interference (Phase1), then 2.0 coordinated power control + precise RB/MCS + multi-beam reception.",
        "DAS: one logical cell, multiple TRPs — coverage fill and UL boost with DmMimoSwitch.",
        "Fusion Cell: virtual 128T by fusing two 64T TRPs; MU-MIMO SINR enhancement and single-TRP SSB options.",
    ])
    ws.row_dimensions[r - 1].height = 110
    r = insert_figure(ws, r, fig("mimo_tdd", "p50_0.png"), COLS, "SU-MIMO basics  (MIMO TDD FPD Ch.5.1.1)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p66_0.png"), COLS, "MU-MIMO basics  (MIMO TDD FPD Ch.6.1.1)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p69_0.png"), COLS, "MU-MIMO spatial multiplexing procedure  (MIMO TDD FPD Ch.6.1.2)")
    return ws


def build_types(wb):
    ws = wb.create_sheet("2. Type of MIMO Config")
    setup_sheet(ws, "Types")
    set_widths(ws, STEP_WIDTHS)
    r = 1
    r = banner(ws, r, COLS, "  2.  Type of MIMO configuration  —  feature IDs, switches, which step")
    r = note_bar(ws, r, COLS, "Each row is one Huawei feature from the eight FPDs §2.2. Deploy in Seq order. Gold rows are capacity licenses that ride on SU/MU layer count.")

    r = section(ws, r, COLS, "2.1  Feature catalogue  (one feature per row)")
    r = headers(ws, r, ["Seq", "Feature ID", "Feature name", "FPD / chapter", "Master switch / MO", "License model", "Workbook step"] + [""] * 3)
    feats = [
        ("1", "FBFD-010003", "MIMO Basic Package", "MIMO TDD Ch.4", "NR cells multi-antenna by default; WeightAlgoSwitch / AdaptiveEdgeExpEnhSwitch", "None (basic functions)", "3. Step1"),
        ("2", "FOFD-010020", "SU-MIMO Multiple Layers", "MIMO TDD Ch.5", "NRDUCellPdsch.MaxMimoLayerNum; NRDUCellPusch.MaxMimoLayerCnt", "NR0S0PREUM00 per Cell + NR0S0DLEPU00 / NR0S0ULEPU00 per 2 layers", "4. Step2"),
        ("3", "FOFD-010010", "MU-MIMO Basic Pairing", "MIMO TDD Ch.6 & Ch.8", "NRDUCellAlgoSwitch.MuMimoSwitch = UL_MU_MIMO_SW & DL_MU_MIMO_SW", "NR0S00MUMM00 per Cell + layer capacity", "5. Step3 / 12. Step10"),
        ("4", "— (Ch.7)", "Massive MIMO Multi-Layer Enhancement", "MIMO TDD Ch.7", "NRDUCellDlMimo.HighLayerMuMimoSw; NRDUCellUlMimo.UlHighLayerMuMimoSwitch", "Layer capacity NR0S0DLEPU00 / NR0S0ULEPU00", "6. Step4"),
        ("5", "FBFD-010015", "Basic Beam Management", "Beam Mgmt Ch.4; mmWave Ch.4", "NRDUCellTrpBeam.CoverageScenario; BeamOptSwitch; PdcchAlgoSwitch", "None listed for basic LF beams", "7. Step5 / 12. Step10"),
        ("6", "FOFD-010100", "3D Coverage Pattern (LF)", "Beam Mgmt Ch.5", "NRDUCellTrpBeam.CoverageScenario = SCENARIO_n / CUSTOMIZED_BEAM_SCENARIO", "NR0SBSC3DC00 per Cell (not SCENARIO_31/37)", "7. Step5"),
        ("7", "FOFD-051301", "AHR Introduction", "AHR Ch.4", "NRDUCellFeatureSw.AhrSwitch = AHR_PHASE1_SW", "NR0S00AET100 per Cell", "8. Step6"),
        ("8", "FOFD-061201", "AHR Experience Turbo 2.0", "AHR Ch.5", "AhrSwitch = AHR_EXP_TURBO_PHASE2_SW", "NR0S00AET200 per Cell", "8. Step6"),
        ("9", "FOFD-061202", "AHR Capacity Upgrade 2.0", "AHR Ch.6", "AhrSwitch = AHR_CAPC_UPGRADE_PHASE2_SW", "NR0S00ACT200 per Cell", "8. Step6"),
        ("10", "FOFD-081201", "iBeam", "iBeam Ch.3", "NRDUCellFeatureSw.HighPrecisionBeamSwitch = ON", "NR0S00BEAM00 per Cell", "9. Step7"),
        ("11", "FOFD-091200", "iBeam 2.0", "iBeam Ch.4", "HighPrecisionBeamPhase2Sw = ON", "NR0S0DLEHR00 per Cell", "9. Step7"),
        ("12", "FOFD-100200", "iBeam 3.0", "iBeam Ch.5", "HighPrecisionBeamPhase3Sw = ON", "NR0S00BEAM30 per Cell", "9. Step7"),
        ("13", "FOFD-091201", "Uplink Boosting", "UL Boosting Ch.4", "NRDUCellFeatureSw.MimoFeatureSwitch = UL_LOW_NOISE_SW", "NR0S00UAHR00 per Cell", "10. Step8"),
        ("14", "FOFD-100201", "Uplink Boosting 2.0", "UL Boosting Ch.5", "MimoFeatureSwitch = UL_LOW_NOISE_PHASE2_SW", "NR0S00UPBT20 per Cell", "10. Step8"),
        ("15", "FOFD-050202", "Distributed Massive MIMO (LampSite)", "DAS Ch.4", "NRDUCellAlgoSwitch.DmMimoSwitch = DM_MIMO_SERVICE_SWITCH", "NR0SDMMIMO00", "11. Step9"),
        ("16", "FOFD-071211", "Distributed Massive MIMO (NR TDD macro)", "DAS Ch.4", "Same DmMimoSwitch on macro", "NR0S00VMMM00", "11. Step9"),
        ("17", "FBFD-091101", "Virtual 128T / Fusion Cell", "Fusion Cell Ch.4", "ADD GNBCLUSTER ClusterType=INTRA_CELL_MIMO; FusionCellAlgoSwitch", "Capacity by TRP count (see License Management)", "11. Step9"),
        ("18", "FOFD-030201", "mmWave 3D Coverage Pattern (+ TA/dense/dynamic beam)", "mmWave Ch.5,7,8,9", "NRDUCELLTRPMMWAVBEAM.CoverageScenario; FLEXIBLE_DENSE_BEAM_SW; DYNAMIC_BEAM_ALLOC_SW", "NR0SMMW3DCP0 per Cell", "12. Step10"),
        ("19", "FOFD-010010", "mmWave MU-MIMO", "MIMO TDD Ch.8", "MuMimoSwitch=UL_MU_MIMO_SW (FR2 example)", "NR0S00MUMM00 + NR0SMMWULE00", "12. Step10"),
        ("20", "— (Ch.9)", "Multi-Beam FDM (FR2)", "MIMO TDD Ch.9", "NRDUCellAlgoSwitch.BeamMultiplexSwitch = VOL_BASED_BEAM_MULTIPLEX_SW", "See Parameter Reference", "12. Step10"),
        ("21", "FBFD-010025", "Basic O&M — Inter-Cell Cable Sequence Detection", "MIMO TDD Ch.10", "STR ANTENNAPORTOPTDET: AntPortOptDetPolicy=INTER_CELL_DETECT", "None", "13. Step11"),
    ]
    for i, rec in enumerate(feats):
        vals = list(rec) + [""] * 3
        r = table_row(ws, r, vals, fills=[alt_fill(i)] * 10, height=40)
        merge(ws, r - 1, 7, r - 1, COLS)

    r = blank(ws, r)
    r = section(ws, r, COLS, "2.2  Configuration mode vs CA workbook analogue")
    r = body(ws, r, COLS,
             "Carrier Aggregation used group-based vs adaptive mode. MIMO does not use that split. "
             "Instead the “mode” column in MML sheets is the feature phase (Basic / SU / MU / AHR Phase1 / iBeam 2.0 / …) "
             "and the Seq column is the order printed in the FPD (activation first, then optional optimization, then deactivation in reverse). "
             "Always finish the previous Seq sheet’s activation verification before turning the next master switch.")

    r = section(ws, r, COLS, "2.3  NSA vs SA")
    r = body(ws, r, COLS,
             "MIMO TDD Ch.2.3: NSA vs SA difference is None for basic MIMO, SU-MIMO, MU-MIMO, multi-layer enhancement, multi-beam FDM, and cable sequence detection. "
             "Verification tip (SU-MIMO Ch.5.4.2): on MAE User Common Monitoring set UEID Type to Random Value or STMSI in NSA, and 5G-Random Value or 5G-STMSI in SA.")
    return ws
