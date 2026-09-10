#!/usr/bin/env python3
"""Append MIMO dump-vs-commercial inconsistency as the LAST sheet.

Copies MIMO_Deployment.xlsx → MIMO_Deployment_v2.0.xlsx (does not overwrite v1).
Source dump: 5G CME 8 Sep 2026 DHK from 5G-RAN-BASIC-TO-ADVANCE_OPTIMIZATION
(Configuration-Dump). 4G dump is not required for NR MIMO air-interface switches.
"""
import os, shutil, zipfile, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from copy import copy
from openpyxl import load_workbook
from mimo_excel_style import *

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "MIMO_Deployment.xlsx")
OUT = os.path.join(ROOT, "MIMO_Deployment_v2.0.xlsx")
ZIP_OUT = os.path.join(ROOT, "MIMO_Deployment_v2.0.zip")

INCON_COLS = 11
INCON_WIDTHS = [6, 10, 18, 10, 58, 24, 32, 42, 32, 48, 36]
SHEET_NAME = "14. MIMO Incon Report"

BLUE_HDR = "5B9BD5"
YELLOW_HDR = "FFC000"
CRIT = "C00000"
MAJOR = "C65911"
MINOR = "BF8F00"
OKC = "548235"
INFOC = "7030A0"
HOLD = "806000"

SEV_FILL = {
    "Critical": "F4B183",
    "Major": "F8CBAD",
    "Minor": "FFF2CC",
    "Introduce": "DDEBF7",
    "OK": "C6EFCE",
    "Info": "E2D5F1",
    "Skip": "E2EFDA",
    "Hold": "FCE4D6",
    "Enable": "F8CBAD",
    "Check": "DDEBF7",
    "Fix": "F4B183",
    "Seq": "1F4E79",
}


def sev_color(sev):
    return SEV_FILL.get(sev, WHITE)


def put11(ws, r, values, fills=None, bolds=None, height=None, center=None, font_color=None):
    for i, v in enumerate(values, 1):
        fh = fills[i - 1] if fills and i - 1 < len(fills) else None
        b = bolds[i - 1] if bolds and i - 1 < len(bolds) else False
        h = "center" if center and i in center else "left"
        col = font_color if font_color else "000000"
        put(ws, r, i, v, size=9, bold=b, color=col, fill_hex=fh, h=h, v="top", border=True)
    if height is None:
        longest = max((len(str(v)) if v is not None else 0) for v in values)
        height = min(72, max(20, 16 + longest // 90 * 10))
    ws.row_dimensions[r].height = height
    return r + 1


def findings_header(ws, r):
    titles = ["ID", "Severity", "Domain", "MO", "Parameter / Switch",
              "Live DHK dump (32T unless noted)", "Commercial good-DL-tput target",
              "Gap / inconsistency", "KPI impact", "Suggested action", "License"]
    fills = [NAVY] * 11
    for i, t in enumerate(titles, 1):
        put(ws, r, i, t, size=9, bold=True, color=WHITE, fill_hex=fills[i - 1],
            h="center", v="center", border=True)
    ws.row_dimensions[r].height = 28
    return r + 1


def mml_header(ws, r):
    titles = ["SN", "RAT", "Doc Name", "Action", "MML Command (Proposed)",
              "MO Name", "Parameter ID", "Live Value (actual in Dump)",
              "proposed Parameter Value", "Purpose/Short Notes", "License"]
    for i, t in enumerate(titles, 1):
        fh = BLUE_HDR if i <= 4 else YELLOW_HDR
        put(ws, r, i, t, size=9, bold=True, color="000000", fill_hex=fh,
            h="center", v="center", border=True)
    ws.row_dimensions[r].height = 28
    return r + 1


def find_row(ws, r, rec):
    sev = rec[1]
    fh = sev_color(sev)
    fills = [fh] * 11
    bolds = [True, True] + [False] * 9
    return put11(ws, r, rec, fills=fills, bolds=bolds, center={1, 2}, height=48)


def mml_row(ws, r, rec):
    action = rec[3]
    fh = sev_color(action)
    if action == "Seq":
        fills = [NAVY] * 11
        return put11(ws, r, rec, fills=fills, bolds=[True] * 11, font_color=WHITE, height=26)
    fills = [PALE_BLUE] * 4 + [fh] + [PALE_GOLD] * 6
    fills[3] = fh
    return put11(ws, r, rec, fills=fills, bolds=[False, False, False, True, True] + [False] * 6,
                 center={1, 2, 4}, height=40)


# ---------------------------------------------------------------------------
# Findings: dump vs commercial 32T NSA with good DL user throughput
# ---------------------------------------------------------------------------
FINDINGS = [
    ("MIMO-001", "Critical", "Beam RF (install)", "NRDUCellTrpBeam",
     "Tilt(degree)",
     "Mode 7° (150); 6° 106; 4° 88; 9° 81; 8° 71; 5° 30; 255 on 28 cells; 10–12° few.",
     "Per-sector RF design. 255 is illegal/default — beamforming not applied. Urban 32T typically 4–9° after drive/cluster.",
     "28 cells with tilt=255 → SSB/CSI/PDSCH beams do not follow planned downtilt. Overshoot + interference on neighbours.",
     "DL user tput, interference, HOSR, edge MCS.",
     "Fix tilt=255 immediately (site design; fallback 7°). Then retune 4° cells if overshoot. Do this before any MIMO switch trial.",
     "No extra license (RF)"),
    ("MIMO-002", "Critical", "Beam RF (install)", "NRDUCellTrpBeam",
     "Azimuth(degree)",
     "0° on 580/581; 10° only on DHTIA82-101.",
     "Electrical/beam azimuth follows sector plan (0 often = mechanical-only if AAU compass is set). Cluster-wide 0° must be audited vs RF design.",
     "If template zeroed planned azimuths, all digital beams point to mechanical 0. Only 1 cell has a non-zero electrical azimuth.",
     "Coverage, overlap, HO, DL user tput.",
     "Audit vs RF design / AAU compass. Correct electrical azimuth per sector before beam-adapt trials.",
     "No extra license (RF)"),
    ("MIMO-003", "OK", "Hardware", "NRDUCellTrp",
     "TxRxMode",
     "32T32R 564; 2T2R 17 (DHTIAA1×10, DHAPT11×5, DHTEJ34×2).",
     "Outdoor AAU = 32T32R. Keep 2T2R only if indoor/special RRU.",
     "17 TRPs are not massive MIMO. MU/AHR already off there (matches hardware).",
     "Capacity on those 17 cells only.",
     "Do not apply 32T MIMO/iBeam/AHR CU trials on DHTIAA1 / DHAPT11 / DHTEJ34. Keep 2T2R MU OFF.",
     "—"),
    ("MIMO-004", "OK", "SU-MIMO (already ON)", "NRDUCellAlgoSwitch",
     "SuMimoSwitch DL_SU_MULTI_LAYER / UL_SU_MULTI_LAYER / DL_SU_PWR_CTRL_ENH",
     "DL+UL SU multi-layer ON 564/564. DL_SU_PWR_CTRL_ENH ON 563/564.",
     "ON for 32T SU peak.",
     "Aligned with commercial SU baseline. Do not re-send.",
     "SU peak layers already available.",
     "Skip. Fix the 1 leftover 32T if SU pwr-ctrl enhance is OFF.",
     "FOFD-010020 already ON"),
    ("MIMO-005", "OK", "MU-MIMO (already ON)", "NRDUCellAlgoSwitch",
     "MuMimoSwitch DL_MU_MIMO / UL_MU_MIMO / PDCCH_MU",
     "DL+UL MU ON 563/564; PDCCH_MU ON 563/564. Group mode ISOLATION_CORRELATION. PreSINR −50. BackToSu 5.",
     "DL+UL MU ON for 32T. Isolation + PreSINR −50 match MIMO TDD activation example.",
     "Core MU is live. 1 extra 32T may still be OFF. IR/optimal-layer still OFF (next wave).",
     "Cell capacity already using basic MU.",
     "Skip re-enable. Confirm leftover 32T MU ON. Keep 2T2R OFF.",
     "FOFD-010010 already ON"),
    ("MIMO-006", "OK", "AHR (already ON)", "NRDUCellFeatureSw / Csirs",
     "AHR_PHASE1_SW + AHR_EXP_TURBO_PHASE2_SW + FD_RESOURCE + TYPE0 + 8PORT + CSIRS_INTRF_STATIC_AVOID",
     "Phase1 + Turbo 2.0 ON 563/564. CSI FD_RESOURCE / TYPE0 / 8PORT / static avoid ON.",
     "Commercial good-DL-tput 32T: AHR Phase1 + Turbo on. CSI baseline FD+TYPE0+8PORT.",
     "AHR introduction and Turbo master are already live. Capacity Upgrade 2.0 is the missing AHR wave.",
     "DL experience / CSI beam adapt already on.",
     "Keep ON 32T. Do not disable Turbo. Next AHR wave is Capacity Upgrade after SRS-IC trial.",
     "FOFD-051301 / 061201 already ON"),
    ("MIMO-007", "OK", "Beam baseline (already ON)", "NRDUCellBeamAlgo / AlgoSwitch / TrpBeam",
     "BeamPerceiveMode / DL_INITIAL_BEAM_SELECT / SSB_BEAM_DENSITY_ADAPT / DL_PMI_SRS_ADAPT",
     "DISTRIBUTED_MODE 564/564. DL_INITIAL_BEAM_SELECT ON 564. SSB density adapt ON. DL_PMI_SRS_ADAPT ON 564.",
     "DISTRIBUTED_MODE + initial DL beam select + density adapt + PMI/SRS adapt = commercial AAU baseline.",
     "Perceive/initial-select/density already match commercial. Tracking / SSB adapt / weight estimate still OFF.",
     "Idle SSB coverage OK; connected BF tracking missing.",
     "Keep DISTRIBUTED_MODE. Do not re-send these four.",
     "FBFD-010003 / 010015 already ON"),
    ("MIMO-008", "OK", "Layer quota", "NRDUCellPdsch",
     "MaxMimoLayerNum",
     "LAYER_16 on 557/564 32T; LAYER_DEFAULT on 7.",
     "Keep live LAYER_16. FPD example LAYER_8 is a lab floor — do not downgrade.",
     "Quota is already better than the FPD sample. 7 cells on DEFAULT should be raised to LAYER_16 if they are 32T.",
     "Peak layer cap already 16.",
     "Never send LAYER_8. Align 7×LAYER_DEFAULT 32T to LAYER_16.",
     "FOFD-010020 layer licenses"),
    ("MIMO-009", "OK", "DL scheduling core", "NRDUCellDlSch / Pdsch",
     "DlSchPolicy=EPF; HEAVY_LOAD_SCH_PRI_OPT; TAIL_PKT_SCH_OPT; DL rank adapt",
     "EPF 564. HEAVY_LOAD_SCH_PRI_OPT ON 563. TAIL_PKT_SCH_OPT ON 563. DL_RANK_ADAPT ON 564.",
     "EPF + heavy-load pri + tail-pkt sch + rank adapt = commercial scheduler baseline.",
     "Core scheduler healthy. Missing are MCS-opt / freq-sel / res-based / iBeam interference-aware bits.",
     "Fairness under load already on.",
     "Skip these. Enable the missing sch bits in W1/W3 (see MIMO-020+).",
     "—"),
    ("MIMO-010", "Major", "Beam weights (MISSING)", "NRDUCellBeamAlgo",
     "WeightAlgoSwitch: SRS_WEIGHT_ESTIMATE_SW / PMI_WEIGHT_OPT_SW / OPEN_LOOP_WEIGHT_OPT_SW",
     "All three OFF 564/564. Perceive mode already DISTRIBUTED_MODE (prerequisite ON).",
     "MIMO TDD Ch.4.4.1.2 activation example: SRS_WEIGHT_ESTIMATE then PMI_WEIGHT_OPT then OPEN_LOOP_WEIGHT_OPT = 1 on 32T AAU.",
     "Largest basic-MIMO gap vs commercial. DL BF stays on template weights; large-packet UEs do not get SRS-estimated weights. Prerequisite (DISTRIBUTED_MODE) is already satisfied.",
     "DL user tput (large packet), mobility MCS/rank, BF gain.",
     "W1 trial on 3–5 loaded 32T: enable SRS_WEIGHT_ESTIMATE_SW=1 first night. Hold PMI + open-loop until SRS-weight KPI is green (do not mix weight sources same night).",
     "No extra license (basic MIMO FBFD-010003)"),
    ("MIMO-011", "Major", "Beam tracking (MISSING)", "NRDUCellBeamAlgo",
     "INTELLIGENT_BEAM_SELECTION_SW / BEAM_TRACKING_SW / BeamAdaptationSwitch",
     "INTELLIGENT_BEAM_SELECTION OFF 564; BEAM_TRACKING OFF 564; Beam Adaptation Switch OFF 564.",
     "Commercial 32T AAU with good DL mobility tput: beam tracking + intelligent beam selection ON in connected mode.",
     "BF does not track the UE after initial select. Visible on roads / mobility. Perceive mode is already correct.",
     "DL/UL user tput in mobility, drop.",
     "Enable BEAM_TRACKING_SW-1 + INTELLIGENT_BEAM_SELECTION_SW-1 on 32T trial. Keep UAV/aerial bits OFF.",
     "FBFD-010015 Beam Management"),
    ("MIMO-012", "Major", "SSB beam adapt (MISSING)", "NRDUCellAlgoSwitch",
     "BeamOptSwitch: SSB_BEAM_ADAPT_SW / SSB_BEAM_VERTICAL_COV_IMP_SW (DL_INITIAL_BEAM_SELECT already ON)",
     "SSB_BEAM_ADAPT OFF 564; SSB_BEAM_VERTICAL_COV_IMP OFF 564. SCH_BEAM_OPT ON only 9/564.",
     "Commercial AAU: keep initial DL beam select; add SSB beam adapt + vertical coverage improve so SSB follows traffic/coverage, not a frozen grid.",
     "SSB beams largely static after initial select. Density adapt is ON but horizontal/vertical adapt is OFF.",
     "Coverage holes, overlap, edge DL tput, HO.",
     "Trial SSB_BEAM_ADAPT_SW-1 + SSB_BEAM_VERTICAL_COV_IMP_SW-1 on 32T after tilt/azimuth audit (MIMO-001/002).",
     "FBFD-010015 / FOFD-010100"),
    ("MIMO-013", "Major", "Beam select opt (MISSING)", "NRDUCellBeamAlgo",
     "ChannelOptAlgoSwitch / BEAM_SELECT_OPT_SW",
     "BEAM_SELECT_OPT_SW OFF 564/564.",
     "iBeam 1.0 + Beam Mgmt TDD ≥8T: BEAM_SELECT_OPT ON with HighPrecisionBeamSwitch.",
     "Connected-mode beam pick stays coarse. Enable with iBeam 1.0 package (do not cherry-pick alone if iBeam license exists).",
     "PDSCH/PDCCH beam quality, DL tput.",
     "Enable with W3 iBeam 1.0 package. If iBeam license missing, still trial BEAM_SELECT_OPT on 32T as Beam Mgmt leftover.",
     "FOFD-081201 if with iBeam; else Beam Mgmt"),
    ("MIMO-014", "Major", "iBeam 1.0 master (MISSING)", "NRDUCellFeatureSw",
     "HighPrecisionBeamSwitch",
     "OFF 564/564. Phase2 also OFF 564. Phase3 not present as ON.",
     "ON for loaded urban 32T (interference-limited). OFF on 2T2R. Commercial good-DL-tput networks run iBeam 1.0 on the capacity layer.",
     "iBeam 1.0 never started. This is the main missing switch for interference-limited DL MIMO (precise MU sch + hybrid interference-random + SRS tight MUX + PDCCH agg compress).",
     "DL user tput in interference, MU quality, PDCCH blocking, tail packets.",
     "LST LICENSE NR0S00BEAM00. 3–5 site 32T trial HighPrecisionBeamSwitch=ON together with the iBeam 1.0 child set (MIMO-015).",
     "FOFD-081201 / NR0S00BEAM00 per cell"),
    ("MIMO-015", "Major", "iBeam 1.0 children (MISSING)", "SrsMeas / Srs / DlSch / Pdcch / DlMimo",
     "SRS_BLIND_IS_MEAS / SRS_TIGHT_MULTIPLEXING / DL_BWP_HYBRID_INTRF_RANDOM / DL_MU_PRECISE_SCH / DL_MU_ANTI_INTRF_SCH / PDCCH_AGG_LVL_COMPR / TAIL_PKT_MCS_OPT / RES_BASED_DL_ADAPT_SCH",
     "All OFF 564/564. Agg-compress threshold already 60. Beam-specific SRS denoise already ON 32T.",
     "iBeam FPD §3.4.1 turns this package ON with HighPrecisionBeamSwitch. Commercial interference-limited 32T runs the package, not the master alone.",
     "Master and all 1.0 scheduling/SRS children are off. Do not stack SRS_TIGHT_MULTIPLEXING with a new AHR SRS_IC the same night.",
     "DL interference tput, MU pairing under ISI, PDCCH, tail MCS.",
     "Enable as one package with HighPrecisionBeamSwitch on the same 3–5 sites. Watch N.SRS / PDCCH blocking / DL IBLER.",
     "FOFD-081201 / NR0S00BEAM00"),
    ("MIMO-016", "Major", "Massive MIMO multi-layer (MISSING)", "NRDUCellDlMimo",
     "HighLayerMuMimoSw: MMIMO_MULTILAYER_ENHANCE_SW / MU_MIMO_PAIRING_PREFERRED_SW / MU_RANK_BOOSTING_SW / SRS_BLIND_IS_SW / SRS_MEAS_ACCELERATING_SW",
     "All five OFF 564/564. Basic MU already ON. LAYER_16 quota already set.",
     "ON for loaded 32T AAU after MU is stable. This is what commercial networks use to convert 32T hardware into higher MU layers / pairing preference.",
     "32T hardware is under-used: no pairing-preferred, no multilayer enhance, no MU rank boosting, no SRS meas accelerate. Quota LAYER_16 cannot be consumed without these bits.",
     "DL capacity, average DL user tput, MU layers.",
     "License check then W5 (after iBeam 1.0 KPI-green): MMIMO_MULTILAYER_ENHANCE_SW-1 + MU_MIMO_PAIRING_PREFERRED_SW-1 + MU_RANK_BOOSTING_SW-1. Add SRS_MEAS_ACCELERATING with that wave.",
     "FOFD mMIMO multi-layer / layer licenses NR0S0DLEPU00"),
    ("MIMO-017", "Major", "AHR Capacity Upgrade 2.0 (MISSING)", "NRDUCellFeatureSw",
     "AhrSwitch / AHR_CAPC_UPGRADE_PHASE2_SW",
     "OFF 564/564. Phase1 + Turbo already ON 563/564.",
     "ON for loaded urban 32T after Phase1+Turbo are stable. High-resolution MU pairing / multi-dimension joint sch live here.",
     "Largest remaining AHR gap. Turbo master is ON but Capacity Upgrade (the capacity wave) never enabled. Also INTEL_PRCS_DL_MU_MIMO_PAIR / EMIMO_PRO / PDCCH_MULTI_DIM_JOINT_SCH all OFF.",
     "Loaded-hour DL cell capacity and average user tput.",
     "After Turbo SRS-IC trial is KPI-green: AHR_CAPC_UPGRADE_PHASE2_SW-1 on loaded 32T. Then INTEL_PRCS_DL_MU_MIMO_PAIR_SW / PDCCH_MULTI_DIM_JOINT_SCH as FPD child set.",
     "FOFD-061202 / NR0S00ACT200 per cell"),
    ("MIMO-018", "Major", "AHR Turbo leftover SRS-IC", "NRDUCellSrs / UlPcConfig",
     "SRS_IC_SW / SRS_INTRF_COORD / SRS_JOINT_PC (Turbo child, master already ON)",
     "SRS_IC OFF 564. SRS_TIGHT_MULTIPLEXING OFF. BEAM_SPECIFIC_SRS_DENOISE already ON ~558.",
     "Turbo 2.0 data-prep enables SRS interference coordination + SRS joint PC. Master is already ON; children are not.",
     "Turbo is half-activated: precise AMC helpers ON, SRS IC path OFF. This limits DL weight quality when SRS collides.",
     "DL tput under SRS collision; UL SRS quality.",
     "W2 on same 3–5 sites: SRS_IC_SW-1 + SRS_JOINT_PC_SW-1. Do not combine with iBeam SRS_TIGHT_MULTIPLEXING the same night.",
     "FOFD-061201 already ON (master)"),
    ("MIMO-019", "Major", "UL Boosting 1.0 (MISSING — UL; secondary for DL target)", "NRDUCellFeatureSw / UlMimo",
     "MimoFeatureSwitch UL_LOW_NOISE_SW + UL_MU_GRP_PAIR / DIFF_WAVEFORM_PAIR / UL_CORR_ACCELERATION",
     "UL_LOW_NOISE OFF 564. UL high-layer MU all OFF. Group/flex pair OFF.",
     "Commercial 32T: UL Boosting 1.0 master UL_LOW_NOISE_SW=1 with the pairing child set. Aerial boost bits stay OFF.",
     "UL Boosting completely off. UL remains basic MU + interference-random (those two are already ON). Not the first DL-tput lever, but needed for UL user tput / SRS quality that feeds DL weights.",
     "UL user tput, UL interference, SRS quality for DL BF.",
     "W4 after DL beam/iBeam waves: UL_LOW_NOISE_SW-1 + child package. Keep OFF on 2T2R. LST LICENSE NR0S00UAHR00.",
     "FOFD-091201 / NR0S00UAHR00"),
    ("MIMO-020", "Major", "DL scheduling extras (MISSING)", "NRDUCellDlSch / FeatureSw",
     "TAIL_PKT_MCS_OPT_SW / RES_BASED_DL_ADAPT_SCH_SW / LOAD_BASED_DL_EXP_SCH_SW / FREQ_SEL_SCH_SW / SMART_SCH_AND_LINK_ADAPT_SW / HIGH_CAPACITY_EXP_IMP_SW / DlBeamPriBasedSch",
     "All OFF 564. (TAIL_PKT_SCH_OPT on PDSCH is already ON — different bit.) Beam-priority schedule OFF 564. Performance Algorithm Switch both bits OFF.",
     "Commercial loaded 32T: tail-pkt MCS opt + resource-based adapt sch with iBeam; freq-sel / load-based exp sch on loaded cells; beam-priority sch when SSB/CSI beams are trusted.",
     "Scheduler is EPF+heavy-load but not interference-aware / freq-selective / beam-priority. SMART_SCH + HIGH_CAPACITY_EXP still off.",
     "DL user tput (tail packets, loaded hour, frequency-selective UEs).",
     "Enable TAIL_PKT_MCS_OPT + RES_BASED_DL_ADAPT_SCH with iBeam 1.0. Trial FREQ_SEL_SCH + LOAD_BASED_DL_EXP_SCH on loaded 32T. Hold DlBeamPriBasedSch until SSB adapt is ON. License-check SMART_SCH / HIGH_CAPACITY_EXP.",
     "iBeam FOFD-081201 for tail/res-based; others performance pack"),
    ("MIMO-021", "Minor", "SRS / UL rank leftovers", "NRDUCellSrs / UlRank / Pusch",
     "SRS_SINR_MEAS_OPT_SW / UL_RANK_FAST_DECREASE_SW / PUSCH_CE_SINR_LEVEL_ENH_SW",
     "All OFF 564. USER_CHARACTER_SRS_ADAPT already ON. SRS period SL80. SrsWeightValidityPeriod already MS400.",
     "MIMO TDD Ch.4.4.1.2 example turns these three ON with basic MIMO weights.",
     "SRS meas opt off → slower UL rank drop and weaker PUSCH CE in interference. Weight validity already correct (MS400).",
     "UL BLER, UL tput, DL BF tracking quality (SRS).",
     "W1 with SRS weight estimate: enable these three on the same 32T trial.",
     "No extra license (basic MIMO)"),
    ("MIMO-022", "Minor", "CSI-RS sweeping / Type2", "NRDUCellCsirs / PdschPrecode",
     "CSIRS_BEAM_SWEEPING_SW / R15_TYPE2_SW / DL_ROBUST_WEIGHT_SW",
     "CSI sweep OFF 564. Type2 codebook all OFF. DL_ROBUST_WEIGHT OFF (iBeam 2.0 child). 8PORT+TYPE0+FD already ON.",
     "CSI-RS sweeping often ON for 8-port CSI on 32T if SSB is static. Type2 / robust weight = later high-res CSI trial (UE CAP + iBeam 2.0).",
     "Static SSB + no CSI sweep + 8-port = coarse spatial info vs commercial high-tput clusters that sweep CSI or run iBeam robust weight.",
     "DL BF gain, MU pair quality.",
     "After SSB scenario freeze: trial CSIRS_BEAM_SWEEPING_SW-1. Robust weight with iBeam 2.0. Type2 = 20-cell trial only.",
     "iBeam 2.0 FOFD-091200 for robust weight"),
    ("MIMO-023", "Minor", "Coverage scenario mix", "NRDUCellTrpBeam",
     "CoverageScenario",
     "SCENARIO_8 367; EXPAND_SCENARIO_1 213; SCENARIO_7 1. Connected-mode scenario DEFAULT 564. AIR_1_GND_4 cluster-wide.",
     "Scenario must match mechanical install. Connected-mode DEFAULT is OK if broadcast scenario is right. Air template should not steer ground beams if no UAV.",
     "Unclear if 213 Expand-1 were RF-planned or copied. 1 cell on SCENARIO_7. Air template looks copy-pasted (aerial algos are OFF — good).",
     "SSB coverage, overlap, HO.",
     "RF audit SCENARIO_8 vs EXPAND_SCENARIO_1 vs design. Fix the SCENARIO_7 outlier. Confirm AIR_1_GND_4 vs ground-only.",
     "FOFD-010100 3D coverage"),
    ("MIMO-024", "Minor", "MU power / gathering / MCS policy", "NRDUCellPdsch / DlMimo",
     "MuMimoPwrAllocSwitch / DL Gathering MU / DL MIMO MCS Optimization Policy / MuMimoOptSwith",
     "MU power alloc OFF 564. Gathering MU OFF. MIMO MCS policy OFF. MU-MIMO Optimization Switch OFF 564. Low-time-corr MU already ON 563.",
     "Low-time-corr MU ON is commercial-good for mobility. Gathering / MCS-policy / MuMimoOpt = optional after multilayer. Power alloc after multilayer enhance.",
     "Power-domain MU still default-off. Do not flip MuMimoOpt cluster-wide on night 1.",
     "MU fairness, edge MCS.",
     "After multilayer enhance: trial MU power alloc + MIMO MCS policy. Keep low-time-corr ON.",
     "FOFD-010010 / multilayer"),
    ("MIMO-025", "Info", "DAS / Fusion / mmWave — N/A", "gNBMimoClusterCell / TrpMmwavBeam / DM_MIMO_*",
     "DM_MIMO_SERVICE_SWITCH / Fusion cluster / FR2 beam MOs",
     "All DM_MIMO bits OFF. gNBCluster / gNBMimoClusterCell ~4 rows. mmWave/FDD beam MOs empty. All cells FR1 TDD n41.",
     "OFF unless multi-AAU DAS / Fusion Cell / FR2 is deployed. This network is n41 FR1 40 MHz 32T NSA.",
     "Correct. mmWave (Step10) is not applicable. Do not copy FR2 / VOL_BASED_BEAM_MULTIPLEX.",
     "N/A",
     "Keep OFF. Do not enable DAS/Fusion/mmWave templates on this cluster.",
     "N/A"),
    ("MIMO-026", "Info", "4G dump", "—",
     "LTE MIMO / TM9 / MeNB",
     "4G dump not parsed for this MIMO sheet. NR n41 32T beam/MU/iBeam/AHR switches are 5G MOs.",
     "4G dump is needed for NSA Option 3x (X2, combo MO, split bearer) — already covered in the NSA inconsistency report. It does not gate NR MIMO air-interface switches.",
     "No 4G MIMO parameter is the missing lever for NR 32T beam/scheduling performance on this n41 capacity layer.",
     "NSA EN-DC add vs NR DL tput are different problems.",
     "Do not wait on 4G MIMO TM changes to improve NR MIMO. Use 5G dump actions below. NSA MeNB issues stay on the NSA report.",
     "—"),
    ("MIMO-027", "Introduce", "iBeam 2.0 / 3.0", "NRDUCellFeatureSw / PdschPrecode",
     "HighPrecisionBeamPhase2Sw / DL_ROBUST_WEIGHT / HighPrecisionBeamPhase3 / DL_HYBRID_PRECODING",
     "Phase2 OFF 564. DL_ROBUST_WEIGHT OFF. Hybrid precoding OFF. Phase3 not ON.",
     "iBeam 2.0 after 1.0 KPI-green (robust weight). iBeam 3.0 last wave.",
     "Correctly not started — 1.0 is missing. Listed so the roadmap is visible.",
     "Second-wave DL robustness in interference.",
     "Do not enable 2.0/3.0 now. Sequence: 1.0 (W3) → 2.0 → 3.0 on the same trial sites.",
     "FOFD-091200 / FOFD-100200"),
    ("MIMO-028", "OK", "PUSCH beam-domain (already ON)", "NRDUCellAlgoSwitch",
     "FullChnCovEnhSwitch PUSCH_BEAM_DOMAIN_ENH / PUSCH_TIME_DOMAIN_ENH / PUSCH_BEAM_SEL_OPT",
     "PUSCH_BEAM_DOMAIN_ENH + TIME_DOMAIN_ENH ON 563/564. PUSCH_BEAM_SEL_OPT ON 564. SSB_BEAM_ENH OFF.",
     "PUSCH beam-domain enhance ON is the commercial UL coverage baseline for 32T.",
     "Aligned. SSB_BEAM_ENH still OFF (optional).",
     "UL coverage already using beam-domain PUSCH.",
     "Skip. Optional later: SSB_BEAM_ENH with RF agreement.",
     "—"),
]


# ---------------------------------------------------------------------------
# 11-col MML (proposed). Action: Check / Skip / Enable / Hold / Fix / Seq
# ---------------------------------------------------------------------------
def mml_rows():
    pid = "{NrDuCellId}"
    return [
        (1, "TDD", "W0 Check", "Seq", "", "", "", "", "",
         "Night 0 on 3–5 loaded 32T macros. Confirm 32T32R, AHR Phase1+Turbo ON, tilt≠255. Exclude DHTIAA1 / DHAPT11 / DHTEJ34.", ""),
        (2, "TDD", "MIMO (TDD)", "Check",
         f"LST NRDUCELLTRP: NrDuCellId={pid};",
         "NRDUCellTrp", "TxRxMode",
         "32T32R 564; 2T2R 17 (DHTIAA1 / DHAPT11 / DHTEJ34)",
         "32T32R",
         "Abort site if 2T2R. Trial is 32T AAU only.",
         "Do not apply on 2T2R"),
        (3, "TDD", "Beam RF", "Fix",
         f"LST NRDUCELLTRPBEAM: NrDuCellTrpId={{NrDuCellTrpId}};",
         "NRDUCellTrpBeam", "Tilt / Azimuth",
         "Tilt 255 on 28 cells; Azimuth 0° on 580/581",
         "Design tilt (not 255); azimuth per RF plan",
         "P0 before any MIMO switch. 255 = beamforming not applied.",
         "No extra license (RF)"),
        (4, "TDD", "AHR", "Check",
         f"LST NRDUCELLFEATURESW: NrDuCellId={pid};",
         "NRDUCellFeatureSw", "AhrSwitch",
         "AHR_PHASE1 + AHR_EXP_TURBO_PHASE2 ON 563/564",
         "Both ON",
         "If Turbo OFF, pick another 32T site.",
         "FOFD-051301 / 061201 already ON"),
        (5, "TDD", "License", "Check",
         "LST LICENSE:;",
         "License", "NR0S00BEAM00 / NR0S00UAHR00 / NR0S00ACT200 / NR0S0DLEPU00",
         "Not in CME dump — confirm on trial gNB",
         "present for trial cells",
         "iBeam 1.0, UL Boosting, AHR CU 2.0, layer capacity.",
         "LST LICENSE first"),
        (6, "TDD", "MIMO (TDD)", "Skip",
         f"LST NRDUCELLPDSCH: NrDuCellId={pid};",
         "NRDUCellPdsch", "MaxMimoLayerNum",
         "LAYER_16 557/564 — do not send LAYER_8",
         "LAYER_16 (keep)",
         "Doc example LAYER_8 is a lab floor. Never downgrade.",
         "FOFD-010020 layer licenses"),
        (7, "TDD", "W0 Skip (already ON)", "Seq", "", "", "", "", "",
         "Do not re-send. Live DHK 32T already matches these RAN10.1 examples.", ""),
        (8, "TDD", "MIMO (TDD)", "Skip",
         f"MOD NRDUCELLALGOSWITCH: NrDuCellId={pid}, SuMimoSwitch=DL_SU_MULTI_LAYER_SW-1&UL_SU_MULTI_LAYER_SW-1;",
         "NRDUCellAlgoSwitch", "SuMimoSwitch",
         "DL+UL SU multi-layer ON 564/564",
         "already ON",
         "SU-MIMO already ON.",
         "FOFD-010020 already ON"),
        (9, "TDD", "MIMO (TDD)", "Skip",
         f"MOD NRDUCELLALGOSWITCH: NrDuCellId={pid}, MuMimoSwitch=DL_MU_MIMO_SW-1&UL_MU_MIMO_SW-1&PDCCH_MU_SW-1;",
         "NRDUCellAlgoSwitch", "MuMimoSwitch",
         "DL+UL MU + PDCCH_MU ON 563/564",
         "already ON",
         "Core MU already ON 32T.",
         "FOFD-010010 already ON"),
        (10, "TDD", "MIMO (TDD)", "Skip",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, BeamPerceiveMode=DISTRIBUTED_MODE;",
         "NRDUCellBeamAlgo", "BeamPerceiveMode",
         "DISTRIBUTED_MODE 564/564",
         "DISTRIBUTED_MODE",
         "Already correct. Prerequisite of weight opt.",
         "No extra license"),
        (11, "TDD", "AHR", "Skip",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, AhrSwitch=AHR_PHASE1_SW-1&AHR_EXP_TURBO_PHASE2_SW-1;",
         "NRDUCellFeatureSw", "AhrSwitch",
         "Phase1+Turbo ON 563/564",
         "already ON",
         "Do not re-send. Next AHR wave is Capacity Upgrade (W5).",
         "FOFD-051301 / 061201"),
        (12, "TDD", "Beam Mgmt", "Skip",
         f"MOD NRDUCELLALGOSWITCH: NrDuCellId={pid}, AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-1;",
         "NRDUCellAlgoSwitch", "AdaptiveEdgeExpEnhSwitch",
         "DL_PMI_SRS_ADAPT ON 564/564",
         "already ON",
         "PMI/SRS adapt already ON. LOW_SNR_CHN_DENOISE stays OFF (Turbo constraint).",
         "FBFD-010003"),
        (13, "TDD", "W1 Enable — Basic MIMO leftovers (this trial night)", "Seq", "", "", "", "", "",
         "Lowest-risk mismatch vs MIMO TDD activation example. Improves DL BF weights + SRS quality. 32T only.", ""),
        (14, "TDD", "MIMO (TDD)", "Enable",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1;",
         "NRDUCellBeamAlgo", "WeightAlgoSwitch",
         "SRS_WEIGHT_ESTIMATE OFF 564/564",
         "SRS_WEIGHT_ESTIMATE_SW-1",
         "MISSING. SRS-based weight estimate for DL large-packet UEs. Perceive mode already DISTRIBUTED.",
         "No extra license (basic MIMO)"),
        (15, "TDD", "MIMO (TDD)", "Hold",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, WeightAlgoSwitch=PMI_WEIGHT_OPT_SW-1;",
         "NRDUCellBeamAlgo", "WeightAlgoSwitch",
         "PMI_WEIGHT_OPT OFF 564/564",
         "PMI_WEIGHT_OPT_SW-1",
         "MISSING vs FPD example. Hold first night — do not mix PMI + SRS weight sources.",
         "No extra license (basic MIMO)"),
        (16, "TDD", "MIMO (TDD)", "Hold",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, WeightAlgoSwitch=OPEN_LOOP_WEIGHT_OPT_SW-1;",
         "NRDUCellBeamAlgo", "WeightAlgoSwitch",
         "OPEN_LOOP_WEIGHT_OPT OFF 564/564",
         "OPEN_LOOP_WEIGHT_OPT_SW-1",
         "MISSING vs FPD example. Hold first night (open-loop mix).",
         "No extra license (basic MIMO)"),
        (17, "TDD", "MIMO (TDD)", "Enable",
         f"MOD NRDUCELLSRS: NrDuCellId={pid}, SrsAlgoSwitch=SRS_SINR_MEAS_OPT_SW-1;",
         "NRDUCellSrs", "SrsAlgoSwitch",
         "SRS_SINR_MEAS_OPT OFF 564/564",
         "SRS_SINR_MEAS_OPT_SW-1",
         "MISSING. SRS SINR measurement opt (MIMO TDD example ON).",
         "No extra license"),
        (18, "TDD", "MIMO (TDD)", "Enable",
         f"MOD NRDUCELLULRANK: NrDuCellId={pid}, UlRankAlgoSw=UL_RANK_FAST_DECREASE_SW-1;",
         "NRDUCellUlRank", "UlRankAlgoSw",
         "UL_RANK_FAST_DECREASE OFF 564/564",
         "UL_RANK_FAST_DECREASE_SW-1",
         "MISSING. Fast UL rank drop on poor SRS. Protects UL BLER.",
         "No extra license"),
        (19, "TDD", "MIMO (TDD)", "Enable",
         f"MOD NRDUCELLPUSCH: NrDuCellId={pid}, PuschPerformanceSwitch=PUSCH_CE_SINR_LEVEL_ENH_SW-1;",
         "NRDUCellPusch", "PuschPerformanceSwitch",
         "PUSCH_CE_SINR_LEVEL_ENH OFF 564/564",
         "PUSCH_CE_SINR_LEVEL_ENH_SW-1",
         "MISSING. PUSCH CE SINR-level enhance. MIMO TDD example ON.",
         "No extra license"),
        (20, "TDD", "Beam Mgmt", "Enable",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, BeamOptAlgoSwitch=BEAM_TRACKING_SW-1&INTELLIGENT_BEAM_SELECTION_SW-1;",
         "NRDUCellBeamAlgo", "BeamOptAlgoSwitch",
         "BEAM_TRACKING + INTELLIGENT_BEAM_SELECTION OFF 564/564",
         "BEAM_TRACKING_SW-1 + INTELLIGENT_BEAM_SELECTION_SW-1",
         "MISSING. Connected-mode BF tracking. Main mobility DL-tput lever after weights.",
         "FBFD-010015"),
        (21, "TDD", "Beam Mgmt", "Enable",
         f"MOD NRDUCELLALGOSWITCH: NrDuCellId={pid}, BeamOptSwitch=SSB_BEAM_ADAPT_SW-1&SSB_BEAM_VERTICAL_COV_IMP_SW-1;",
         "NRDUCellAlgoSwitch", "BeamOptSwitch",
         "SSB_BEAM_ADAPT + VERTICAL_COV_IMP OFF 564/564 (DL_INITIAL_BEAM_SELECT already ON)",
         "SSB_BEAM_ADAPT_SW-1 + SSB_BEAM_VERTICAL_COV_IMP_SW-1",
         "MISSING. SSB still static after initial select. Do after tilt/azimuth fix.",
         "FBFD-010015"),
        (22, "TDD", "W2 Enable — AHR Turbo leftover (week 2, same sites)", "Seq", "", "", "", "", "",
         "Turbo master already ON; SRS IC path still OFF. Do not stack with iBeam SRS_TIGHT_MULTIPLEXING same night.", ""),
        (23, "TDD", "AHR Turbo", "Enable",
         f"MOD NRDUCELLSRS: NrDuCellId={pid}, SrsDetectionAlgoSwitch=SRS_IC_SW-1;",
         "NRDUCellSrs", "SrsDetectionAlgoSwitch",
         "SRS_IC OFF 564; BEAM_SPECIFIC_SRS_DENOISE already ON",
         "SRS_IC_SW-1",
         "MISSING Turbo child. Watch N.SRS collision / UL SRS quality.",
         "FOFD-061201 master already ON"),
        (24, "TDD", "AHR Turbo", "Enable",
         f"MOD NRDUCELLULPCCONFIG: NrDuCellId={pid}, UlPwrCtrlAlgoSwitch=SRS_JOINT_PC_SW-1;",
         "NRDUCellUlPcConfig", "UlPwrCtrlAlgoSwitch",
         "SRS_JOINT_PC OFF 564/564",
         "SRS_JOINT_PC_SW-1",
         "Recommended with SRS IC. Live OFF.",
         "FOFD-061201"),
        (25, "TDD", "W3 Enable — iBeam 1.0 (week 3) — main DL interference package", "Seq", "", "", "", "", "",
         "Massive MIMO iBeam RAN10.1 §3.4.1.2. Master then children. This is the biggest missing DL-tput switch pack.", ""),
        (26, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, HighPrecisionBeamSwitch=ON;",
         "NRDUCellFeatureSw", "HighPrecisionBeamSwitch",
         "OFF 564/564",
         "ON",
         "MISSING master. iBeam 1.0. Loaded urban 32T.",
         "FOFD-081201 / NR0S00BEAM00 — LST LICENSE"),
        (27, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLSRSMEAS: NrDuCellId={pid}, SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-1;",
         "NRDUCellSrsMeas", "SrsMeasOptSwitch",
         "SRS_BLIND_IS_MEAS OFF 564/564",
         "SRS_BLIND_IS_MEAS_SW-1",
         "MISSING. Blind IS for SRS measurement.",
         "FOFD-081201"),
        (28, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLSRS: NrDuCellId={pid}, SrsAlgoSwitch=SRS_TIGHT_MULTIPLEXING_SW-1;",
         "NRDUCellSrs", "SrsAlgoSwitch",
         "SRS_TIGHT_MULTIPLEXING OFF 564/564",
         "SRS_TIGHT_MULTIPLEXING_SW-1",
         "MISSING. Do not stack with new SRS_IC the same night.",
         "FOFD-081201"),
        (29, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLBEAMALGO: NrDuCellId={pid}, ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1;",
         "NRDUCellBeamAlgo", "ChannelOptAlgoSwitch",
         "BEAM_SELECT_OPT OFF 564/564",
         "BEAM_SELECT_OPT_SW-1",
         "MISSING. Beam selection optimization.",
         "FOFD-081201"),
        (30, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-1;",
         "NRDUCellDlMimo", "DLMuMimoSchSupplementSw",
         "DL_MU_PRECISE_SCH OFF 564/564",
         "DL_MU_PRECISE_SCH_SW-1",
         "MISSING. DL MU precise scheduling — MU efficiency under interference.",
         "FOFD-081201"),
        (31, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, DLMuMimoSchSupplementSw=DL_MU_ANTI_INTRF_SCH_SW-1;",
         "NRDUCellDlMimo", "DLMuMimoSchSupplementSw",
         "DL_MU_ANTI_INTRF_SCH OFF 564/564",
         "DL_MU_ANTI_INTRF_SCH_SW-1",
         "MISSING. DL MU anti-interference scheduling.",
         "FOFD-081201"),
        (32, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLSCH: NrDuCellId={pid}, DlSchAlgoSwitch=DL_BWP_HYBRID_INTRF_RANDOM_SW-1;",
         "NRDUCellDlSch", "DlSchAlgoSwitch",
         "DL_BWP_HYBRID_INTRF_RANDOM OFF 564/564",
         "DL_BWP_HYBRID_INTRF_RANDOM_SW-1",
         "MISSING. DL multi-BWP interference randomization.",
         "FOFD-081201"),
        (33, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLSCH: NrDuCellId={pid}, DlSchAlgoSwitch=TAIL_PKT_MCS_OPT_SW-1;",
         "NRDUCellDlSch", "DlSchAlgoSwitch",
         "TAIL_PKT_MCS_OPT OFF 564/564 (PDSCH TAIL_PKT_SCH_OPT already ON — different bit)",
         "TAIL_PKT_MCS_OPT_SW-1",
         "MISSING. Tail-packet MCS opt. Complements already-ON TAIL_PKT_SCH_OPT.",
         "FOFD-081201"),
        (34, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLSCH: NrDuCellId={pid}, DlSchAlgoSwitch=RES_BASED_DL_ADAPT_SCH_SW-1;",
         "NRDUCellDlSch", "DlSchAlgoSwitch",
         "RES_BASED_DL_ADAPT_SCH OFF 564/564",
         "RES_BASED_DL_ADAPT_SCH_SW-1",
         "MISSING. Resource-based adaptive DL scheduling.",
         "FOFD-081201"),
        (35, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLDLSCH: NrDuCellId={pid}, DlSchAlgoSwitch=DL_RLC_STAT_RPT_MERGE_SCH_SW-1;",
         "NRDUCellDlSch", "DlSchAlgoSwitch",
         "DL_RLC_STAT_RPT_MERGE_SCH OFF 564/564",
         "DL_RLC_STAT_RPT_MERGE_SCH_SW-1",
         "MISSING. Merged DL RLC status-report scheduling.",
         "FOFD-081201"),
        (36, "TDD", "iBeam 1.0", "Enable",
         f"MOD NRDUCELLPDCCH: NrDuCellId={pid}, PdcchAlgoSwitch=PDCCH_AGG_LVL_COMPR_SW-1;",
         "NRDUCellPdcch", "PdcchAlgoSwitch",
         "PDCCH_AGG_LVL_COMPR OFF 564 (threshold already 60)",
         "PDCCH_AGG_LVL_COMPR_SW-1",
         "MISSING. PDCCH aggregation-level compression. Threshold already staged at 60.",
         "FOFD-081201"),
        (37, "TDD", "W3b Hold — extra DL sch (same week only if load is high)", "Seq", "", "", "", "", "",
         "Optional scheduling bits. Not in the iBeam 1.0 minimum package. Trial on loaded cells only.", ""),
        (38, "TDD", "DL sch", "Hold",
         f"MOD NRDUCELLDLSCH: NrDuCellId={pid}, DlSchAlgoSwitch=FREQ_SEL_SCH_SW-1;",
         "NRDUCellDlSch", "DlSchAlgoSwitch",
         "FREQ_SEL_SCH OFF 564/564",
         "FREQ_SEL_SCH_SW-1",
         "MISSING vs loaded commercial. Hold until CSI quality is stable (after W1/W3).",
         "Performance pack"),
        (39, "TDD", "DL sch", "Hold",
         f"MOD NRDUCELLDLSCH: NrDuCellId={pid}, DlSchAlgoSwitch=LOAD_BASED_DL_EXP_SCH_SW-1;",
         "NRDUCellDlSch", "DlSchAlgoSwitch",
         "LOAD_BASED_DL_EXP_SCH OFF 564/564",
         "LOAD_BASED_DL_EXP_SCH_SW-1",
         "MISSING. Load-based DL experience scheduling. Hold for busy-hour 32T.",
         "Performance pack"),
        (40, "TDD", "DL sch", "Hold",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, PerformanceAlgoSwitch=SMART_SCH_AND_LINK_ADAPT_SW-1&HIGH_CAPACITY_EXP_IMP_SW-1;",
         "NRDUCellFeatureSw", "PerformanceAlgoSwitch",
         "Both bits OFF 564/564",
         "SMART_SCH_AND_LINK_ADAPT_SW-1 + HIGH_CAPACITY_EXP_IMP_SW-1",
         "MISSING performance-algorithm pack. License + Huawei TAC before cluster.",
         "Performance algorithm license"),
        (41, "TDD", "W4 Enable — UL Boosting 1.0 (week 4)", "Seq", "", "", "", "", "",
         "UL path. Helps SRS quality that feeds DL weights. Skip RedCap / BWP2. 32T only.", ""),
        (42, "TDD", "UL Boosting 1.0", "Enable",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, MimoFeatureSwitch=UL_LOW_NOISE_SW-1;",
         "NRDUCellFeatureSw", "MimoFeatureSwitch",
         "UL_LOW_NOISE_SW OFF 564/564",
         "UL_LOW_NOISE_SW-1",
         "MISSING master. Uplink Boosting 1.0.",
         "FOFD-091201 / NR0S00UAHR00 — LST LICENSE"),
        (43, "TDD", "UL Boosting 1.0", "Enable",
         f"MOD NRDUCELLULMIMO: NrDuCellId={pid}, UlMuMimoAlgoSwitch=UL_MU_GRP_PAIR_SW-1;",
         "NRDUCellUlMimo", "UlMuMimoAlgoSwitch",
         "UL_MU_GRP_PAIR OFF 564/564",
         "UL_MU_GRP_PAIR_SW-1",
         "MISSING. UL MU grouping/pairing.",
         "FOFD-091201"),
        (44, "TDD", "UL Boosting 1.0", "Enable",
         f"MOD NRDUCELLULMIMO: NrDuCellId={pid}, UlMuMimoAlgoSwitch=DIFF_WAVEFORM_PAIR_SW-1;",
         "NRDUCellUlMimo", "UlMuMimoAlgoSwitch",
         "DIFF_WAVEFORM_PAIR OFF 564/564",
         "DIFF_WAVEFORM_PAIR_SW-1",
         "MISSING. Mixed-waveform UL MU pairing.",
         "FOFD-091201"),
        (45, "TDD", "UL Boosting 1.0", "Enable",
         f"MOD NRDUCELLULMIMO: NrDuCellId={pid}, UlMuMimoAlgoSwitch=UL_CORR_ACCELERATION_SW-1;",
         "NRDUCellUlMimo", "UlMuMimoAlgoSwitch",
         "UL_CORR_ACCELERATION OFF 564/564",
         "UL_CORR_ACCELERATION_SW-1",
         "MISSING. UL correlation-acceleration pairing.",
         "FOFD-091201"),
        (46, "TDD", "W5 Hold — multilayer + AHR Capacity Upgrade (after trial KPI-green)", "Seq", "", "", "", "", "",
         "Not first night. Needs load + layer license. Converts 32T hardware into higher MU layers. After W2 SRS-IC + W3 iBeam 1.0.", ""),
        (47, "TDD", "MM Multi-Layer", "Hold",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;",
         "NRDUCellDlMimo", "HighLayerMuMimoSw",
         "MMIMO_MULTILAYER_ENHANCE OFF 564/564",
         "MMIMO_MULTILAYER_ENHANCE_SW-1",
         "MISSING. Master DL multi-layer enhance. LAYER_16 quota already set — this bit spends it.",
         "FOFD multilayer / NR0S0DLEPU00"),
        (48, "TDD", "MM Multi-Layer", "Hold",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, HighLayerMuMimoSw=MU_MIMO_PAIRING_PREFERRED_SW-1;",
         "NRDUCellDlMimo", "HighLayerMuMimoSw",
         "MU_MIMO_PAIRING_PREFERRED OFF 564/564",
         "MU_MIMO_PAIRING_PREFERRED_SW-1",
         "MISSING. Prefer MU pairing when gain exists.",
         "FOFD multilayer"),
        (49, "TDD", "MM Multi-Layer", "Hold",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, HighLayerMuMimoSw=MU_RANK_BOOSTING_SW-1;",
         "NRDUCellDlMimo", "HighLayerMuMimoSw",
         "MU_RANK_BOOSTING OFF 564/564",
         "MU_RANK_BOOSTING_SW-1",
         "MISSING. MU rank boosting.",
         "FOFD multilayer"),
        (50, "TDD", "MM Multi-Layer", "Hold",
         f"MOD NRDUCELLDLMIMO: NrDuCellId={pid}, HighLayerMuMimoSw=SRS_MEAS_ACCELERATING_SW-1;",
         "NRDUCellDlMimo", "HighLayerMuMimoSw",
         "SRS_MEAS_ACCELERATING OFF 564/564",
         "SRS_MEAS_ACCELERATING_SW-1",
         "MISSING. Faster SRS meas for multilayer.",
         "FOFD multilayer"),
        (51, "TDD", "AHR CU 2.0", "Hold",
         f"MOD NRDUCELLFEATURESW: NrDuCellId={pid}, AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1;",
         "NRDUCellFeatureSw", "AhrSwitch",
         "AHR_CAPC_UPGRADE_PHASE2 OFF 564/564 (Phase1+Turbo already ON)",
         "AHR_CAPC_UPGRADE_PHASE2_SW-1",
         "MISSING. AHR Capacity Upgrade 2.0 master. Loaded 32T after Turbo SRS-IC.",
         "FOFD-061202 / NR0S00ACT200"),
        (52, "TDD", "AHR CU 2.0", "Hold",
         f"MOD NRDUCELLALGOSWITCH: NrDuCellId={pid}, MuMimoSwitch=PDCCH_MULTI_DIM_JOINT_SCH_SW-1;",
         "NRDUCellAlgoSwitch", "MuMimoSwitch",
         "PDCCH_MULTI_DIM_JOINT_SCH OFF 564/564",
         "PDCCH_MULTI_DIM_JOINT_SCH_SW-1",
         "MISSING Capacity-Upgrade child. Multi-dimension joint PDCCH/PDSCH sch.",
         "FOFD-061202"),
        (53, "TDD", "Do not send", "Skip",
         f"MOD NRDUCELLPDSCH: NrDuCellId={pid}, MaxMimoLayerNum=LAYER_8;",
         "NRDUCellPdsch", "MaxMimoLayerNum",
         "LAYER_16 live — DO NOT SEND LAYER_8",
         "DO NOT SEND",
         "FPD sample only. Sending LAYER_8 would cut peak layers.",
         "—"),
        (54, "TDD", "mmWave / DAS / Fusion", "Skip",
         "—",
         "NRDUCellTrpMmwavBeam / DM_MIMO / gNBMimoClusterCell",
         "FR2 / DAS / Fusion",
         "Empty / OFF — n41 FR1 40 MHz only",
         "keep OFF",
         "Not applicable on this network. Sibling analysis: mmWave is not FR2.",
         "N/A"),
    ]


def patch_cover(wb):
    ws = wb["0. Cover & Index"]
    # Banner
    ws["A1"].value = "  5G MIMO (all features together)  —  Deployment Workbook  v2.0"
    # How-to-use row: find and extend
    last = ws.max_row + 2
    r = last
    r = section(ws, r, 10, "v2.0 addition — dump vs commercial MIMO inconsistency (LAST SHEET)")
    r = note_bar(ws, r, 10,
                 "Sheet 14 compares the live DHK 5G CME dump (8 Sep 2026, 564×32T32R n41 NSA) against a commercial "
                 "32T NSA reference with good DL user throughput (Huawei RAN10.1 FPD activation examples for MIMO TDD / "
                 "Beam Management / AHR Capacity Upgrade / iBeam). 4G dump is not required for NR MIMO air-interface switches. "
                 "v1 file MIMO_Deployment.xlsx is unchanged. This file is MIMO_Deployment_v2.0.xlsx.")
    r = headers(ws, r, ["#", "Sheet", "Maps to", "What you will find"] + [""] * 6)
    r = table_row(ws, r,
                  ["14", "14. MIMO Incon Report", "Live dump 8 Sep 2026 DHK vs commercial 32T good-DL-tput",
                   "Missing beam / scheduling / multilayer / iBeam / AHR CU switches + 11-col proposed MML"] + [""] * 6,
                  fills=[PALE_ORANGE] * 10, height=36)
    merge(ws, r - 1, 4, r - 1, 10)
    return r


def build_incon_sheet(wb):
    ws = wb.create_sheet(SHEET_NAME)
    setup_sheet(ws, SHEET_NAME)
    set_widths(ws, INCON_WIDTHS)
    ws.oddHeader.left.text = "5G MIMO Inconsistency Report v2.0 (dump vs commercial good DL tput)"
    ws.oddFooter.left.text = "Dump: 5G CME 8 Sep 2026 DHK · 564×32T32R n41 FR1 TDD 40 MHz NSA Option 3x · 2T2R excluded from trials"
    ws.freeze_panes = "A4"
    ws.sheet_properties.tabColor = "C00000"

    r = 1
    r = banner(ws, r, INCON_COLS,
               "  14.  MIMO Inconsistency Report  v2.0  —  Live dump vs commercial 32T (good DL user throughput)",
               size=16, height=32)
    r = note_bar(ws, r, INCON_COLS,
                 "Reference = established commercial Huawei NSA Option 3x 32T AAU cluster with good DL user throughput "
                 "(RAN10.1 FPD activation examples for MIMO TDD + Beam Management + AHR + iBeam, loaded urban). "
                 "Dump = 5G CME ConfigurationData 8 Sep 2026 DHK from repo 5G-RAN-BASIC-TO-ADVANCE_OPTIMIZATION "
                 "(branch Configuration-Dump). Counts below are 32T cells (n=564) unless noted. "
                 "Indoor 2T2R DHTIAA1 / DHAPT11 / DHTEJ34 stay OFF. mmWave/FR2 not applicable (n41 FR1). "
                 "Do not stack AHR SRS_IC + iBeam SRS_TIGHT_MULTIPLEXING the same night.")

    r = section(ws, r, INCON_COLS, "A.  Dump identity and verdict")
    r = headers(ws, r, ["Item", "Value"] + [""] * 9)
    identity = [
        ("5G dump", "Dump_5G_ConfigurationData__8Sep26_DHK.xlsb  ·  CME 8 Sep 2026  ·  DHK  ·  NSA Option 3x  ·  n41 FR1 TDD 40 MHz"),
        ("Hardware", "564 × 32T32R AAU  +  17 × 2T2R indoor (DHTIAA1 / DHAPT11 / DHTEJ34)  ·  trial scope = 32T only"),
        ("4G dump", "Not required for NR MIMO air-interface (beam / MU / iBeam / AHR / scheduler). NSA MeNB issues stay on the NSA report."),
        ("Commercial target", "Loaded urban 32T NSA with good DL user throughput = FPD ON-set for weights + beam tracking + iBeam 1.0 + multilayer + AHR CU 2.0, with SU/MU/AHR Phase1+Turbo already live (as in this dump)."),
        ("Verdict", "Foundation is already commercial-grade (SU/MU/AHR Phase1+Turbo/LAYER_16/DISTRIBUTED_MODE). DL user tput is limited by MISSING beam-weight + beam-tracking + iBeam 1.0 scheduling + multilayer/AHR-CU bits — not by missing MU or AHR Phase1."),
        ("Do not send", "LAYER_8 (live is LAYER_16). mmWave / DAS / Fusion / 2T2R MU. Do not re-enable SU/MU/AHR Phase1/Turbo."),
    ]
    for a, b in identity:
        r = put11(ws, r, [a, b] + [""] * 9, fills=[PALE_BLUE, WHITE] + [WHITE] * 9,
                  bolds=[True] + [False] * 10, height=32)
        merge(ws, r - 1, 2, r - 1, INCON_COLS)

    r = section(ws, r, INCON_COLS, "B.  Finding counts (this MIMO sheet only)")
    counts = {"Critical": 0, "Major": 0, "Minor": 0, "Introduce": 0, "OK": 0, "Info": 0}
    for rec in FINDINGS:
        counts[rec[1]] = counts.get(rec[1], 0) + 1
    r = headers(ws, r, ["Severity", "Count", "Meaning for DL MIMO tput"] + [""] * 8)
    meanings = {
        "Critical": "RF tilt/azimuth — fix before any switch trial (beams not applied on tilt=255).",
        "Major": "Missing switch vs commercial good-DL-tput 32T (beam / iBeam / multilayer / AHR CU / scheduler).",
        "Minor": "Helps DL/UL but second wave (SRS meas, CSI sweep, scenario mix, MU power).",
        "Introduce": "Roadmap (iBeam 2.0/3.0) — do not enable until 1.0 is green.",
        "OK": "Already matches commercial — Skip / do not re-send.",
        "Info": "Out of scope (DAS/Fusion/mmWave/4G dump).",
    }
    for sev in ["Critical", "Major", "Minor", "Introduce", "OK", "Info"]:
        r = put11(ws, r, [sev, counts[sev], meanings[sev]] + [""] * 8,
                  fills=[sev_color(sev)] * 11, bolds=[True, True, False] + [False] * 8, height=24)
        merge(ws, r - 1, 3, r - 1, INCON_COLS)
    r = put11(ws, r, ["Total", sum(counts.values()), "MIMO-focused only (not the full NSA 193-finding report)"] + [""] * 8,
              fills=[NAVY] * 11, bolds=[True] * 11, font_color=WHITE, height=22)
    merge(ws, r - 1, 3, r - 1, INCON_COLS)

    r = section(ws, r, INCON_COLS, "C.  What we are missing vs commercial (ranked for DL user throughput)")
    r = bullets(ws, r, INCON_COLS, [
        "P0 RF: fix 28× tilt=255 and audit cluster-wide azimuth=0° — otherwise beam switches cannot help those cells.",
        "P1 Beam weights (Step1 leftover): SRS_WEIGHT_ESTIMATE_SW is the first missing BF lever (PMI/open-loop HOLD same night). Prerequisite DISTRIBUTED_MODE is already ON.",
        "P1 Beam tracking (Step5 leftover): BEAM_TRACKING_SW + INTELLIGENT_BEAM_SELECTION_SW — connected-mode BF follow. SSB_BEAM_ADAPT + VERTICAL_COV_IMP after tilt audit.",
        "P1 iBeam 1.0 (Step7) — THE main missing DL-interference package: HighPrecisionBeamSwitch + precise/anti-intrf MU sch + hybrid interference-random + tail-pkt MCS + res-based adapt sch + beam-select opt + SRS tight MUX. This is what commercial interference-limited 32T networks run and this dump does not.",
        "P2 AHR Turbo leftover: SRS_IC / SRS_JOINT_PC (master Turbo already ON). Separate night from iBeam SRS_TIGHT_MULTIPLEXING.",
        "P2 Massive MIMO multi-layer (Step4): MMIMO_MULTILAYER_ENHANCE + PAIRING_PREFERRED + MU_RANK_BOOSTING — LAYER_16 quota is already set but unused.",
        "P2 AHR Capacity Upgrade 2.0: AHR_CAPC_UPGRADE_PHASE2_SW + PDCCH_MULTI_DIM_JOINT_SCH — loaded-hour capacity wave; Phase1+Turbo already ON.",
        "P3 UL Boosting 1.0 (Step8): UL_LOW_NOISE_SW — UL tput / SRS quality for DL weights. Not the first DL lever.",
        "Already ON (do not chase): SU/MU/PDCCH MU, AHR Phase1+Turbo, CSI FD_RESOURCE/TYPE0/8PORT, LAYER_16, EPF, HEAVY_LOAD_SCH_PRI, TAIL_PKT_SCH_OPT, DL_PMI_SRS_ADAPT, DL_INITIAL_BEAM_SELECT, SSB density adapt, PUSCH beam-domain enhance, DL 256QAM.",
    ], fill_hex=PALE_GOLD)
    ws.row_dimensions[r - 1].height = 160

    r = section(ws, r, INCON_COLS, "D.  Findings table  (dump vs commercial good-DL-tput 32T)")
    r = findings_header(ws, r)
    for rec in FINDINGS:
        r = find_row(ws, r, rec)

    r = section(ws, r, INCON_COLS, "E.  Proposed MML  (11-col)  ·  3–5 loaded 32T trial  ·  mismatch vs commercial")
    r = note_bar(ws, r, INCON_COLS,
                 "Action: Enable = send on this trial wave; Hold = missing but not first night; Skip = already ON or must not send; "
                 "Fix = RF before switches; Check = LST only. Replace {NrDuCellId} / {NrDuCellTrpId}. "
                 "Yellow columns = MML through License. Do not apply on 2T2R. Do not send LAYER_8.")
    r = mml_header(ws, r)
    mml_start = r
    for rec in mml_rows():
        r = mml_row(ws, r, rec)
    ws.auto_filter.ref = f"A{mml_start - 1}:K{r - 1}"

    r = section(ws, r, INCON_COLS, "F.  Mismatch-only list  (Enable + Fix only — copy this to the trial MO)")
    r = note_bar(ws, r, INCON_COLS,
                 "Filtered from section E: Action = Enable or Fix. This is the set of important parameters/switches "
                 "missing vs a commercial 32T network with good DL user throughput.")
    r = mml_header(ws, r)
    sn = 1
    for rec in mml_rows():
        if rec[3] in ("Enable", "Fix"):
            new = (sn,) + rec[1:]
            r = mml_row(ws, r, new)
            sn += 1

    r = section(ws, r, INCON_COLS, "G.  Trial rules / KPI / licenses")
    r = bullets(ws, r, INCON_COLS, [
        "Sites: 3–5 loaded 32T 3-sector macros where AHR Phase1 + Turbo are already ON and tilt ≠ 255. Exclude DHTIAA1 / DHAPT11 / DHTEJ34.",
        "KPI (MAE, same busy hour, trial vs neighbour control): User DL Average Throughput (DU), User UL Average Throughput (DU), DL IBLER, MCS/rank, MU paired UE ratio, N.ChMeas.MIMO.*.Layer, PDCCH blocking, N.SRS, HOSR, drop.",
        "Rollback: FPD deactivation MML on the same sheet sequence (set the enabled bit back to 0 / HighPrecisionBeamSwitch=OFF).",
        "Licenses to confirm (not in CME dump): NR0S00BEAM00 (iBeam 1.0), NR0S00UAHR00 (UL Boosting), NR0S00ACT200 (AHR CU 2.0), NR0S0DLEPU00 / NR0S0ULEPU00 (layers — already implied by LAYER_16).",
        "Do not stack same night: AHR SRS_IC + iBeam SRS_TIGHT_MULTIPLEXING; PMI_WEIGHT_OPT + SRS_WEIGHT_ESTIMATE (enable SRS first); iBeam 2.0/3.0 before 1.0 KPI-green.",
        "4G dump: not used here. NR MIMO missing switches are all 5G MOs. NSA EN-DC / MeNB work stays on the NSA inconsistency report.",
    ], fill_hex=PALE_GREEN)
    ws.row_dimensions[r - 1].height = 120
    return ws


def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"missing {SRC} — build v1 first (python3 tools/build_mimo_workbook.py)")
    print("copy v1 → v2.0")
    shutil.copy2(SRC, OUT)
    wb = load_workbook(OUT)
    print("patch cover...")
    patch_cover(wb)
    print("incon sheet...")
    build_incon_sheet(wb)
    colors = ["1F4E79", "2E75B6", "0D7377", "C00000", "C65911", "548235",
              "7030A0", "1F4E79", "2E75B6", "0D7377", "C00000", "C65911",
              "548235", "7030A0", "C00000"]
    for i, ws in enumerate(wb.worksheets):
        ws.sheet_view.showGridLines = False
        if i < len(colors):
            ws.sheet_properties.tabColor = colors[i]
    print("saving", OUT)
    wb.save(OUT)
    print("ok", os.path.getsize(OUT), "sheets", len(wb.worksheets), wb.sheetnames[-1])
    with zipfile.ZipFile(ZIP_OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(OUT, os.path.basename(OUT))
    print("zip", ZIP_OUT, os.path.getsize(ZIP_OUT))


if __name__ == "__main__":
    main()
