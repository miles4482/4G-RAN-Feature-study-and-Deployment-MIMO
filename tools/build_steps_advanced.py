# Steps 7–11: iBeam, UL Boosting, DAS+Fusion, mmWave, Cable Sequence Detection.
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from mimo_excel_style import *
from mimo_common import *


def build_step7(wb):
    ws, r = start_step(
        wb, "9. Step7 iBeam",
        "  Step 7 —  Massive MIMO iBeam   (iBeam FPD)   iBeam → 2.0 → 3.0",
        "Mandatory inner sequence: HighPrecisionBeamSwitch (iBeam) then HighPrecisionBeamPhase2Sw then HighPrecisionBeamPhase3Sw. "
        "Do not enable 3.0 without 1.0. Requires Massive MIMO MU (Step3). Doc NrDuCellId=0.")
    r = section(ws, r, COLS, "A.  Principal")
    r = insert_figure(ws, r, fig("ibeam", "p14_0.png"), COLS, "iBeam interference suppression enhancement (Ch.3.1.1)")
    r = insert_figure(ws, r, fig("ibeam", "p15_0.png"), COLS, "iBeam scheduling enhancement (Ch.3.1.2)")
    r = insert_figure(ws, r, fig("ibeam", "p16_0.png"), COLS, "iBeam scheduling / pairing (Ch.3.1.2)")
    r = bullets(ws, r, COLS, [
        "iBeam 1.0 (FOFD-081201): high-precision beams, SRS blind IS measurement, tight SRS multiplexing, interference-random BWP, precise MU sch.",
        "iBeam 2.0 (FOFD-091200): robust weights (SRS_H_BASED_CALCULATION), dynamic cluster grouping, correlation acceleration, optional inter-cell interference avoid.",
        "iBeam 3.0 (FOFD-100200): self-fusion weights, PDCCH robust weights, smart AMC, small-packet shaping, optional AI DL MU MCS.",
        "Trigger: corresponding HighPrecisionBeam* = ON. Leave: set to OFF (FPD lists companion bits to 0).",
    ])
    ws.row_dimensions[r - 1].height = 88

    r = section(ws, r, COLS, "B.  Parameters (masters)")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellFeatureSw", "HighPrecisionBeamSwitch", "OFF", "ON", "iBeam 1.0 master.", "Seq 7.1")
    r = add_param(ws, r, 2, "NRDUCellFeatureSw", "HighPrecisionBeamPhase2Sw", "OFF", "ON after 1.0", "iBeam 2.0 master.", "Seq 7.2")
    r = add_param(ws, r, 3, "NRDUCellFeatureSw", "HighPrecisionBeamPhase3Sw", "OFF", "ON after 2.0", "iBeam 3.0 master.", "Seq 7.3")
    r = add_param(ws, r, 4, "NRDUCellPdschPrecode", "PrecodingAlgoSwitch / DL_ROBUST_WEIGHT_SW", "OFF", "1 (2.0)", "Robust weight.", "RobustWtPhaseCalcMethod=SRS_H_BASED_CALCULATION")
    r = add_param(ws, r, 5, "NRDUCellPdschPrecode", "PrecodingAlgoSwitch / DL_SELF_FUSION_WEIGHT_SW", "OFF", "1 (3.0)", "Self-fusion weight.", "SelfFusionWtSrsPresinrThld=2, SelfFusionWtRhoThld=80")
    r = add_param(ws, r, 6, "NRDUCellPdsch", "DlMuMimoGroupMode", "ISOLATION_CORRELATION in Step3", "DYNAMIC_CLUSTER_GROUP (2.0 doc)", "Dynamic MU clustering for iBeam 2.0.", "Overrides Step3 grouping while 2.0 is ON.")

    r = section(ws, r, COLS, "C.  MML  (Ch.3.4.1.2 → Ch.4.4.1.2 → Ch.5.4.1.2 sequence)")
    r = add_mmls(ws, r, [
        (1, "7.1.1", "iBeam", "FR1 TDD", "Activation — 1.0 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, HighPrecisionBeamSwitch=ON;",
         "Ch.3.4.1.2."),
        (2, "7.1.2", "iBeam", "FR1 TDD", "Activation — SRS blind IS meas",
         "MOD NRDUCELLSRSMEAS: NrDuCellId=0, SrsMeasOptSwitch=SRS_BLIND_IS_MEAS_SW-1;",
         ""),
        (3, "7.1.3", "iBeam", "FR1 TDD", "Activation — tight SRS mux",
         "MOD NRDUCELLSRS: NrDuCellId=0, SrsAlgoSwitch=SRS_TIGHT_MULTIPLEXING_SW-1;",
         ""),
        (4, "7.1.4", "iBeam", "FR1 TDD", "Activation — BWP IF random",
         "MOD NRDUCELLDLSCH: NrDuCellId=0, DlSchAlgoSwitch=DL_BWP_HYBRID_INTRF_RANDOM_SW-1;",
         ""),
        (5, "7.1.5", "iBeam", "FR1 TDD", "Activation — RLC merge sch",
         "MOD NRDUCELLDLSCH: NrDuCellId=0, DlSchAlgoSwitch=DL_RLC_STAT_RPT_MERGE_SCH_SW-1;",
         ""),
        (6, "7.1.6", "iBeam", "FR1 TDD", "Activation — PDCCH agg compress",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, PdcchAlgoSwitch=PDCCH_AGG_LVL_COMPR_SW-1;",
         "AggLvlComprCceUsageThld=60."),
        (7, "7.1.7", "iBeam", "FR1 TDD", "Activation — beam select opt",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, ChannelOptAlgoSwitch=BEAM_SELECT_OPT_SW-1;",
         ""),
        (8, "7.1.8", "iBeam", "FR1 TDD", "Activation — precise + anti-IF MU sch",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, DLMuMimoSchSupplementSw=DL_MU_PRECISE_SCH_SW-1;",
         "Also DL_MU_ANTI_INTRF_SCH_SW-1, TAIL_PKT_MCS_OPT_SW, RES_BASED_DL_ADAPT_SCH_SW."),
        (9, "7.2.1", "iBeam2", "FR1 TDD", "Activation — 2.0 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, HighPrecisionBeamPhase2Sw=ON;",
         "Ch.4.4.1.2. After 7.1 verified."),
        (10, "7.2.2", "iBeam2", "FR1 TDD", "Activation — robust weight",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, PrecodingAlgoSwitch=DL_ROBUST_WEIGHT_SW-1;",
         "RobustWtPhaseCalcMethod=SRS_H_BASED_CALCULATION; SrsIntrfThld=3; SrsBlindIsDegree=7."),
        (11, "7.2.3", "iBeam2", "FR1 TDD", "Activation — dynamic cluster",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlMuMimoGroupMode=DYNAMIC_CLUSTER_GROUP;",
         "MuMimoIblerTarget=15; PRECISE_MUMIMO_EVAL / FAR_UE_RANK_OPT / DL_CORR_ACCELERATION."),
        (12, "7.2.4", "iBeam2", "FR1 TDD", "Activation — optional inter-cell IF avoid",
         "MOD NRDUCELLCOLLABSERV: NrDuCellId=0, MultiCellMimoSwitch=INTER_CELL_INTRF_AVOID_SW-1;",
         "Needs cluster/collab ready."),
        (13, "7.3.1", "iBeam3", "FR1 TDD", "Activation — 3.0 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, HighPrecisionBeamPhase3Sw=ON;",
         "Ch.5.4.1.2."),
        (14, "7.3.2", "iBeam3", "FR1 TDD", "Activation — self-fusion weight",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, PrecodingAlgoSwitch=DL_SELF_FUSION_WEIGHT_SW-1;",
         "SelfFusionWtSrsPresinrThld=2, SelfFusionWtRhoThld=80."),
        (15, "7.3.3", "iBeam3", "FR1 TDD", "Activation — PDCCH robust wt",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, PdcchAlgoSwitch=PDCCH_ROBUST_WEIGHT_SW-1;",
         "PdcchRobWtSrsPresinrThId=-13, PdcchRobWtBFGainCoeff=6."),
        (16, "7.3.4", "iBeam3", "FR1 TDD", "Activation — smart AMC",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, DLMuMimoSchSupplementSw=DL_SMART_AMC_SW-1;",
         "ROBUST_RANK_SW; optional DL_MU_MCS_INTEL_OPT_SW."),
        (17, "7.D.1", "iBeam", "FR1 TDD", "Deactivation — 1.0 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, HighPrecisionBeamSwitch=OFF;",
         "Turn Phase3 then Phase2 then 1.0 companion bits 0 as in FPD."),
    ])
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 TDD", "FOFD-081201", "iBeam", "NR0S00BEAM00", "per Cell", "Seq 7.1")
    r = add_lic(ws, r, 2, "FR1 TDD", "FOFD-091200", "iBeam 2.0", "NR0S0DLEHR00", "per Cell", "Seq 7.2")
    r = add_lic(ws, r, 3, "FR1 TDD", "FOFD-100200", "iBeam 3.0", "NR0S00BEAM30", "per Cell", "Seq 7.3")
    return ws


def build_step8(wb):
    ws, r = start_step(
        wb, "10. Step8 UL Boosting",
        "  Step 8 —  Massive MIMO Uplink Boosting   (UL Boosting FPD)   Phase1 → 2.0",
        "Inner sequence: UL_LOW_NOISE_SW (FOFD-091201) then UL_LOW_NOISE_PHASE2_SW (FOFD-100201). "
        "Phase1 reduces PUSCH/PUCCH/SRS interference; 2.0 adds coordinated PC, precise RB/MCS, multi-beam RX.")
    r = section(ws, r, COLS, "A.  Principal")
    r = insert_figure(ws, r, fig("ulboost", "p15_0.png"), COLS, "UL Boosting — MU pairing optimization (Ch.4.1.1.1)")
    r = insert_figure(ws, r, fig("ulboost", "p21_0.png"), COLS, "UL scheduling optimization (Ch.4.1.1.2)")
    r = insert_figure(ws, r, fig("ulboost", "p23_0.png"), COLS, "Control-channel interference reduction (Ch.4.1.2)")
    r = insert_figure(ws, r, fig("ulboost", "p55_0.png"), COLS, "Uplink Boosting 2.0 — PUSCH coordinated power control (Ch.5.1.1)")
    r = bullets(ws, r, COLS, [
        "Phase1: MU pairing opt (UL_MU_GRP_PAIR_SW, DIFF_WAVEFORM_PAIR_SW, UL_CORR_ACCELERATION_SW), PUSCH resource adapt, OLLA, PDCCH symbol smart alloc, PUCCH/CSI period opt, IRC-based PDP detect.",
        "Phase2: PUSCH_COORD_PWR_CTRL_SW + INTRF_SC_LINK_PERF_PC_SW, UL retrans precise RB, precise MCS (optional AI), adapt-aware / experience scheduling, multi-beam RX, precise freq-offset and channel estimation.",
        "Trigger: MimoFeatureSwitch bits. Leave: set to 0. Phase2 coordinated PC needs inter-gNB time sync — FPD mentions DSP CLKTST.",
    ])
    ws.row_dimensions[r - 1].height = 100

    r = section(ws, r, COLS, "B.  Parameters (masters)")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellFeatureSw", "MimoFeatureSwitch / UL_LOW_NOISE_SW", "OFF", "1", "Uplink Boosting 1.0 master.", "Seq 8.1")
    r = add_param(ws, r, 2, "NRDUCellFeatureSw", "MimoFeatureSwitch / UL_LOW_NOISE_PHASE2_SW", "OFF", "1 after 1.0", "Uplink Boosting 2.0 master.", "Seq 8.2")
    r = add_param(ws, r, 3, "NRDUCellUlMimo", "UlMuMimoAlgoSwitch / UL_MU_GRP_PAIR_SW", "OFF", "1", "UL MU group pairing optimization.", "Phase1")
    r = add_param(ws, r, 4, "NRDUCellUlPcConfig", "UlPwrCtrlAlgoExtSwitch / PUSCH_COORD_PWR_CTRL_SW", "OFF", "1 (2.0)", "Multi-cell PUSCH coordinated PC.", "Needs sync. UlCpcUeSsbRsrpThld=-120 on NRCELLMEASCONFIG.")
    r = add_param(ws, r, 5, "NRDUCellPusch", "SinrThldforWaveformSel", "See Parameter Reference", "32 (doc)", "Waveform selection SINR threshold.", "Phase1")

    r = section(ws, r, COLS, "C.  MML  (Ch.4.4.1.2 then Ch.5.4.1.2 sequence)")
    r = add_mmls(ws, r, [
        (1, "8.1.1", "ULB1", "FR1 TDD", "Activation — Phase1 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, MimoFeatureSwitch=UL_LOW_NOISE_SW-1;",
         "Ch.4.4.1.2."),
        (2, "8.1.2", "ULB1", "FR1 TDD", "Activation — UL MU group pair",
         "MOD NRDUCELLULMIMO: NrDuCellId=0, UlMuMimoAlgoSwitch=UL_MU_GRP_PAIR_SW-1;",
         "Also DIFF_WAVEFORM_PAIR_SW-1, UL_CORR_ACCELERATION_SW-1."),
        (3, "8.1.3", "ULB1", "FR1 TDD", "Activation — PUSCH res adapt",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, PuschAlgoExtSwitch=PUSCH_RES_ADAPT_ALLOC_SW-1;",
         "UlPreallocationSwitch=UL_PREALLOCATION_PERIOD_ADJ_SW-1; PKT_LEN_BASED_SCH_OPT_SW-1."),
        (4, "8.1.4", "ULB1", "FR1 TDD", "Activation — UL AMC",
         "MOD NRDUCELLULAMC: NrDuCellId=0, UlAmcAlgoSw=UL_CELL_OLLA_SW-1;",
         "SMALL_PKT_OL_ADAPT_ADJ_SW; SR_BASED_SCH_MCS_OPT_SW."),
        (5, "8.1.5", "ULB1", "FR1 TDD", "Activation — waveform SINR",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, SinrThldforWaveformSel=32;",
         ""),
        (6, "8.1.6", "ULB1", "FR1 TDD", "Activation — PDCCH symbols",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, PdcchAlgoEnhSwitch=PDCCH_SYM_SMART_ALLOC_SW-1;",
         "CCE_SYMBOL_ALLOC_OPT_SW; PDCCH_BLIND_DET_ASSIGN_OPT_SW."),
        (7, "8.1.7", "ULB1", "FR1 TDD", "Activation — PUCCH/CSI BWP2",
         "MOD NRDUCELLPUCCH: NrDuCellId=0, CsiResoureAlgoSwitch=BWP2_CSI_PERIOD_OPT_SWITCH-1;",
         "BWP2_PUCCH_POS_OPT_SW; F1_ACK_CODE_CHN_INTRF_OPT_SW; IRC_BASED_PDP_DETECT_SW."),
        (8, "8.2.1", "ULB2", "FR1 TDD", "Activation — Phase2 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, MimoFeatureSwitch=UL_LOW_NOISE_PHASE2_SW-1;",
         "Ch.5.4.1.2. After 8.1."),
        (9, "8.2.2", "ULB2", "FR1 TDD", "Activation — coordinated PC",
         "MOD NRDUCELLULPCCONFIG: NrDuCellId=0, UlPwrCtrlAlgoExtSwitch=PUSCH_COORD_PWR_CTRL_SW-1;",
         "MOD NRCELLMEASCONFIG: NrCellId=0, UlCpcUeSsbRsrpThld=-120; UlCpcUeSrsRsrpThld=-98, UlCpcUeIntrfRsrpThld=-110."),
        (10, "8.2.3", "ULB2", "FR1 TDD", "Activation — link-perf PC",
         "MOD NRDUCELLULPCCONFIG: NrDuCellId=0, UlPwrCtrlAlgoExtSwitch=INTRF_SC_LINK_PERF_PC_SW-1;",
         ""),
        (11, "8.2.4", "ULB2", "FR1 TDD", "Activation — precise retrans RB",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, PuschAlgoExtSwitch=UL_RETRANS_PREC_RB_CTRL_SW-1;",
         "UlFirstRetransMinRbPct=20."),
        (12, "8.2.5", "ULB2", "FR1 TDD", "Activation — precise MCS / aware sch",
         "MOD NRDUCELLAIALGO: NrDuCellId=0, AiAmcAlgoSwitch=UL_PRECISE_MCS_OPT_SW-1;",
         "ADAPT_AWARE_SCH_SW; UL_EXP_SCH_OPT_SW; UL_MU_PBMCSSEL_SW."),
        (13, "8.2.6", "ULB2", "FR1 TDD", "Activation — multi-beam RX + CE",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, PuschAlgoExtSwitch=MULTI_BEAM_RX_ENH_SW-1;",
         "INTRF_PREC_FREQ_OFS_EST_SW; PREC_CHANNEL_EST_SW."),
        (14, "8.V", "Verify", "FR1 TDD", "Sync check for Phase2 PC",
         "DSP CLKTST",
         "FPD: view clock/sync monitoring before coordinated PC."),
        (15, "8.D.1", "ULB", "FR1 TDD", "Deactivation — Phase1 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, MimoFeatureSwitch=UL_LOW_NOISE_SW-0;",
         "Turn PHASE2_SW-0 first if enabled."),
    ])
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 TDD", "FOFD-091201", "Uplink Boosting", "NR0S00UAHR00", "per Cell", "Seq 8.1")
    r = add_lic(ws, r, 2, "FR1 TDD", "FOFD-100201", "Uplink Boosting 2.0", "NR0S00UPBT20", "per Cell", "Seq 8.2")
    return ws


def build_step9(wb):
    ws, r = start_step(
        wb, "11. Step9 DAS+Fusion Cell",
        "  Step 9 —  Distributed Antenna Solutions then Fusion Cell   (site architecture, after air-interface MIMO is green)",
        "Sequence: DAS (multi-TRP one cell) and Fusion Cell (Virtual 128T) are site-architecture steps — run after Steps 1–3 (and typically 5). "
        "Do not mix blindly: Fusion Cell uses two 64T TRPs + GNBCLUSTER INTRA_CELL_MIMO. DAS uses master+slave TRP and DmMimoSwitch.")
    r = section(ws, r, COLS, "A.  Distributed Massive MIMO  (DAS FPD Ch.4)")
    r = insert_figure(ws, r, fig("das", "p11_0.png"), COLS, "Distributed Massive MIMO — TRP layout (DAS FPD)")
    r = insert_figure(ws, r, fig("das", "p13_0.png"), COLS, "Distributed MIMO service / TRP selection (DAS FPD)")
    r = bullets(ws, r, COLS, [
        "One NR DU cell, multiple TRPs (master TxRxMode e.g. 4T4R, slave TrpType=SLAVE, NrDuCellId=65535 in the doc sample).",
        "LampSite feature FOFD-050202 (NR0SDMMIMO00); macro FOFD-071211 (NR0S00VMMM00). Also layer-capacity licenses.",
        "Master switch: DmMimoSwitch = DM_MIMO_SERVICE_SWITCH. Optional PDCCH TRP trans, TRP-select weight, UL beam boost, PDCCH beam opt.",
        "Deactivation: DEA NRCELL then DmMimoSwitch service bit 0.",
    ])
    ws.row_dimensions[r - 1].height = 88
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellTrp", "TxRxMode / TrpType", "Site", "Master 4T4R; slave TrpType=SLAVE (doc)", "TRP roles.", "Slave NrDuCellId=65535 in sample.")
    r = add_param(ws, r, 2, "NRDUCellAlgoSwitch", "DmMimoSwitch / DM_MIMO_SERVICE_SWITCH", "OFF", "1", "Distributed MIMO service master.", "")
    r = add_param(ws, r, 3, "NRDUCellAlgoSwitch", "DmMimoSwitch / DM_MIMO_PDCCH_TRP_TRANS_SW", "OFF", "1 optional", "PDCCH TRP transmission.", "Same parent DmMimoSwitch.", related=True)
    r = add_param(ws, r, 4, "NRDUCellAlgoSwitch", "DmMimoSwitch / DM_MIMO_TRP_SEL_WEIGHT_SW", "OFF", "1 optional", "TRP selection weights.", "Same parent DmMimoSwitch.", related=True)

    r = add_mmls(ws, r, [
        (1, "9.D.1", "DAS", "FR1 TDD", "Activation — master TRP",
         "MOD NRDUCELLTRP: NrDuCellTrpId=1, NrDuCellId=0, TxRxMode=4T4R;",
         "Ch.4.4.1.2. Complete coverage MO as planned."),
        (2, "9.D.2", "DAS", "FR1 TDD", "Activation — slave TRP",
         "MOD NRDUCELLTRP: NrDuCellTrpId=2, NrDuCellId=65535, TxRxMode=4T4R, TrpType=SLAVE;",
         "Doc sample."),
        (3, "9.D.3", "DAS", "FR1 TDD", "Activation — service switch",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, DmMimoSwitch=DM_MIMO_SERVICE_SWITCH-1;",
         ""),
        (4, "9.D.4", "DAS", "FR1 TDD", "Activation — optional PDCCH TRP",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, DmMimoSwitch=DM_MIMO_PDCCH_TRP_TRANS_SW-1;",
         "Also DM_MIMO_TRP_SEL_WEIGHT_SW-1."),
        (5, "9.D.V", "DAS", "FR1 TDD", "Verification",
         "DSP NRDUCELL  /  DSP NRDUCELLTRP",
         "Check distributed MIMO status and TRP state."),
        (6, "9.D.D1", "DAS", "FR1 TDD", "Deactivation",
         "DEA NRCELL: NrCellId=0;",
         "Then DM_MIMO_SERVICE_SWITCH-0 (and companion bits 0)."),
    ])

    r = section(ws, r, COLS, "B.  Fusion Cell / Virtual 128T  (Fusion Cell FPD Ch.4, FBFD-091101)")
    r = insert_figure(ws, r, fig("fusion", "p10_0.png"), COLS, "Fusion Cell overview (two TRPs, one cell)")
    r = insert_figure(ws, r, fig("fusion", "p13_0.png"), COLS, "Fusion Cell enhancements")
    r = insert_figure(ws, r, fig("fusion", "p14_0.png"), COLS, "Fusion Cell MU-MIMO SINR enhancement")
    r = bullets(ws, r, COLS, [
        "Fuse two 64T TRPs into one cell (virtual 128T). Sample uses NrCellId/NrDuCellId/ClusterId=90, MCC/MNC 302/220, gNodeBId=1.",
        "ClusterType=INTRA_CELL_MIMO. License consumption is by TRP count (two TRPs → two units).",
        "Enhancements: FusionCellAlgoSwitch MUMIMO_SINR_ENH_SW, SINGLE_TRP_SSB_TRANS_SW; GNODEBALGO ChnCalibPolSw=FUSION_CALIB_INTRF_AVOID_SW.",
        "Service-affecting: ADD cell/TRP/cluster then ACT NRCELL. Rollback: DEA, RMV coverage/TRP/cluster, set NrDuCellNetworkingMode=NORMAL_CELL.",
    ])
    ws.row_dimensions[r - 1].height = 88
    r = add_mmls(ws, r, [
        (1, "9.F.1", "Fusion", "FR1 TDD", "Activation — NR cell",
         'ADD NRCELL: NrCellId=90, CellName="90", CellId=90, FrequencyBand=N78, DuplexMode=CELL_TDD;',
         "Ch.4.4.1.2 sample IDs — replace. Command continues with site RF params in FPD."),
        (2, "9.F.2", "Fusion", "FR1 TDD", "Activation — NR DU cell",
         'ADD NRDUCELL: NrDuCellId=90, NrDuCellName="90", DuplexMode=CELL_TDD, CellId=90;',
         ""),
        (3, "9.F.3", "Fusion", "FR1 TDD", "Activation — TRP 90",
         "ADD NRDUCELLTRP: NrDuCellTrpId=90, TrpType=DEFAULT, NrDuCellId=90, TxRxMode=64T64R;",
         "Then ADD NRDUCELLCOVERAGE: NrDuCellTrpId=90, NrDuCellCoverageId=90, SectorEqmId=90;"),
        (4, "9.F.4", "Fusion", "FR1 TDD", "Activation — TRP 91",
         "ADD NRDUCELLTRP: NrDuCellTrpId=91, TrpType=DEFAULT, NrDuCellId=90, TxRxMode=64T64R;",
         "Then coverage 91."),
        (5, "9.F.5", "Fusion", "FR1 TDD", "Activation — cluster",
         "ADD GNBCLUSTER: ClusterId=90, ClusterType=INTRA_CELL_MIMO;",
         ""),
        (6, "9.F.6", "Fusion", "FR1 TDD", "Activation — cluster cell",
         'ADD GNBMIMOCLUSTERCELL: ClusterId=90, Mcc="302", Mnc="220", gNodeBId=1, CellId=90;',
         "Replace PLMN/gNodeBId."),
        (7, "9.F.7", "Fusion", "FR1 TDD", "Activation — cell on",
         "ACT NRCELL: NrCellId=90;",
         ""),
        (8, "9.F.8", "Fusion", "FR1 TDD", "Activation — MU SINR enh",
         "MOD NRDUCELLMULTITRP: NrDuCellId=90, FusionCellAlgoSwitch=MUMIMO_SINR_ENH_SW-1;",
         ""),
        (9, "9.F.9", "Fusion", "FR1 TDD", "Activation — calib IF avoid",
         "MOD GNODEBALGO: ChnCalibPolSw=FUSION_CALIB_INTRF_AVOID_SW-1;",
         ""),
        (10, "9.F.10", "Fusion", "FR1 TDD", "Activation — single-TRP SSB",
         "MOD NRDUCELLMULTITRP: NrDuCellId=90, FusionCellAlgoSwitch=SINGLE_TRP_SSB_TRANS_SW-1;",
         "Optional enhancement."),
        (11, "9.F.V", "Fusion", "FR1 TDD", "Verification",
         "DSP NRDUCELL  /  DSP NRDUCELLTRP  /  DSP NRCELLCFGCHK  /  DSP GNBMIMOCLUSTERCELL",
         "Check DU cell state, TRP, cfg check, OTA calibration."),
        (12, "9.F.D1", "Fusion", "FR1 TDD", "Deactivation — cell off",
         "DEA NRCELL: NrCellId=90;",
         "Then RMV coverage 91, RMV TRP 91, RMV cluster cell, RMV cluster, MOD NRDUCELL NetworkingMode=NORMAL_CELL."),
    ])
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 LampSite", "FOFD-050202", "Distributed Massive MIMO", "NR0SDMMIMO00", "see FPD", "Indoor DAS")
    r = add_lic(ws, r, 2, "FR1 macro", "FOFD-071211", "Distributed Massive MIMO (NR TDD)", "NR0S00VMMM00", "see FPD", "Macro DAS")
    r = add_lic(ws, r, 3, "FR1 TDD", "FBFD-091101", "Virtual 128T / Fusion Cell", "capacity by TRP count", "per TRP", "Two 64T TRPs → two units (FPD)")
    return ws


def build_step10(wb):
    ws, r = start_step(
        wb, "12. Step10 mmWave",
        "  Step 10 —  mmWave (FR2) beam management + MU-MIMO + multi-beam FDM",
        "Parallel track to FR1 Steps 1–9 — not a substitute. Sequence on FR2: basic mmWave beams → 3D coverage / dense beam / TA / dynamic beam → mmWave MU-MIMO → multi-beam FDM. "
        "Cell deactivate/activate is required around NRDUCELLTRPMMWAVBEAM changes (doc DEA/ACT NrCellId=0..3).")
    r = section(ws, r, COLS, "A.  Principal")
    r = insert_figure(ws, r, fig("mmwave", "p17_0.png"), COLS, "mmWave beam management overview (mmWave FPD Ch.3)")
    r = insert_figure(ws, r, fig("mmwave", "p30_0.png"), COLS, "mmWave 3D coverage pattern (Ch.5)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p163_0.png"), COLS, "mmWave MU-MIMO basics (MIMO TDD Ch.8)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p174_0.png"), COLS, "Multi-beam FDM (MIMO TDD Ch.9)")
    r = bullets(ws, r, COLS, [
        "FR2 uses PMI-based weights only (no SRS-based DL weights).",
        "Basic beams: CoverageScenario=DEFAULT, Tilt=255 on NRDUCELLTRPMMWAVBEAM (FBFD-010015).",
        "3D coverage FOFD-030201 (NR0SMMW3DCP0): e.g. SCENARIO_101 + SSB_MEAS_POS_POLICY_SW. LampSite scenario-based beam, short TA period, FLEXIBLE_DENSE_BEAM_SW, DYNAMIC_BEAM_ALLOC_SW sit in the same feature ID.",
        "mmWave MU-MIMO: MuMimoSwitch=UL_MU_MIMO_SW, UlMuSrsPreSinrThld=0, UlMuIsolationMeasThld=2, MaxMimoLayerCnt=LAYER_4 (MIMO TDD Ch.8.4.1.2).",
        "Multi-beam FDM: BeamMultiplexSwitch=VOL_BASED_BEAM_MULTIPLEX_SW plus PDCCH PUSCH_DTX_AGG_LVL_ADAPT_SW and optional CSIRS_2PORT_COV_ENH_SW.",
    ])
    ws.row_dimensions[r - 1].height = 120

    r = section(ws, r, COLS, "B.  MML  (mmWave FPD then MIMO TDD Ch.8–9 sequence)")
    r = add_mmls(ws, r, [
        (1, "10.1", "mmWave basic", "FR2 TDD", "Activation — deactivate cells",
         "DEA NRCELL: NrCellId=0;",
         "Repeat NrCellId=1,2,3 as in FPD before beam MOD."),
        (2, "10.2", "mmWave basic", "FR2 TDD", "Activation — default beams",
         "MOD NRDUCELLTRPMMWAVBEAM: NrDuCellTrpId=0, CoverageScenario=DEFAULT, Tilt=255;",
         "Repeat TrpId=1,2,3. Then ACT NRCELL 0..3."),
        (3, "10.3", "mmWave 3D", "FR2 TDD", "Activation — scenario 101",
         "MOD NRDUCELLTRPMMWAVBEAM: NrDuCellTrpId=0, CoverageScenario=SCENARIO_101, Tilt=255;",
         "DEA first. Repeat TRPs. Consumes NR0SMMW3DCP0."),
        (4, "10.4", "mmWave 3D", "FR2 TDD", "Activation — SSB meas pos",
         "MOD NRCELLALGOSWITCH: NrCellId=0, MeasPolicySwitch=SSB_MEAS_POS_POLICY_SW-1;",
         "Repeat cells 1–3 then ACT."),
        (5, "10.5", "mmWave dense", "FR2 TDD", "Activation — flexible dense beam",
         "MOD NRDUCELLTRPMMWAVBEAM: NrDuCellTrpId=0, BeamPerformanceSw=FLEXIBLE_DENSE_BEAM_SW-1;",
         "DEA/ACT around this. Repeat TRPs. Optional CsiRsBeamMeasPeriod=SLOT320/640."),
        (6, "10.6", "mmWave TA", "FR2 TDD", "Activation — short TA",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, ShortTaCommandPeriodSW=SHORT_TAC_PERIOD-1;",
         "Ch.7 of mmWave FPD."),
        (7, "10.7", "mmWave dyn", "FR2 TDD", "Activation — dynamic beam alloc",
         "MOD NRDUCELLMOBILEBHALGO: NrDuCellId=0, MobileBackhaulAlgoSw=DYNAMIC_BEAM_ALLOC_SW-1;",
         "DynBeamAllocDelayDiffThld=15."),
        (8, "10.8", "mmWave MU", "FR2 TDD", "Activation — UL MU",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, MuMimoSwitch=UL_MU_MIMO_SW-1;",
         "MIMO TDD Ch.8.4.1.2."),
        (9, "10.9", "mmWave MU", "FR2 TDD", "Activation — isolation",
         "MOD NRDUCELLULMIMO: NrDuCellId=0, UlMuSrsPreSinrThld=0, UlMuIsolationMeasThld=2;",
         ""),
        (10, "10.10", "mmWave MU", "FR2 TDD", "Activation — UL layers",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, MaxMimoLayerCnt=LAYER_4;",
         ""),
        (11, "10.11", "FDM", "FR2 TDD", "Activation — volume-based beam mux",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, BeamMultiplexSwitch=VOL_BASED_BEAM_MULTIPLEX_SW-1;",
         "MIMO TDD Ch.9.4.1.2."),
        (12, "10.12", "FDM", "FR2 TDD", "Activation — PDCCH DTX agg",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, PdcchAlgoEnhSwitch=PUSCH_DTX_AGG_LVL_ADAPT_SW-1;",
         "Optional CSIRS_2PORT_COV_ENH_SW."),
        (13, "10.V", "Verify", "FR2 TDD", "Verification",
         "DSP NRDUCELLTRP",
         "Observe Actual Beam coverage scenario."),
        (14, "10.D.1", "mmWave MU", "FR2 TDD", "Deactivation — UL MU",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, MuMimoSwitch=UL_MU_MIMO_SW-0;",
         ""),
        (15, "10.D.2", "FDM", "FR2 TDD", "Deactivation — beam mux",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, BeamMultiplexSwitch=VOL_BASED_BEAM_MULTIPLEX_SW-0;",
         ""),
    ])
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR2 TDD", "FBFD-010015", "Basic Beam Management", "—", "—", "mmWave Ch.4")
    r = add_lic(ws, r, 2, "FR2 TDD", "FOFD-030201", "mmWave 3D Coverage Pattern", "NR0SMMW3DCP0", "per Cell", "Also TA / dense / dynamic beam chapters")
    r = add_lic(ws, r, 3, "FR2 TDD", "FOFD-010010", "MU-MIMO Basic Pairing", "NR0S00MUMM00", "per Cell", "Plus NR0SMMWULE00 layer capacity")
    return ws


def build_step11(wb):
    ws, r = start_step(
        wb, "13. Step11 Cable Sequence",
        "  Step 11 —  Inter-Cell Cable Sequence Detection   (MIMO TDD Ch.10, FBFD-010025)",
        "Mandatory commissioning sequence (user requirement: must use sequence). "
        "Detects crossed antenna cables between intra-base-station intra-frequency 4T4R NORMAL_CELL cells. "
        "Run after RF is up; run again if MIMO KPI is poor and VSWR/DAS is not the cause. No license.")
    r = section(ws, r, COLS, "A.  Principal  (Ch.10.1)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p180_0.png"), COLS, "Inter-cell cable sequence detection principle (Ch.10.1)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p181_0.png"), COLS, "Detection based on RSSI / correlation between cells")
    r = bullets(ws, r, COLS, [
        "Polarization of physical and logical antenna ports must match (Ch.4.3.4). Crossed inter-cell jumpers are hard to find visually.",
        "The function correlates uplink RSSI between intra-gNodeB intra-frequency 4T4R cells and reports CROSSED vs CORRECT.",
        "Trigger: STR ANTENNAPORTOPTDET with AntPortOptDetPolicy=INTER_CELL_DETECT (entire base station in the doc example).",
        "Leave / stop: STP ANTENNAPORTOPTDET with the same policy.",
        "Inaccurate in indoor DAS, abnormal cell service, or VSWR alarms (Ch.10.3.3).",
    ])
    ws.row_dimensions[r - 1].height = 100

    r = section(ws, r, COLS, "B.  Requirements  (Ch.10.3)")
    r = impact_table(ws, r, [
        ("1", "Prerequisite", "FR1 TDD", "None (Basic O&M package)", "FBFD-010025", "No extra feature must be ON", "Ch.10.3.1 Licenses = None"),
        ("2", "Constraint", "FR1 TDD", "Cell type", "TxRxMode=4T4R; NrDuCellNetworkingMode=NORMAL_CELL", "Not Hyper Cell / not Fusion unless FPD allows — FPD says NORMAL_CELL", "Ch.10.3.3"),
        ("3", "Constraint", "FR1 TDD", "Cell status", "≥2 intra-freq cells activated; no channel shutdown/derating", "NI+noise average of 3 samples in 15 min ≤ −90 dBm", "Else ERROR_EXCESSIVE_INTERFERENCE"),
        ("4", "Constraint", "FR1 TDD", "Do not use as-is", "Indoor DAS coverage; VSWR alarms; abnormal service", "Results may be inaccurate", "Ch.10.3.3"),
    ])

    r = section(ws, r, COLS, "C.  Parameters")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "STR ANTENNAPORTOPTDET", "AntPortOptDetPolicy", "—", "INTER_CELL_DETECT (doc)", "Antenna port optimize detection policy.", "Table 10-1. Only value used for this feature.")

    r = section(ws, r, COLS, "D.  MML  (Ch.10.4.1.2 sequence)")
    r = add_mmls(ws, r, [
        (1, "11.1", "O&M", "FR1 4T4R", "Activation — start detection (entire gNodeB)",
         "STR ANTENNAPORTOPTDET: AntPortOptDetPolicy=INTER_CELL_DETECT;",
         "Document activation example. Confirm ≥2 intra-freq cells ON."),
        (2, "11.2", "O&M", "FR1 4T4R", "Verification — display result",
         "DSP ANTENNAPORTOPTDET",
         "Read Inter-Cell Line Sequence Detection Result, Crossed NR DU Cell ID, Latest Detection Complete Time."),
        (3, "11.3", "O&M", "FR1 4T4R", "Deactivation — stop detection",
         "STP ANTENNAPORTOPTDET: AntPortOptDetPolicy=INTER_CELL_DETECT;",
         "Document deactivation."),
    ])

    r = section(ws, r, COLS, "E.  Result codes  (Table 10-2 — treat as KPI of the sequence test)")
    r = headers(ws, r, ["SN", "Result value", "Meaning", "Action"] + [""] * 6)
    results = [
        ("1", "CROSSED", "Cross-connection with another cell", "Fix jumpers; Crossed NR DU Cell ID names the peer. Re-run Seq 11.1."),
        ("2", "CORRECT", "No cross-connection", "Sequence OK. Proceed with MIMO KPI."),
        ("3", "DETECTING", "Detection in progress", "Wait; then DSP again."),
        ("4", "NOT_DETECTED", "Cell not selected", "Select the cell on MAE or re-run site-wide STR."),
        ("5", "ERROR_SERVICE_INTERRUPT", "Service interrupted during detection", "Restore cell; re-run."),
        ("6", "ERROR_CHANNEL_SHUTDOWN", "Channel shutdown during detection", "Clear shutdown; re-run."),
        ("7", "ERROR_CHANNEL_DERATING", "Channel derating during detection", "Clear fault/derating; re-run."),
        ("8", "ERROR_INSUFFICIENT_DATA", "Not enough samples", "Keep cells up; re-run when traffic/RSSI present."),
        ("9", "ERROR_LOW_CORRELATION", "Channel correlation too low", "Check antenna/DAS; not a simple swap."),
        ("10", "ERROR_RRSI_NO_CHANGE", "RSSI unchanged (FPD spelling RRSI)", "No detection possible; check RF."),
        ("11", "ERROR_NO_DETECT_RESULT", "Only this intra-freq cell was active", "Activate a second intra-freq cell."),
        ("12", "ERROR_EXCESSIVE_INTERFERENCE", "Interference > −90 dBm", "Clean UL NI then re-run."),
        ("13", "ERROR_INVALID_RSSI_DATA", "RSSI near interference", "Accuracy not guaranteed."),
    ]
    for i, rec in enumerate(results):
        vals = list(rec) + [""] * 6
        fh = PALE_RED if rec[1].startswith("CROSSED") or rec[1].startswith("ERROR") else (PALE_GREEN if rec[1] == "CORRECT" else alt_fill(i))
        r = table_row(ws, r, vals, fills=[fh] * 10, height=28)
        merge(ws, r - 1, 4, r - 1, COLS)
    r = note_bar(ws, r, COLS,
                 "If detection succeeds then a later run fails on some cells, Result shows the new error while Crossed NR DU Cell ID and Latest Detection Complete Time keep the last success. "
                 "Cell reset/deactivate/reestablish keeps the previous result; RST APP clears it. MAE path: Inter-Cell Cable Sequence Detection, RAT=gNodeB, export the result file.")
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 TDD", "FBFD-010025", "Basic O&M Package (cable sequence)", "—", "None", "Ch.10.3.1")
    return ws
