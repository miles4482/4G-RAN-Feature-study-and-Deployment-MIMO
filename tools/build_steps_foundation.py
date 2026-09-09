# Steps 1–6: Basic MIMO, SU-MIMO, MU-MIMO, Multi-layer, Beam Management, AHR.
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from mimo_excel_style import *
from mimo_common import *


def build_step1(wb):
    ws, r = start_step(
        wb, "3. Step1 Basic MIMO",
        "  Step 1 —  Basic functions of MIMO   (MIMO TDD Ch.4, FBFD-010003)",
        "First sheet in the mandatory sequence. NR cells are already multi-antenna; this step enables weight / sensing optimizations. "
        "No license for basic functions. MML Seq follows Ch.4.4.1.2 (activation then deactivation). Doc uses NrDuCellId=0.")
    r = section(ws, r, COLS, "A.  Principal  (Ch.4.1)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p19_0.png"), COLS, "Figure 4-1  Uplink receive diversity  y = W(Hx+N)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p21_0.png"), COLS, "Downlink beamforming  (MIMO TDD Ch.4.1.2)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p21_1.png"), COLS, "Beamforming weight types  (PMI / SRS / open-loop)")
    r = bullets(ws, r, COLS, [
        "Uplink: combine M receive antennas with weight W to raise uplink SINR.",
        "Downlink: PMI-based weights and/or SRS-based weights (FR1). FR2: PMI only.",
        "MIMO performance enhancements (Ch.4.1.3): UL optimization (SRS SINR meas, UL rank fast decrease, PUSCH CE SINR level) and DL optimization (PMI/SRS adapt, SRS weight validity, MetaAAU beam sensing).",
        "MetaAAU beam sensing: NRDUCellBeamAlgo.BeamPerceiveMode = DISTRIBUTED_MODE is a prerequisite of PMI-based and open-loop weight optimization.",
        "Network impact (Table 4-4): MetaAAU sensing fluctuates UE throughput, PRB, IBLER, MCS, CQI, NI/RSSI; PMI-weight opt increases RRC reconfigurations; open-loop weight opt changes MCS/rank/IBLER and can raise scheduled UE count.",
    ])
    ws.row_dimensions[r - 1].height = 110

    r = section(ws, r, COLS, "B.  Prerequisite / exclusive")
    r = impact_table(ws, r, [
        ("1", "Prerequisite", "FR1 TDD", "MetaAAU beam sensing", "NRDUCellBeamAlgo.BeamPerceiveMode=DISTRIBUTED_MODE", "Set before PMI/open-loop weight opt", "Ch.4.3.2.1"),
        ("2", "Prerequisite", "FR1 TDD", "SRS- and PMI-based weight adaptation", "NRDUCellAlgoSwitch.AdaptiveEdgeExpEnhSwitch → DL_PMI_SRS_ADAPT_SW", "Required before SrsNonASFixedWeightType takes effect", "Ch.4.3.2.1"),
        ("3", "Exclusive", "FR1 TDD", "High-speed railway superior experience", "NRDUCell.HighSpeedFlag=HIGH_SPEED", "PUSCH beam-domain enhancement cannot be enabled on 32T/64T", "Ch.4.3.2.2"),
        ("4", "Exclusive", "FR1 TDD", "Hyper Cell", "NRDUCell.NrDuCellNetworkingMode=HYPER_CELL", "If Hyper Cell, BeamPerceiveMode must be CENTRALIZED_MODE (not DISTRIBUTED)", "Ch.4.3.2.2"),
        ("5", "Constraint", "FR1/FR2", "Hardware", "TxRxMode 2T2R/4T4R/8T8R/32T32R/64T64R; port polarization match", "Ch.4.3.3–4.3.4", "Wrong polarization → MIMO gain collapse"),
    ])

    r = section(ws, r, COLS, "C.  Parameters  (Ch.4.4.1.1 — document MML values)")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellTrp", "TxRxMode", "Site-planned", "64T64R (doc example)", "Antenna configuration of the TRP.", "Doc activation starts with 64T64R.")
    r = add_param(ws, r, 2, "NRDUCellAlgoSwitch", "AdaptiveEdgeExpEnhSwitch / DL_PMI_SRS_ADAPT_SW", "See Parameter Reference", "1", "SRS- and PMI-based weight adaptation.", "Parent of SrsNonASFixedWeightType.")
    r = add_param(ws, r, 3, "NRDUCellPdschPrecode", "SrsWeightValidityPeriod", "See Parameter Reference", "MS400 (doc)", "How long an SRS weight stays valid.", "Related to SRS-based DL precoding.", related=True)
    r = add_param(ws, r, 4, "NRDUCellPdsch", "SrsPreSinrJudgeThld", "See Parameter Reference", "-100 (doc)", "SRS pre-SINR judge threshold (dB).", "")
    r = add_param(ws, r, 5, "NRDUCellPdsch", "DlSchOptTimeThld", "See Parameter Reference", "300 (doc)", "DL scheduling optimization time threshold.", "0 disables (deactivation MML).")
    r = add_param(ws, r, 6, "NRDUCellDlSch", "DlAdaptSchTimeThld", "See Parameter Reference", "30 (doc)", "Adaptive DL scheduling time threshold.", "")
    r = add_param(ws, r, 7, "GNodeBParam", "NrBoardPerformanceSw / CHN_MEASURE_CPU_DEC_SW", "See Parameter Reference", "1", "Channel-measurement CPU decrease.", "gNodeB-level.")
    r = add_param(ws, r, 8, "NRDUCellSrs", "SrsAlgoSwitch / SRS_SINR_MEAS_OPT_SW", "See Parameter Reference", "1", "SRS SINR measurement optimization.", "UL enhancement.")
    r = add_param(ws, r, 9, "NRDUCellUlRank", "UlRankAlgoSw / UL_RANK_FAST_DECREASE_SW", "See Parameter Reference", "1", "Fast UL rank decrease on poor channel.", "")
    r = add_param(ws, r, 10, "NRDUCellPusch", "PuschPerformanceSwitch / PUSCH_CE_SINR_LEVEL_ENH_SW", "See Parameter Reference", "1", "PUSCH channel-estimation SINR-level enhancement.", "")
    r = add_param(ws, r, 11, "NRDUCellPdschPrecode", "SrsNonASFixedWeightType", "See Parameter Reference", "PMI_WEIGHT (doc)", "Weight type when SRS is not available.", "Needs DL_PMI_SRS_ADAPT_SW.")
    r = add_param(ws, r, 12, "NRDUCellCsirs", "FR1MaxCellCsirsPortNum", "See Parameter Reference", "8PORT (doc)", "Max CSI-RS ports in the FR1 cell.", "")
    r = add_param(ws, r, 13, "NRDUCellBeamAlgo", "BeamPerceiveMode", "See Parameter Reference", "DISTRIBUTED_MODE", "MetaAAU beam sensing mode.", "CENTRALIZED_MODE if Hyper Cell.")
    r = add_param(ws, r, 14, "NRDUCellBeamAlgo", "WeightAlgoSwitch / SRS_WEIGHT_ESTIMATE_SW", "See Parameter Reference", "1", "SRS-based weight estimation for DL large-packet UEs.", "Related bits PMI_WEIGHT_OPT_SW, OPEN_LOOP_WEIGHT_OPT_SW.", related=True)
    r = add_param(ws, r, 15, "NRDUCellBeamAlgo", "WeightAlgoSwitch / PMI_WEIGHT_OPT_SW", "See Parameter Reference", "1", "PMI-based weight optimization.", "Needs beam sensing DISTRIBUTED_MODE.")
    r = add_param(ws, r, 16, "NRDUCellBeamAlgo", "WeightAlgoSwitch / OPEN_LOOP_WEIGHT_OPT_SW", "See Parameter Reference", "1", "Open-loop weight optimization.", "Needs beam sensing DISTRIBUTED_MODE.")

    r = section(ws, r, COLS, "D.  MML  (Ch.4.4.1.2, document sequence)")
    r = note_bar(ws, r, COLS, "Seq 1.x = activation in FPD order. Seq 1.D.x = deactivation (document prints them after activation). Replace NrDuCellId=0.")
    r = add_mmls(ws, r, [
        (1, "1.1", "Basic", "FR1 TDD", "Activation — TRP",
         "MOD NRDUCELLTRP: NrDuCellTrpId=0, NrDuCellId=0, TxRxMode=64T64R;",
         "Ch.4.4.1.2 sample. Use planned TxRxMode."),
        (2, "1.2", "Basic", "FR1 TDD", "Activation — PMI/SRS adapt",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-1;",
         "Prerequisite for SrsNonASFixedWeightType."),
        (3, "1.3", "Basic", "FR1 TDD", "Activation — SRS weight validity",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0,SrsWeightValidityPeriod=MS400;",
         ""),
        (4, "1.4", "Basic", "FR1 TDD", "Activation — SRS pre-SINR",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, SrsPreSinrJudgeThld=-100;",
         ""),
        (5, "1.5", "Basic", "FR1 TDD", "Activation — DL sch time",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlSchOptTimeThld=300;",
         ""),
        (6, "1.6", "Basic", "FR1 TDD", "Activation — adapt sch time",
         "MOD NRDUCELLDLSCH: NrDuCellId=0, DlAdaptSchTimeThld=30;",
         ""),
        (7, "1.7", "Basic", "FR1 TDD", "Activation — board CPU",
         "MOD GNODEBPARAM: NrBoardPerformanceSw=CHN_MEASURE_CPU_DEC_SW-1;",
         "gNodeB-level."),
        (8, "1.8", "Basic", "FR1 TDD", "Activation — SRS SINR meas",
         "MOD NRDUCELLSRS: NrDuCellId=0, SrsAlgoSwitch=SRS_SINR_MEAS_OPT_SW-1;",
         "UL optimization."),
        (9, "1.9", "Basic", "FR1 TDD", "Activation — UL rank",
         "MOD NRDUCELLULRANK: NrDuCellId=0, UlRankAlgoSw=UL_RANK_FAST_DECREASE_SW-1;",
         ""),
        (10, "1.10", "Basic", "FR1 TDD", "Activation — PUSCH CE",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, PuschPerformanceSwitch=PUSCH_CE_SINR_LEVEL_ENH_SW-1;",
         ""),
        (11, "1.11", "Basic", "FR1 TDD", "Activation — non-AS weight",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, SrsNonASFixedWeightType=PMI_WEIGHT;",
         ""),
        (12, "1.12", "Basic", "FR1 TDD", "Activation — CSI-RS ports",
         "MOD NRDUCELLCSIRS: NrDuCellId=0, FR1MaxCellCsirsPortNum=8PORT;",
         ""),
        (13, "1.13", "Basic", "FR1 TDD", "Activation — beam sensing",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, BeamPerceiveMode=DISTRIBUTED_MODE;",
         "Required before PMI/open-loop weight opt."),
        (14, "1.14", "Basic", "FR1 TDD", "Activation — SRS weight estimate",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-1;",
         ""),
        (15, "1.15", "Basic", "FR1 TDD", "Activation — PMI weight opt",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, WeightAlgoSwitch=PMI_WEIGHT_OPT_SW-1;",
         ""),
        (16, "1.16", "Basic", "FR1 TDD", "Activation — open-loop weight",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, WeightAlgoSwitch=OPEN_LOOP_WEIGHT_OPT_SW-1;",
         ""),
        (17, "1.D.1", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, AdaptiveEdgeExpEnhSwitch=DL_PMI_SRS_ADAPT_SW-0;",
         "Document deactivation sequence."),
        (18, "1.D.2", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlSchOptTimeThld=0;",
         ""),
        (19, "1.D.3", "Basic", "FR1 TDD", "Deactivation",
         "MOD GNODEBPARAM: NrBoardPerformanceSw=CHN_MEASURE_CPU_DEC_SW-0;",
         ""),
        (20, "1.D.4", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLSRS: NrDuCellId=0, SrsAlgoSwitch=SRS_SINR_MEAS_OPT_SW-0;",
         ""),
        (21, "1.D.5", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLULRANK: NrDuCellId=0, UlRankAlgoSw=UL_RANK_FAST_DECREASE_SW-0;",
         ""),
        (22, "1.D.6", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, PuschPerformanceSwitch=PUSCH_CE_SINR_LEVEL_ENH_SW-0;",
         ""),
        (23, "1.D.7", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, WeightAlgoSwitch=SRS_WEIGHT_ESTIMATE_SW-0;",
         ""),
        (24, "1.D.8", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, WeightAlgoSwitch=PMI_WEIGHT_OPT_SW-0;",
         ""),
        (25, "1.D.9", "Basic", "FR1 TDD", "Deactivation",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=0, WeightAlgoSwitch=OPEN_LOOP_WEIGHT_OPT_SW-0;",
         ""),
    ])

    r = section(ws, r, COLS, "E.  Verification / KPI / licenses")
    r = body(ws, r, COLS, "Ch.4.4.2–4.4.3: basic MIMO needs no extra activation verification (cells are multi-antenna by default). Monitor User Uplink Average Throughput (DU) and User Downlink Average Throughput (DU).")
    r = add_kpi_header(ws, r)
    r = add_kpi(ws, r, 1, "User UL Average Throughput (DU)", "MAE KPI (FPD Ch.4.4.3)", "Mbit/s", "Baseline before SU/MU")
    r = add_kpi(ws, r, 2, "User DL Average Throughput (DU)", "MAE KPI (FPD Ch.4.4.3)", "Mbit/s", "Baseline before SU/MU")
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1/FR2 TDD", "FBFD-010003", "MIMO Basic Package", "—", "—", "No license requirements for basic functions (Ch.4.3.1)")
    return ws


def build_step2(wb):
    ws, r = start_step(
        wb, "4. Step2 SU-MIMO",
        "  Step 2 —  SU-MIMO Multiple Layers   (MIMO TDD Ch.5, FOFD-010020)",
        "Sequence: after Step1. License FOFD-010020 must be purchased. 32T/64T also need 2-layer capacity licenses. "
        "Do not enable MU-MIMO yet if you want to observe SU-only layer counters (Ch.5.4.2).")
    r = section(ws, r, COLS, "A.  Principal")
    r = insert_figure(ws, r, fig("mimo_tdd", "p50_0.png"), COLS, "SU-MIMO basics — N layers of one UE (Ch.5.1.1)")
    r = bullets(ws, r, COLS, [
        "If a UE supports N layers, UL or DL peak rate is theoretically N × single-layer.",
        "DL enhancements: weight, scheduling, power, other (Ch.5.1.2).",
        "UL UE SINR optimization is supported only in high-frequency TDD (Ch.2.3) — see Step10.",
        "Max layers must not exceed licensed 2-layer units (NR0S0DLEPU00 DL / NR0S0ULEPU00 UL on FR1).",
    ])
    ws.row_dimensions[r - 1].height = 72

    r = section(ws, r, COLS, "B.  Prerequisite / exclusive")
    r = impact_table(ws, r, [
        ("1", "Prerequisite", "FR1/FR2", "MIMO basic package / Step1", "Multi-antenna cell already on", "Complete Step1", "Ch.5.3.2.1"),
        ("2", "Constraint", "FR1 32T+", "Layer capacity licenses", "NR0S0DLEPU00 / NR0S0ULEPU00", "One unit = 2 layers per cell", "Ch.5.3.1"),
        ("3", "Constraint", "FR2", "mmWave layer capacity", "NR0SMMWDLE00 / NR0SMMWULE00", "per 2 layers per 100 MHz per cell", "Ch.5.3.1"),
        ("4", "Inherit", "FR1/FR2", "Step1 exclusives", "Hyper Cell / HSR as applicable", "Keep consistent with Step1", ""),
    ])

    r = section(ws, r, COLS, "C.  Parameters")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellPusch", "MaxMimoLayerCnt", "See Parameter Reference", "LAYER_2 (doc)", "Max UL SU-MIMO layers.", "Doc activation example.")
    r = add_param(ws, r, 2, "NRDUCellPdsch", "MaxMimoLayerNum", "See Parameter Reference", "LAYER_DEFAULT (doc)", "Max DL SU-MIMO layers.", "LAYER_DEFAULT lets gNodeB follow license/UE cap.")
    r = add_param(ws, r, 3, "NRDUCellPdsch", "DlLinkAdaptEnhancementSw / DL_RANK_ADAPT_SW", "See Parameter Reference", "1", "DL rank adaptation.", "")
    r = add_param(ws, r, 4, "NRDUCellAlgoSwitch", "CompSwitch / INTRA_GNB_DL_JT_SW", "See Parameter Reference", "1 (doc example)", "Intra-gNB DL joint transmission interaction.", "Set 0 to deactivate.")
    r = add_param(ws, r, 5, "NRDUCellDlAmc", "SuMimoPwrCtrlProtectThld", "See Parameter Reference", "100 (doc)", "SU-MIMO power-control protect threshold.", "")
    r = add_param(ws, r, 6, "NRDUCellDmrs", "DlDmrsSwitch / SU_DMRS_OH_ADAPT_DEDUCT_SW", "See Parameter Reference", "1", "Adaptive DMRS overhead deduction for SU.", "")
    r = add_param(ws, r, 7, "NRDUCellPdschPrecode", "DlPrecodeOptOnSrsDtxSw / SRS_PRECODE_OPT_SW", "See Parameter Reference", "1", "Precoding optimization on SRS DTX.", "")

    r = section(ws, r, COLS, "D.  MML  (Ch.5.4.1.2 sequence)")
    r = add_mmls(ws, r, [
        (1, "2.1", "SU", "FR1/FR2", "Activation — UL layers",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, MaxMimoLayerCnt=LAYER_2;",
         "Ch.5.4.1.2."),
        (2, "2.2", "SU", "FR1/FR2", "Activation — DL layers",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, MaxMimoLayerNum=LAYER_DEFAULT;",
         ""),
        (3, "2.3", "SU", "FR1/FR2", "Activation — DL rank adapt",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlLinkAdaptEnhancementSw=DL_RANK_ADAPT_SW-1;",
         ""),
        (4, "2.4", "SU", "FR1/FR2", "Activation — JT interaction",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, CompSwitch=INTRA_GNB_DL_JT_SW-1;",
         "As printed; confirm JT is intended."),
        (5, "2.5", "SU", "FR1/FR2", "Activation — SU power protect",
         "MOD NRDUCELLDLAMC: NrDuCellId=0, SuMimoPwrCtrlProtectThld=100;",
         ""),
        (6, "2.6", "SU", "FR1/FR2", "Activation — DMRS OH",
         "MOD NRDUCELLDMRS: NrDuCellId=0, DlDmrsSwitch=SU_DMRS_OH_ADAPT_DEDUCT_SW-1;",
         ""),
        (7, "2.7", "SU", "FR1/FR2", "Activation — SRS DTX precoding",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, DlPrecodeOptOnSrsDtxSw=SRS_PRECODE_OPT_SW-1;",
         ""),
        (8, "2.D.1", "SU", "FR1/FR2", "Deactivation",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlLinkAdaptEnhancementSw=DL_RANK_ADAPT_SW-0;",
         ""),
        (9, "2.D.2", "SU", "FR1/FR2", "Deactivation",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, CompSwitch=INTRA_GNB_DL_JT_SW-0;",
         ""),
        (10, "2.D.3", "SU", "FR1/FR2", "Deactivation",
         "MOD NRDUCELLDMRS: NrDuCellId=0, DlDmrsSwitch=SU_DMRS_OH_ADAPT_DEDUCT_SW-0;",
         ""),
        (11, "2.D.4", "SU", "FR1/FR2", "Deactivation",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, DlPrecodeOptOnSrsDtxSw=SRS_PRECODE_OPT_SW-0;",
         ""),
    ])

    r = section(ws, r, COLS, "E.  Counters / KPI / licenses")
    r = add_ctr_header(ws, r)
    r = add_ctr(ws, r, 1, "N.ChMeas.MIMO.DL.Transmission.Layer.Max", "Max DL layers on a PRB", "Maximum number of downlink layers on a PRB in a cell", "SU-only observation if MU is OFF (Ch.5.4.2)")
    r = add_ctr(ws, r, 2, "N.ChMeas.MIMO.UL.Trans.Layer.Max", "Max UL layers on a PRB", "Maximum number of uplink layers on a PRB in a cell", "Same rule")
    r = add_kpi_header(ws, r)
    r = add_kpi(ws, r, 1, "User UL Average Throughput (DU)", "MAE", "Mbit/s", "Expect rise vs Step1")
    r = add_kpi(ws, r, 2, "User DL Average Throughput (DU)", "MAE", "Mbit/s", "Expect rise vs Step1")
    r = add_kpi(ws, r, 3, "Uplink MAC Throughput (single UE)", "MAE User Common Monitoring", "Mbit/s", "NSA: Random/STMSI; SA: 5G-Random/5G-STMSI")
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 & FR2 TDD", "FOFD-010020", "SU-MIMO Multiple Layers", "NR0S0PREUM00", "per Cell", "Must be purchased before activation")
    r = add_lic(ws, r, 2, "FR1 32T+", "capacity", "Massive MIMO DL 2-Layers Extended Processing Unit", "NR0S0DLEPU00", "per 2 Layers per Cell", "Max layers ≤ licensed")
    r = add_lic(ws, r, 3, "FR1 32T+", "capacity", "Massive MIMO UL 2-Layers Extended Processing Unit", "NR0S0ULEPU00", "per 2 Layers per Cell", "Max layers ≤ licensed")
    r = add_lic(ws, r, 4, "FR2 TDD", "capacity", "mmWave DL 2-Layers Extended Processing Unit", "NR0SMMWDLE00", "per 2 Layers per 100MHz per Cell", "Step10 also")
    r = add_lic(ws, r, 5, "FR2 TDD", "capacity", "mmWave UL 2-Layers Extended Processing Unit", "NR0SMMWULE00", "per 2 Layers per 100MHz per Cell", "Step10 also")
    return ws


def build_step3(wb):
    ws, r = start_step(
        wb, "5. Step3 MU-MIMO",
        "  Step 3 —  MU-MIMO Basic Pairing  (low-frequency TDD)   (MIMO TDD Ch.6, FOFD-010010)",
        "Sequence: after Step2 SU-MIMO is verified. FR1 8T example uses TxRxMode=8T8R; 32T example uses 32T32R. "
        "Master switch MuMimoSwitch UL+DL. Isolation / correlation thresholds from document MML.")
    r = section(ws, r, COLS, "A.  Principal")
    r = insert_figure(ws, r, fig("mimo_tdd", "p66_0.png"), COLS, "MU-MIMO basics (Ch.6.1.1)")
    r = insert_figure(ws, r, fig("mimo_tdd", "p69_0.png"), COLS, "MU-MIMO spatial multiplexing procedure (Ch.6.1.2)")
    r = bullets(ws, r, COLS, [
        "gNodeB pairs UEs with enough spatial isolation onto the same time-frequency resource.",
        "PDSCH MU enhancements: paired-layer number, scheduling, power, other (Ch.6.1.3).",
        "PUSCH MU enhancements: paired-layer number, scheduling, beam-domain, other (Ch.6.1.4).",
        "Benefits: cell spectral efficiency and average UE throughput in medium/heavy load. Impacts: IBLER/MCS fluctuation; edge access/HO success may dip; PDCCH MU can lower CCE success and slightly raise drop rate (Ch.6.2).",
        "Triggering: MU pairing starts when MuMimoSwitch is ON, UEs report sufficient isolation (DlPmiMuMimoSpaceIsoThld / DlSrsMuMimoSpaceIsoThld / UlMuMimoCorrThld), and layer/SINR gates pass.",
        "Leaving: UE falls back to SU when isolation fails or DlMuBackToSuSeThld spectral-efficiency test trips (doc example 5).",
    ])
    ws.row_dimensions[r - 1].height = 120
    r = formula_box(
        ws, r, COLS,
        "MU pairing isolation (document thresholds) + sample",
        "DL PMI-based MU: isolation metric compared with NRDUCellPdsch.DlPmiMuMimoSpaceIsoThld (doc = 140).\n"
        "DL SRS-based MU: compared with DlSrsMuMimoSpaceIsoThld (doc = 50) and DlMuMimoSrsPreSinrThld (doc = −50).\n"
        "UL MU: UlMuMimoCorrThld (doc = 9) and UlMuMimoSinrThld (doc = −20 dB).\n"
        "Fallback to SU: DlMuBackToSuSeThld (doc = 5).",
        "Assume SRS isolation measured = 80, threshold = 50, SRS pre-SINR = −10 dB vs Thld −50 dB.\n"
        "80 > 50 AND −10 > −50 → both true → UE is a DL SRS-MU candidate.\n"
        "If isolation falls to 40 (< 50) → pairing rejected → SU-MIMO.",
        "Tune isolation UP to pair fewer (cleaner) UEs; DOWN to pair more (capacity) at IBLER risk.")

    r = section(ws, r, COLS, "B.  Prerequisite / exclusive")
    r = impact_table(ws, r, [
        ("1", "Prerequisite", "FR1 TDD", "SU-MIMO Multiple Layers (Step2)", "FOFD-010020 active; MaxMimoLayer*", "Complete Step2 first", "Ch.6.3.2.1"),
        ("2", "Constraint", "FR1 32T+", "Layer capacity", "NR0S0DLEPU00 / NR0S0ULEPU00", "MU layers count against the same 2-layer units", "Ch.6.3.1"),
        ("3", "Exclusive", "FR1 TDD", "See FPD Ch.6.3.2.2", "Mutually exclusive functions table", "Deactivate listed functions before MU", "Do not skip the FPD exclusive list"),
    ])

    r = section(ws, r, COLS, "C.  Parameters  (Ch.6.4.1.1 / MML)")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellAlgoSwitch", "MuMimoSwitch / UL_MU_MIMO_SW & DL_MU_MIMO_SW", "OFF", "1 & 1 (doc)", "Master UL+DL MU-MIMO switch.", "Parent. Set both bits.")
    r = add_param(ws, r, 2, "NRDUCellTrp", "TxRxMode", "Site", "32T32R or 8T8R (doc has both)", "Antenna mode for the MU cell.", "32T sample then 8T sample in Ch.6.4.1.2.")
    r = add_param(ws, r, 3, "NRDUCellPdsch", "MaxMimoLayerNum", "See Parameter Reference", "LAYER_DEFAULT", "DL layer cap during MU.", "Related to SU Step2.", related=True)
    r = add_param(ws, r, 4, "NRDUCellPdcch", "MaxPairLayerNum", "See Parameter Reference", "LAYER_2 (doc)", "Max PDCCH MU paired layers.", "")
    r = add_param(ws, r, 5, "NRDUCellPusch", "MaxMimoLayerCnt", "See Parameter Reference", "LAYER_4 (doc)", "UL layer cap during MU.", "")
    r = add_param(ws, r, 6, "NRDUCellPdsch", "DlPmiMuMimoSpaceIsoThld", "See Parameter Reference", "140 (doc)", "PMI-based DL MU isolation threshold.", "Greater → harder to pair.")
    r = add_param(ws, r, 7, "NRDUCellPdsch", "DlSrsMuMimoSpaceIsoThld", "See Parameter Reference", "50 (doc)", "SRS-based DL MU isolation threshold.", "", related=True)
    r = add_param(ws, r, 8, "NRDUCellDlMimo", "DlMuMimoSrsPreSinrThld", "See Parameter Reference", "-50 (doc)", "SRS pre-SINR gate for DL MU.", "")
    r = add_param(ws, r, 9, "NRDUCellPdsch", "DlMuMimoGroupMode", "See Parameter Reference", "ISOLATION_CORRELATION (doc)", "How MU groups are built.", "")
    r = add_param(ws, r, 10, "NRDUCellDlAmc", "DlMuMimoSirScaleFactor", "See Parameter Reference", "10 (doc)", "DL MU SIR scale.", "")
    r = add_param(ws, r, 11, "NRDUCellDlMimo", "DlMuBackToSuSeThld", "See Parameter Reference", "5 (doc)", "SE threshold to fall back from MU to SU.", "0 used later in multi-layer Step4.")
    r = add_param(ws, r, 12, "NRDUCellDlMimo", "DlMuPmiBeamNumThld", "See Parameter Reference", "3 (doc)", "PMI beam-number threshold for MU.", "")
    r = add_param(ws, r, 13, "NRDUCellDmrs", "DlMuEstRbPolicy", "See Parameter Reference", "PSEUDOORTHOG_DMRS_ADAPT_DEDUCT (doc)", "DL MU DMRS RB policy.", "")
    r = add_param(ws, r, 14, "NRDUCellDlRank", "DlSrsMuMimoRank / DlPmiMuMimoRank", "See Parameter Reference", "RANK_2 / RANK_2 (doc)", "Rank used when pairing.", "")
    r = add_param(ws, r, 15, "NRDUCellPdschPrecode", "PrecodingIntrfSupprValue", "See Parameter Reference", "DEFAULT (doc)", "Interference suppression in precoding.", "")
    r = add_param(ws, r, 16, "NRDUCellDlSch", "DlSchAlgoSwitch / HEAVY_LOAD_SCH_PRI_OPT_SW", "See Parameter Reference", "1", "Heavy-load scheduling priority opt.", "")
    r = add_param(ws, r, 17, "NRDUCellUlMimo", "UlMuMimoCorrThld", "See Parameter Reference", "9 (doc)", "UL MU correlation threshold.", "")
    r = add_param(ws, r, 18, "NRDUCellUlMimo", "UlMuMimoSinrThld", "See Parameter Reference", "-20 (doc)", "UL MU SINR threshold (dB).", "", related=True)

    r = section(ws, r, COLS, "D.  MML  (Ch.6.4.1.2 sequence — 32T then 8T samples)")
    r = add_mmls(ws, r, [
        (1, "3.1", "MU 32T", "FR1 TDD", "Activation — TRP",
         "MOD NRDUCELLTRP: NrDuCellTrpId=0, NrDuCellId=0, TxRxMode=32T32R;",
         "32T sample."),
        (2, "3.2", "MU", "FR1 TDD", "Activation — DL layers",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, MaxMimoLayerNum=LAYER_DEFAULT;",
         ""),
        (3, "3.3", "MU", "FR1 TDD", "Activation — PDCCH pair layers",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, MaxPairLayerNum=LAYER_2;",
         ""),
        (4, "3.4", "MU", "FR1 TDD", "Activation — UL layers",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, MaxMimoLayerCnt=LAYER_4;",
         ""),
        (5, "3.5", "MU 8T", "FR1 TDD", "Activation — 8T TRP (alternate sample)",
         "MOD NRDUCELLTRP: NrDuCellTrpId=0, NrDuCellId=0, TxRxMode=8T8R;",
         "Second sample in the same chapter — pick one TxRxMode."),
        (6, "3.6", "MU", "FR1 TDD", "Activation — master switch",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, MuMimoSwitch=UL_MU_MIMO_SW-1&DL_MU_MIMO_SW-1;",
         "Both UL and DL bits."),
        (7, "3.7", "MU", "FR1 TDD", "Activation — isolation",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlPmiMuMimoSpaceIsoThld=140, DlSrsMuMimoSpaceIsoThld=50;",
         ""),
        (8, "3.8", "MU", "FR1 TDD", "Activation — SRS pre-SINR",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, DlMuMimoSrsPreSinrThld=-50;",
         ""),
        (9, "3.9", "MU", "FR1 TDD", "Activation — group mode",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlMuMimoGroupMode=ISOLATION_CORRELATION;",
         ""),
        (10, "3.10", "MU", "FR1 TDD", "Activation — SIR scale",
         "MOD NRDUCELLDLAMC: NrDuCellId=0, DlMuMimoSirScaleFactor=10;",
         ""),
        (11, "3.11", "MU", "FR1 TDD", "Activation — fallback / PMI beams",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, DlMuBackToSuSeThld=5, DlMuPmiBeamNumThld=3;",
         ""),
        (12, "3.12", "MU", "FR1 TDD", "Activation — DMRS policy",
         "MOD NRDUCELLDMRS: NrDuCellId=0, DlMuEstRbPolicy=PSEUDOORTHOG_DMRS_ADAPT_DEDUCT;",
         ""),
        (13, "3.13", "MU", "FR1 TDD", "Activation — MU rank",
         "MOD NRDUCELLDLRANK: NrDuCellId=0, DlSrsMuMimoRank=RANK_2, DlPmiMuMimoRank=RANK_2;",
         ""),
        (14, "3.14", "MU", "FR1 TDD", "Activation — IS value",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, PrecodingIntrfSupprValue=DEFAULT;",
         ""),
        (15, "3.15", "MU", "FR1 TDD", "Activation — heavy-load sch",
         "MOD NRDUCELLDLSCH: NrDuCellId=0, DlSchAlgoSwitch=HEAVY_LOAD_SCH_PRI_OPT_SW-1;",
         ""),
        (16, "3.16", "MU", "FR1 TDD", "Activation — UL corr/SINR",
         "MOD NRDUCELLULMIMO: NrDuCellId=0, UlMuMimoCorrThld=9, UlMuMimoSinrThld=-20;",
         ""),
        (17, "3.D.1", "MU", "FR1 TDD", "Deactivation — master",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, MuMimoSwitch=UL_MU_MIMO_SW-0&DL_MU_MIMO_SW-0;",
         "Document deactivation."),
    ])

    r = section(ws, r, COLS, "E.  Counters / licenses")
    r = add_ctr_header(ws, r)
    r = add_ctr(ws, r, 1, "N.ChMeas.MIMO.DL.Transmission.Layer.Max", "Max DL layers/PRB", "Includes MU paired layers once MU is ON", "Should rise vs SU-only")
    r = add_ctr(ws, r, 2, "N.ChMeas.MIMO.UL.Trans.Layer.Max", "Max UL layers/PRB", "UL MU pairing", "Watch with IBLER")
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 TDD", "FOFD-010010", "MU-MIMO Basic Pairing", "NR0S00MUMM00", "per Cell", "Must be purchased")
    r = add_lic(ws, r, 2, "FR1 32T+", "capacity", "DL/UL 2-layer processing", "NR0S0DLEPU00 / NR0S0ULEPU00", "per 2 Layers per Cell", "Shared with SU")
    return ws


def build_step4(wb):
    ws, r = start_step(
        wb, "6. Step4 MM Multi-Layer",
        "  Step 4 —  Massive MIMO multi-layer enhancement   (MIMO TDD Ch.7, FR1 only)",
        "Sequence: after Step3 MU-MIMO is stable on 32T/64T. Not supported on FR2. "
        "DL: HighLayerMuMimoSw bits. UL: UlHighLayerMuMimoSwitch bits. Do not skip MU-MIMO.")
    r = section(ws, r, COLS, "A.  Principal")
    r = bullets(ws, r, COLS, [
        "As MU paired-layer count grows, inter-UE interference, scheduling, and weight accuracy become the bottleneck.",
        "DL (Ch.7.1): MMIMO_MULTILAYER_ENHANCE_SW, MU rank boosting, SRS blind IS, hybrid precoding, PDCCH CCE/BWP0/blind-detect opts, PUCCH interference coordination.",
        "UL (Ch.7.2): multilayer demod enhancement, flexible MU pairing, resource/latency-based adaptive UL scheduling, optional AI UL SU SINR predict.",
        "Trigger: HighLayerMuMimoSw ON and MaxMimoLayerNum raised (doc LAYER_8). Leave: turn the same bits to 0.",
    ])
    ws.row_dimensions[r - 1].height = 88

    r = section(ws, r, COLS, "B.  Prerequisite")
    r = impact_table(ws, r, [
        ("1", "Prerequisite", "FR1 TDD", "MU-MIMO Basic Pairing (Step3)", "MuMimoSwitch UL+DL ON", "Complete Step3", "Ch.7.1.3.2.1 / 7.2.3.2.1"),
        ("2", "Constraint", "FR1 only", "Not for FR2", "—", "Use Step10 for mmWave MU", "Ch.2.3"),
        ("3", "Constraint", "32T/64T", "Layer licenses", "NR0S0DLEPU00 / NR0S0ULEPU00", "LAYER_8 needs enough 2-layer units", ""),
    ])

    r = section(ws, r, COLS, "C.  Parameters")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellPdsch", "MaxMimoLayerNum", "LAYER_DEFAULT typical", "LAYER_8 (doc DL)", "Raise DL layer cap for multi-layer MU.", "Needs license headroom.")
    r = add_param(ws, r, 2, "NRDUCellDlMimo", "HighLayerMuMimoSw / MMIMO_MULTILAYER_ENHANCE_SW", "OFF", "1", "Master DL multi-layer enhance.", "Parent of other HighLayer bits.")
    r = add_param(ws, r, 3, "NRDUCellDlMimo", "HighLayerMuMimoSw / MU_RANK_BOOSTING_SW", "OFF", "1", "MU rank boosting.", "Same parent HighLayerMuMimoSw.", related=True)
    r = add_param(ws, r, 4, "NRDUCellDlMimo", "HighLayerMuMimoSw / SRS_BLIND_IS_SW", "OFF", "1", "SRS blind interference suppression.", "Same parent HighLayerMuMimoSw.", related=True)
    r = add_param(ws, r, 5, "NRDUCellSrs", "SrsBlindIsDegree", "See Parameter Reference", "4 (doc)", "Degree of SRS blind IS.", "Related to SRS_BLIND_IS_SW.")
    r = add_param(ws, r, 6, "NRDUCellDlMimo", "DlMuBackToSuSeThld", "5 in Step3", "0 (doc multi-layer)", "Relax SU-fallback so high-layer MU can stay.", "Changed from Step3 value 5.")
    r = add_param(ws, r, 7, "NRDUCellPdschPrecode", "PrecodingAlgoSwitch / DL_HYBRID_PRECODING_SW", "OFF", "1", "DL hybrid precoding.", "")
    r = add_param(ws, r, 8, "NRDUCellPusch", "MaxMimoLayerCnt", "LAYER_4 in Step3", "LAYER_4 (doc UL)", "UL layers for UL multi-layer chapter.", "")
    r = add_param(ws, r, 9, "NRDUCellUlMimo", "UlHighLayerMuMimoSwitch / MULTILAYER_DEMOD_ENH_SW", "OFF", "1", "UL multilayer demodulation enhance.", "")
    r = add_param(ws, r, 10, "NRDUCellUlMimo", "UlHighLayerMuMimoSwitch / MU_MIMO_FLEX_PAIR_SW", "OFF", "1", "Flexible UL MU pairing.", "Same parent UlHighLayerMuMimoSwitch.", related=True)
    r = add_param(ws, r, 11, "NRDUCellAiAlgo", "AiAmcAlgoSwitch / UL_SU_SINR_INTEL_PREDICT_SW", "OFF", "1 (doc optional)", "AI UL SU SINR predict.", "Needs AI model; DSP NRDUCELLAISCHMODEL.")

    r = section(ws, r, COLS, "D.  MML  (Ch.7.1.4.1.2 then Ch.7.2.4.1.2 sequence)")
    r = add_mmls(ws, r, [
        (1, "4.1", "DL ML", "FR1 TDD", "Activation — 8 layers",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, MaxMimoLayerNum=LAYER_8;",
         "Ch.7.1.4.1.2."),
        (2, "4.2", "DL ML", "FR1 TDD", "Activation — master",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-1;",
         ""),
        (3, "4.3", "DL ML", "FR1 TDD", "Activation — small-pkt robust",
         "MOD NRDUCELLDLSCHRES: NrDuCellId=0, SmallPktType1RobustSchPol=LEVEL3;",
         ""),
        (4, "4.4", "DL ML", "FR1 TDD", "Activation — tail MCS",
         "MOD NRDUCELLDLSCH: NrDuCellId=0, DlSchAlgoSwitch=TAIL_PKT_MCS_OPT_SW-1;",
         ""),
        (5, "4.5", "DL ML", "FR1 TDD", "Activation — MU rank boost",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, HighLayerMuMimoSw=MU_RANK_BOOSTING_SW-1;",
         ""),
        (6, "4.6", "DL ML", "FR1 TDD", "Activation — PUCCH coord",
         "MOD NRDUCELLPUCCH: NrDuCellId=0, PucchAlgoSwitch=PUCCH_INTRF_COORD_SW-1;",
         ""),
        (7, "4.7", "DL ML", "FR1 TDD", "Activation — relax SU fallback",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, DlMuBackToSuSeThld=0;",
         "Was 5 in Step3."),
        (8, "4.8", "DL ML", "FR1 TDD", "Activation — SRS blind IS",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, HighLayerMuMimoSw=SRS_BLIND_IS_SW-1;",
         ""),
        (9, "4.9", "DL ML", "FR1 TDD", "Activation — blind IS degree",
         "MOD NRDUCELLSRS: NrDuCellId=0, SrsBlindIsDegree=4;",
         ""),
        (10, "4.10", "DL ML", "FR1 TDD", "Activation — hybrid precoding",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, PrecodingAlgoSwitch=DL_HYBRID_PRECODING_SW-1;",
         ""),
        (11, "4.11", "DL ML", "FR1 TDD", "Activation — CCE opt",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, PdcchAlgoEnhSwitch=CCE_RESOURCE_OPT_SW-1;",
         ""),
        (12, "4.12", "DL ML", "FR1 TDD", "Activation — BWP0 PDSCH res",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, PdcchAlgoEnhSwitch=UE_BWP0_PDSCH_RES_OPT_SW-1;",
         ""),
        (13, "4.13", "DL ML", "FR1 TDD", "Activation — PDCCH blind det",
         "MOD NRDUCELLPDCCH: NrDuCellId=0, PdcchAlgoSwitch=PDCCH_BLIND_DET_ASSIGN_OPT_SW-1;",
         ""),
        (14, "4.14", "UL ML", "FR1 TDD", "Activation — UL layers",
         "MOD NRDUCELLPUSCH: NrDuCellId=0, MaxMimoLayerCnt=LAYER_4;",
         "Ch.7.2.4.1.2."),
        (15, "4.15", "UL ML", "FR1 TDD", "Activation — AI SINR (optional)",
         "MOD NRDUCELLAIALGO: NrDuCellId=0, AiAmcAlgoSwitch=UL_SU_SINR_INTEL_PREDICT_SW-1;",
         "Verify model with DSP NRDUCELLAISCHMODEL."),
        (16, "4.16", "UL ML", "FR1 TDD", "Activation — multilayer demod",
         "MOD NRDUCELLULMIMO: NrDuCellId=0, UlHighLayerMuMimoSwitch=MULTILAYER_DEMOD_ENH_SW-1;",
         ""),
        (17, "4.17", "UL ML", "FR1 TDD", "Activation — flex pair",
         "MOD NRDUCELLULMIMO: NrDuCellId=0, UlHighLayerMuMimoSwitch=MU_MIMO_FLEX_PAIR_SW-1;",
         ""),
        (18, "4.18", "UL ML", "FR1 TDD", "Activation — res-based UL sch",
         "MOD NRDUCELLULSCH: NrDuCellId=0, UlSchAlgoSwitch=RES_BASED_ADAPT_ULSCH_SW-1;",
         ""),
        (19, "4.19", "UL ML", "FR1 TDD", "Activation — latency UL sch",
         "MOD NRDUCELLULSCH: NrDuCellId=0, UlSchAlgoSwitch=LATENCY_BASED_ADAPT_ULSCH_SW-1;",
         ""),
        (20, "4.D.1", "DL ML", "FR1 TDD", "Deactivation — master last among DL bits",
         "MOD NRDUCELLDLMIMO: NrDuCellId=0, HighLayerMuMimoSw=MMIMO_MULTILAYER_ENHANCE_SW-0;",
         "Turn companion bits 0 as in Ch.7.1.4.1.2 deactivation list."),
    ])
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 TDD", "capacity", "DL/UL 2-layer processing", "NR0S0DLEPU00 / NR0S0ULEPU00", "per 2 Layers per Cell", "LAYER_8 consumes more units")
    return ws


def build_step5(wb):
    ws, r = start_step(
        wb, "7. Step5 Beam Management",
        "  Step 5 —  Beam Management (low-frequency TDD)   (Beam Mgmt FPD Ch.4–5)",
        "Sequence: after foundation MIMO (Steps 1–3 at least). Basic beams FBFD-010015 first, then 3D Coverage Pattern FOFD-010100. "
        "Doc MML uses NrDuCellId=1 / NrDuCellTrpId=1. SCENARIO_31/37 do not consume NR0SBSC3DC00.")
    r = section(ws, r, COLS, "A.  Principal")
    r = bullets(ws, r, COLS, [
        "Broadcast beams: SSB/CSI-RS coverage. 32T/64T use coverage scenarios + tilt/azimuth; 2T/4T/8T use RET/RVD mechanical/electrical tilt (Ch.4.1.1).",
        "Control beams: PDCCH initial beam select, PDCCH beam robustness (DTX threshold), SRS beam select opt (Ch.4.1.2).",
        "3D Coverage Pattern: scenario catalogue for traditional AAU, MetaAAU, 8T, plus customized broadcast beams (ADD NRDUCELLTRPCUSTBEAM) (Ch.5).",
        "Trigger: CoverageScenario + optional BeamOptSwitch. Leave: set CoverageScenario=DEFAULT and remove customized beams.",
    ])
    ws.row_dimensions[r - 1].height = 88

    r = section(ws, r, COLS, "B.  Parameters")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellTrpBeam", "CoverageScenario / Tilt / Azimuth", "DEFAULT / 255 / 0 typical", "DEFAULT then SCENARIO_n (doc)", "Broadcast coverage pattern.", "Tilt=255 means product default in several samples.")
    r = add_param(ws, r, 2, "NRDUCell", "SsbPeriod / SsbTimePos", "See Parameter Reference", "MS20 / DEFAULT (doc)", "SSB broadcast timing.", "")
    r = add_param(ws, r, 3, "NRDUCellCsirs", "CsiPeriod / CsirsCellResourceNum", "See Parameter Reference", "SLOT40 / 4_RESOURCE (doc)", "CSI-RS resources for beam management.", "")
    r = add_param(ws, r, 4, "NRDUCellPdcchAlgo", "PdcchPrecodeEnhPolicy", "See Parameter Reference", "ENH_PMI (doc)", "PDCCH precoding policy.", "")
    r = add_param(ws, r, 5, "NRDUCellAlgoSwitch", "BeamOptSwitch / DL_INITIAL_BEAM_SELECT_SW", "OFF", "1", "DL initial beam selection.", "")
    r = add_param(ws, r, 6, "NRDUCellPdcch", "PdcchAlgoSwitch / PDCCH_BEAM_ROBUST_SW", "OFF", "1", "PDCCH beam robustness.", "")
    r = add_param(ws, r, 7, "NRDUCellPdcchAlgo", "PdcchBeamRobDtxThld", "See Parameter Reference", "4 (doc)", "DTX count before robust beam action.", "Used with PDCCH_BEAM_ROBUST_SW.", related=True)
    r = add_param(ws, r, 8, "NRDUCellSrs", "SrsAlgoSwitch / SRS_BEAM_SELECT_OPT_SW", "OFF", "1", "SRS-based beam select optimization.", "")
    r = add_param(ws, r, 9, "NRDUCellBeamAlgo", "WeightAlgoSwitch / SCENARIO_BEAM_OPT_SW", "OFF", "1 with customized scenario", "Optimize weights for coverage scenario.", "3D coverage.")
    r = add_param(ws, r, 10, "NRDUCellTrpCustBeam", "BeamId / Tilt / Azimuth", "none", "Doc BEAM_0..3 sample", "Customized broadcast beams.", "ADD then CoverageScenario=CUSTOMIZED_BEAM_SCENARIO.")

    r = section(ws, r, COLS, "C.  MML  (Ch.4.4.1.2 then Ch.5.4.1.2 sequence)")
    r = add_mmls(ws, r, [
        (1, "5.1", "Basic beam", "FR1 TDD", "Activation — default coverage",
         "MOD NRDUCELLTRPBEAM: NrDuCellTrpId=1, CoverageScenario=DEFAULT, Tilt=255, Azimuth=0;",
         "Ch.4.4.1.2. NrDuCellTrpId=1 in this FPD."),
        (2, "5.2", "Basic beam", "FR1 TDD", "Activation — SSB",
         "MOD NRDUCELL: NrDuCellId=1, DuplexMode=CELL_TDD, SsbPeriod=MS20, SsbTimePos=DEFAULT;",
         ""),
        (3, "5.3", "Basic beam", "FR1 TDD", "Activation — CSI-RS",
         "MOD NRDUCELLCSIRS: NrDuCellId=1, CsiPeriod=SLOT40, CsirsCellResourceNum=4_RESOURCE;",
         ""),
        (4, "5.4", "Basic beam", "FR1 TDD", "Activation — PDCCH precoding",
         "MOD NRDUCELLPDCCHALGO: NrDuCellId=1, PdcchPrecodeEnhPolicy=ENH_PMI;",
         ""),
        (5, "5.5", "Basic beam", "FR1 TDD", "Activation — initial beam select",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=1, BeamOptSwitch=DL_INITIAL_BEAM_SELECT_SW-1;",
         ""),
        (6, "5.6", "Basic beam", "FR1 TDD", "Activation — PDCCH robust",
         "MOD NRDUCELLPDCCH: NrDuCellId=1, PdcchAlgoSwitch=PDCCH_BEAM_ROBUST_SW-1;",
         ""),
        (7, "5.7", "Basic beam", "FR1 TDD", "Activation — DTX thld",
         "MOD NRDUCELLPDCCHALGO: NrDuCellId=1, PdcchBeamRobDtxThld=4;",
         ""),
        (8, "5.8", "Basic beam", "FR1 TDD", "Activation — SRS beam select",
         "MOD NRDUCELLSRS: NrDuCellId=1, SrsAlgoSwitch=SRS_BEAM_SELECT_OPT_SW-1;",
         ""),
        (9, "5.9", "3D cov", "FR1 TDD", "Activation — scenario 1",
         "MOD NRDUCELLTRPBEAM: NrDuCellTrpId=1, CoverageScenario=SCENARIO_1, Tilt=255, Azimuth=0;",
         "Ch.5.4.1.2. Consumes NR0SBSC3DC00."),
        (10, "5.10", "3D cov", "FR1 TDD", "Activation — custom beams",
         "ADD NRDUCELLTRPCUSTBEAM: NrDuCellTrpId=1, BeamId=BEAM_0, Tilt=6, Azimuth=-40;",
         "Repeat BEAM_1 Tilt=6 Azimuth=-22; BEAM_2 -7; BEAM_3 +7."),
        (11, "5.11", "3D cov", "FR1 TDD", "Activation — enable custom scenario",
         "MOD NRDUCELLTRPBEAM: NrDuCellTrpId=1, CoverageScenario=CUSTOMIZED_BEAM_SCENARIO;",
         ""),
        (12, "5.12", "3D cov", "FR1 TDD", "Activation — scenario beam opt",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=1, WeightAlgoSwitch=SCENARIO_BEAM_OPT_SW-1;",
         ""),
        (13, "5.13", "3D cov", "FR1 TDD", "Activation — scenario 3 + RVD",
         "MOD NRDUCELLTRPBEAM: NrDuCellTrpId=1, CoverageScenario=SCENARIO_3, Tilt=255, Azimuth=0;",
         "Optional RVD: MOD RVDSUBUNIT: DEVICENO=3, SUBUNITNO=1, VBEAMWIDTH=200;"),
        (14, "5.14", "3D cov", "FR1 TDD", "Activation — scenario 29 + RET",
         "MOD NRDUCELLTRPBEAM: NrDuCellTrpId=1, CoverageScenario=SCENARIO_29;",
         "Optional RET: MOD RETSUBUNIT: DEVICENO=0, SUBUNITNO=1, TILT=80;"),
        (15, "5.V", "Verify", "FR1 TDD", "Verification",
         "DSP NRDUCELLTRP",
         "Check Actual Beam / coverage scenario. 32T+: LST NRDUCELLTRPSUBBEAM. Custom: LST NRDUCELLTRPCUSTBEAM."),
        (16, "5.D.1", "3D cov", "FR1 TDD", "Deactivation — remove custom beams",
         "RMV NRDUCELLTRPCUSTBEAM: NrDuCellTrpId=1, BeamId=BEAM_0;",
         "Repeat BEAM_1/2/3 then CoverageScenario=DEFAULT."),
        (17, "5.D.2", "3D cov", "FR1 TDD", "Deactivation — scenario opt",
         "MOD NRDUCELLBEAMALGO: NrDuCellId=1, WeightAlgoSwitch=SCENARIO_BEAM_OPT_SW-0;",
         ""),
    ])
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 TDD", "FBFD-010015", "Basic Beam Management", "—", "—", "Basic package")
    r = add_lic(ws, r, 2, "FR1 TDD", "FOFD-010100", "3D Coverage Pattern", "NR0SBSC3DC00", "per Cell", "Not required for SCENARIO_31/37")
    return ws


def build_step6(wb):
    ws, r = start_step(
        wb, "8. Step6 AHR",
        "  Step 6 —  Massive MIMO AHR   (AHR FPD)   Phase1 → Turbo 2.0 → Capacity Upgrade 2.0",
        "Mandatory inner sequence: AHR Introduction (AHR_PHASE1_SW) first, then Experience Turbo 2.0, then Capacity Upgrade 2.0. "
        "Do not enable Capacity Upgrade without Phase1. Doc uses NrDuCellId=0.")
    r = section(ws, r, COLS, "A.  Principal")
    r = insert_figure(ws, r, fig("ahr", "p13_0.png"), COLS, "AHR adaptive CSI-RS multi-beam / multi-stream (Ch.4.1)")
    r = insert_figure(ws, r, fig("ahr", "p27_0.png"), COLS, "AHR Experience Turbo 2.0 — SRS interference coordination (Ch.5.1.1)")
    r = insert_figure(ws, r, fig("ahr", "p47_0.png"), COLS, "AHR Capacity Upgrade 2.0 (Ch.6)")
    r = bullets(ws, r, COLS, [
        "Phase1 (FOFD-051301): adaptive CSI-RS multi-beam adjustment + adaptive multi-stream perception. Switch AHR_PHASE1_SW.",
        "Turbo 2.0 (FOFD-061201): SRS interference coordination + precise AMC (optional AI DL SU MCS). Switch AHR_EXP_TURBO_PHASE2_SW.",
        "Capacity Upgrade 2.0 (FOFD-061202): MU pairing/gathering, multi-dimension joint scheduling, MU IR precoding. Switch AHR_CAPC_UPGRADE_PHASE2_SW.",
        "Trigger: corresponding AhrSwitch bit = 1 after licenses and MU-MIMO (for capacity phase). Leave: set the same bit to 0 (do not leave Phase1 ON if you intended to roll back the whole AHR stack).",
    ])
    ws.row_dimensions[r - 1].height = 100

    r = section(ws, r, COLS, "B.  Parameters")
    r = add_param_header(ws, r)
    r = add_param(ws, r, 1, "NRDUCellFeatureSw", "AhrSwitch / AHR_PHASE1_SW", "OFF", "1", "AHR Introduction master.", "Seq 6.A first.")
    r = add_param(ws, r, 2, "NRDUCellCsirs", "CsiSwitch / CSIRS_INTRF_STATIC_AVOID_SW", "OFF", "1 (doc)", "CSI-RS static interference avoid.", "Phase1 companion.")
    r = add_param(ws, r, 3, "NRDUCellCsirs", "CsirsCellResourceNum / CsirsBeamType", "See Parameter Reference", "FD_RESOURCE / TYPE0 (doc)", "CSI-RS resource/beam type for AHR.", "Phase1 companion.", related=True)
    r = add_param(ws, r, 4, "NRDUCellFeatureSw", "AhrSwitch / AHR_EXP_TURBO_PHASE2_SW", "OFF", "1 after Phase1", "Experience Turbo 2.0 master.", "Seq 6.B.")
    r = add_param(ws, r, 5, "NRDUCellSrsMeas", "SrsIntrfThld", "See Parameter Reference", "5 (doc)", "SRS interference threshold.", "Turbo 2.0.")
    r = add_param(ws, r, 6, "NRDUCellUlPcConfig", "UlPwrCtrlAlgoSwitch / SRS_JOINT_PC_SW", "OFF", "1", "SRS joint power control.", "")
    r = add_param(ws, r, 7, "NRDUCellFeatureSw", "AhrSwitch / AHR_CAPC_UPGRADE_PHASE2_SW", "OFF", "1 after Turbo if capacity is the goal", "Capacity Upgrade 2.0 master.", "Seq 6.C. Needs MU.")
    r = add_param(ws, r, 8, "NRDUCellAlgoSwitch", "MuMimoOptSwith", "See Parameter Reference", "ON (doc spelling)", "MU-MIMO optimization for capacity phase.", "Printed MuMimoOptSwith in FPD.")

    r = section(ws, r, COLS, "C.  MML  (Ch.4 then Ch.5 then Ch.6 sequence)")
    r = add_mmls(ws, r, [
        (1, "6.A.1", "AHR P1", "FR1 TDD", "Activation — CSI-RS avoid",
         "MOD NRDUCELLCSIRS: NrDuCellId=1, CsiSwitch=CSIRS_INTRF_STATIC_AVOID_SW-1;",
         "Ch.4.4.1.2 (NrDuCellId=1 on this command in FPD)."),
        (2, "6.A.2", "AHR P1", "FR1 TDD", "Activation — Phase1 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, AhrSwitch=AHR_PHASE1_SW-1;",
         "Remaining Phase1 commands use NrDuCellId=0."),
        (3, "6.A.3", "AHR P1", "FR1 TDD", "Activation — CSI-RS type",
         "MOD NRDUCELLCSIRS: NrDuCellId=0, CsirsCellResourceNum=FD_RESOURCE, CsirsBeamType=TYPE0;",
         ""),
        (4, "6.A.4", "AHR P1", "FR1 TDD", "Activation — AMC start",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, FixedAmcStepValue=40, DlInitialMcsAdjValue=0;",
         ""),
        (5, "6.A.5", "AHR P1", "FR1 TDD", "Activation — exp-based sch",
         "MOD NRDUCELLSERVEXP: NrDuCellId=0, ServiceExpAlgoSwitch=EXP_BASED_MM_ADAPT_SCH_SW-1;",
         ""),
        (6, "6.A.6", "AHR P1", "FR1 TDD", "Activation — delay buffer",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlDelaySchBufferThld=1;",
         ""),
        (7, "6.A.7", "AHR P1", "FR1 TDD", "Activation — res-based sch",
         "MOD NRDUCELLSERVEXP: NrDuCellId=0, ServiceExpAlgoSwitch=RES_BASED_MM_ADAPT_SCH_SW-1;",
         ""),
        (8, "6.A.8", "AHR P1", "FR1 TDD", "Activation — sch time",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlSchOptTimeThld=300;",
         ""),
        (9, "6.A.9", "AHR P1", "FR1 TDD", "Activation — adapt sch time",
         "MOD NRDUCELLDLSCH: NrDuCellId=0, DlAdaptSchTimeThld=30;",
         ""),
        (10, "6.B.1", "AHR T2", "FR1 TDD", "Activation — Turbo 2.0 master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, AhrSwitch=AHR_EXP_TURBO_PHASE2_SW-1;",
         "Only after 6.A is green."),
        (11, "6.B.2", "AHR T2", "FR1 TDD", "Activation — SRS IF thld",
         "MOD NRDUCELLSRSMEAS: NrDuCellId=0, SrsIntrfThld=5;",
         ""),
        (12, "6.B.3", "AHR T2", "FR1 TDD", "Activation — SRS joint PC",
         "MOD NRDUCELLULPCCONFIG: NrDuCellId=0, UlPwrCtrlAlgoSwitch=SRS_JOINT_PC_SW-1;",
         ""),
        (13, "6.B.4", "AHR T2", "FR1 TDD", "Activation — SRS PC range",
         "MOD NRDUCELLULPCCONFIG: NrDuCellId=0,IntrfUeSrsPcMinSinrTarget=50, MaxSrsPoAdjustAmount=6;",
         ""),
        (14, "6.B.5", "AHR T2", "FR1 TDD", "Activation — precise AMC set",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlPdschAlgoSwitch=TAIL_PKT_SCH_OPT_SW-1;",
         "Also DL_MCS_ADJ_OPT_SW, RANK_AND_SINR_ESTIMATE_OPT_SW, optional DL_SU_MCS_INTEL_OPT_SW."),
        (15, "6.C.1", "AHR CU", "FR1 TDD", "Activation — Capacity Upgrade master",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, AhrSwitch=AHR_CAPC_UPGRADE_PHASE2_SW-1;",
         "Needs MU-MIMO (Step3)."),
        (16, "6.C.2", "AHR CU", "FR1 TDD", "Activation — MU opt",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0, MuMimoOptSwith=ON;",
         "Spelling as in FPD."),
        (17, "6.C.3", "AHR CU", "FR1 TDD", "Activation — gather / pre-SINR",
         "MOD NRDUCELLPDSCH: NrDuCellId=0, DlMuGatherOptSw=ON, DlSrsMuMimoPreSinrThld=0;",
         ""),
        (18, "6.C.4", "AHR CU", "FR1 TDD", "Activation — multi-dim sch",
         "MOD NRDUCELLALGOSWITCH: NrDuCellId=0,MuMimoSwitch=PDCCH_MULTI_DIM_JOINT_SCH_SW-1;",
         ""),
        (19, "6.C.5", "AHR CU", "FR1 TDD", "Activation — MU IR precoding",
         "MOD NRDUCELLPDSCHPRECODE: NrDuCellId=0, DlMuIRPrecodePol=MU_ALLUSER_POWER_ENH;",
         ""),
        (20, "6.D.1", "AHR", "FR1 TDD", "Deactivation example",
         "MOD NRDUCELLFEATURESW: NrDuCellId=0, AhrSwitch=AHR_PHASE1_SW-0;",
         "Turn Turbo/Capacity bits 0 first if they were enabled (FPD deactivation lists)."),
    ])
    r = add_lic_header(ws, r)
    r = add_lic(ws, r, 1, "FR1 TDD", "FOFD-051301", "AHR Introduction", "NR0S00AET100", "per Cell", "Seq 6.A")
    r = add_lic(ws, r, 2, "FR1 TDD", "FOFD-061201", "AHR Experience Turbo 2.0", "NR0S00AET200", "per Cell", "Seq 6.B")
    r = add_lic(ws, r, 3, "FR1 TDD", "FOFD-061202", "AHR Capacity Upgrade 2.0", "NR0S00ACT200", "per Cell", "Seq 6.C")
    return ws
