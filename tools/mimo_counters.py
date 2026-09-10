# Performance and Monitoring Counter tables (FPD Counter Changes / KPI chapters).
# Appended as the last section on every workbook sheet.
from mimo_excel_style import *
from mimo_common import add_kpi_header, add_kpi

SECTION_TITLE = "Performance and Monitoring Counter"

# (counter_id, name, function, use)
# Keep each sheet to the FPD counters that actually move after that step.
COUNTERS = {
    "0. Cover & Index": [
        ("N.ThpVol.DL / N.RLC.ThpTime.DL.Cell", "User DL Average Throughput (DU)",
         "Primary commercial KPI for this workbook. MAE DU cell throughput.",
         "Busy-hour trial vs neighbour control. Must not drop after any MIMO wave."),
        ("N.ThpVol.UL / N.RLC.ThpTime.UL.Cell", "User UL Average Throughput (DU)",
         "UL sister KPI. SRS quality that feeds DL weights shows up here first.",
         "Watch with UL IBLER and N.UL.SRS.PreSINR after UL Boosting / SRS bits."),
        ("N.ChMeas.MIMO.DL.Transmission.Layer.Max", "Max DL layers on a PRB",
         "SU+MU layer cap actually used (not the MaxMimoLayerNum quota).",
         "Already LAYER_16 quota on DHK 32T — this counter shows consumption."),
        ("N.ChMeas.MIMO.DL.Pair.Layer.Avg", "Average DL MU paired layers",
         "MU spatial multiplexing in service.",
         "Should rise only after MU / iBeam precise-sch / multilayer, not after RF-only."),
        ("N.DL.SCH.*.ErrTB.Ibler / N.DL.SCH.*.TB", "DL IBLER by MCS table",
         "Initial BLER of PDSCH TB per modulation.",
         "Must stay in planned band when enabling interference-aware bits."),
        ("N.CCE.Used.Avg / N.CCE.Avail.*", "PDCCH CCE load",
         "Control-channel occupancy. iBeam PDCCH_AGG_LVL_COMPR targets this.",
         "Blocking up = DL tput down even if PDSCH MCS looks fine."),
        ("N.UL.SRS.PreSINR.Index* / N.SRS.NI.Avg", "SRS quality / SRS NI",
         "SRS that drives DL BF weights and UL rank.",
         "Mandatory after SRS_WEIGHT_ESTIMATE / SRS_TIGHT_MULTIPLEXING / SRS_JOINT_PC."),
        ("N.UECntx.AbnormRel / N.RRC.SetupReq.Succ", "Drop / RRC success",
         "Safety net. MIMO waves must not raise abnormal release.",
         "Rollback if drop or HOSR moves against the control cluster."),
    ],
    "1. Overview of MIMO": [
        ("N.ThpVol.DL", "DL throughput volume",
         "Bytes scheduled on PDSCH (FPD MIMO TDD KPI chapter).",
         "Pair with N.RLC.ThpTime.DL.Cell for User DL Average Throughput."),
        ("N.ThpVol.UL", "UL throughput volume",
         "Bytes on PUSCH.",
         "Pair with N.RLC.ThpTime.UL.Cell."),
        ("N.PRB.DL.Used.Avg / N.PRB.DL.Avail.Avg", "DL PRB usage",
         "Load. MIMO gain is largest when PRB is occupied but MCS/rank still headroom.",
         "Trial on loaded 32T, not empty rural cells."),
        ("N.PRB.UL.Used.Avg / N.PRB.UL.Avail.Avg", "UL PRB usage",
         "UL load sister.",
         "Needed to read UL MU pair-PRB counters."),
        ("N.ChMeas.PDSCH.MCS.k", "PDSCH MCS histogram",
         "DL MCS distribution (k = MCS index).",
         "Beam/iBeam waves should shift mass toward higher MCS if IBLER holds."),
        ("N.ChMeas.PUSCH.MCS.k", "PUSCH MCS histogram",
         "UL MCS distribution.",
         "UL Boosting / PUSCH CE SINR-level enhance."),
        ("N.ChMeas.MIMO.DL.Pair.PRB", "DL MU paired PRBs",
         "How much DL bandwidth is MU-scheduled.",
         "Core MU already ON in DHK; iBeam precise/anti-intrf should raise quality not just count."),
        ("N.User.RRCConn.Avg", "Average RRC-connected users",
         "Normalise tput and pairing by user count.",
         "Do not compare raw tput of a 20-user cell vs a 200-user cell."),
    ],
    "2. Type of MIMO Config": [
        ("N.ChMeas.MIMO.DL.Transmission.Layer.Max", "Max DL layers (SU license FOFD-010020)",
         "Confirms layer-capacity licenses are consumed.",
         "Step2 / Step4. Do not send LAYER_8 on DHK (live LAYER_16)."),
        ("N.ChMeas.MIMO.UL.Trans.Layer.Max", "Max UL layers (NR0S0ULEPU00)",
         "UL SU/MU layer consumption.",
         "Step2 / Step3 / Step8."),
        ("N.ChMeas.MIMO.DL.Pair.Layer / .Avg", "DL MU pair layers (FOFD-010010)",
         "MU-MIMO Basic Pairing in service.",
         "Step3. Already ON 32T DHK."),
        ("N.User.OptimalSSBBeam.Avg", "Users on optimal SSB (FBFD-010015)",
         "Broadcast-beam / SSB adapt effectiveness.",
         "Step5. Moves after SSB_BEAM_ADAPT."),
        ("N.ChMeas.CQI.SingleCW.k", "Wideband CQI (AHR CSI)",
         "CSI quality after AHR Phase1/Turbo.",
         "Step6. Already ON DHK; Capacity Upgrade would move pair-layer."),
        ("N.CCE.DL.AllocReq.Num / AggLvl*", "PDCCH agg-level mix (iBeam)",
         "PDCCH_AGG_LVL_COMPR / hybrid interference-random.",
         "Step7. CR01 turns this pack ON on 10 sites."),
        ("N.ChMeas.MIMO.UL.Pair.Layer / .PRB", "UL MU pairing (UL Boosting)",
         "UL_LOW_NOISE_SW child set (group pair / corr accel).",
         "Step8. CR01 master + children ON 10 sites."),
        ("N.ThpVol.DL.LastSlot / N.UL.NI.Avg.PRB*", "DAS/Fusion/UL NI",
         "Architecture / NI — not MIMO switch KPI.",
         "Step9–10 N/A on this n41 32T NSA cluster."),
    ],
    "3. Step1 Basic MIMO": [
        ("N.ThpVol.DL / User DL Average Throughput (DU)", "DL user throughput",
         "MIMO TDD Ch.4.4.3 primary KPI after weight algorithms.",
         "Expect rise after SRS_WEIGHT_ESTIMATE on large-packet UEs (CR01 W1)."),
        ("N.ThpVol.UL / User UL Average Throughput (DU)", "UL user throughput",
         "Ch.4.4.3. PUSCH CE / SRS meas opt.",
         "Moves with PUSCH_CE_SINR_LEVEL_ENH and SRS_SINR_MEAS_OPT (CR01)."),
        ("N.UL.SRS.PreSINR.Index0–5", "SRS PreSINR histogram",
         "Quality of SRS used for DL weights.",
         "Must stay healthy after SRS_WEIGHT_ESTIMATE. Watch N.SRS.NI.Avg."),
        ("N.SRS.NI.Avg", "SRS noise+interference",
         "SRS collision / IC path.",
         "CR01 does not enable SRS_IC (hold vs iBeam tight MUX). NI should not jump."),
        ("N.ChMeas.PDSCH.MCS.k", "DL MCS",
         "Beamforming gain appears as MCS upshift.",
         "Compare same PRB load vs control."),
        ("N.PDSCH.InitTbDl.Rank1–4", "Initial DL rank mix",
         "Rank used at first TX. Weight quality feeds rank.",
         "Rank2+ share should not collapse."),
        ("N.UL.RSSI.Avg", "UL RSSI",
         "Receive-diversity / PUSCH CE baseline.",
         "Safety. Sudden RSSI step = RF not MIMO."),
        ("N.PRB.DL.Used.Avg", "DL PRB usage",
         "Normalise MCS/tput by load.",
         "Same busy hour, trial vs control."),
    ],
    "4. Step2 SU-MIMO": [
        ("N.ChMeas.MIMO.DL.Transmission.Layer.Max", "Max DL layers on a PRB",
         "MIMO TDD Ch.5.4.2. Observe SU-only if MU is OFF.",
         "DHK MU already ON — this includes MU layers. Quota is LAYER_16."),
        ("N.ChMeas.MIMO.UL.Trans.Layer.Max", "Max UL layers on a PRB",
         "UL SU layer consumption.",
         "CR01 UL rank fast-decrease protects BLER; max layer may dip on poor SRS."),
        ("N.PDSCH.InitTbDl.Rank1–4", "DL initial rank",
         "DL_RANK_ADAPT already ON DHK.",
         "Do not downgrade MaxMimoLayerNum."),
        ("N.PUSCH.InitTbUl.Rank.k / N.PUSCH.TbUl.Rank1–2", "UL rank mix",
         "UL_RANK_FAST_DECREASE (CR01) speeds Rank2→1 on bad SRS.",
         "UL IBLER should fall or hold; UL tput should not crash."),
        ("User UL Average Throughput (DU)", "MAE UL user tput",
         "Ch.5.4.2 / 5.4.3.",
         "NSA: MAE User Common Monitoring UEID = Random/STMSI."),
        ("User DL Average Throughput (DU)", "MAE DL user tput",
         "Ch.5.4.3.",
         "NSA: Random/STMSI. SA would be 5G-Random/5G-STMSI."),
        ("Uplink MAC Throughput (single UE)", "Single-UE MAC UL",
         "User Common Monitoring.",
         "Use for SU peak check on a test UE, not cluster mean."),
        ("N.PRB.UL.Used.Avg", "UL PRB usage",
         "Rank mix is load-dependent.",
         "Read with rank counters."),
    ],
    "5. Step3 MU-MIMO": [
        ("N.ChMeas.MIMO.DL.Pair.Layer", "DL MU paired-layer count",
         "How many DL MU layers are paired (Ch.6).",
         "Basic MU already ON 32T. iBeam precise/anti-intrf (CR01) should raise efficient pairing."),
        ("N.ChMeas.MIMO.DL.Pair.Layer.Avg", "Average DL MU pair layers",
         "Smoothed pairing depth.",
         "Watch vs DL IBLER — pairing up + IBLER up = over-aggressive MU."),
        ("N.ChMeas.MIMO.DL.Pair.PRB", "DL MU paired PRBs",
         "MU bandwidth share.",
         "CR01 DL_MU_PRECISE_SCH / ANTI_INTRF."),
        ("N.ChMeas.MIMO.UL.Pair.Layer", "UL MU paired layers",
         "UL MU already ON. CR01 adds group-pair / corr-accel / diff-waveform.",
         "Step8 children on the same CR."),
        ("N.ChMeas.MIMO.UL.Pair.PRB", "UL MU paired PRBs",
         "UL MU bandwidth.",
         "Read with UL IBLER."),
        ("N.DL.SCH.256QAM.ErrTB.Ibler / .TB", "DL 256QAM IBLER",
         "High MCS must survive MU pairing.",
         "DHK DL 256QAM already ON."),
        ("N.UL.SCH.256QAM.ErrTB.Ibler / .TB", "UL 256QAM IBLER",
         "UL 256QAM FIXED already ON DHK.",
         "UL_RANK_FAST_DECREASE should protect this."),
        ("N.ChMeas.PDSCH.MCS.k", "DL MCS under MU",
         "MU isolation / PreSINR −50 already set.",
         "MCS collapse = isolation or interference-sch problem."),
    ],
    "6. Step4 MM Multi-Layer": [
        ("N.ChMeas.MIMO.DL.Pair.Layer.Avg", "Average DL MU layers",
         "Multilayer enhance + pairing-preferred consume LAYER_16 quota.",
         "NOT in CR01 — still missing vs sheet 14 (MIMO-016). Baseline for later wave."),
        ("N.ChMeas.MIMO.DL.Pair.PRB", "DL MU PRBs",
         "Should rise when MMIMO_MULTILAYER_ENHANCE lands (future).",
         "CR01 does not turn HighLayerMuMimoSw."),
        ("N.ChMeas.MIMO.UL.Pair.Layer", "UL multilayer pairing",
         "UL_MMIMO_MULTILAYER_ENHANCE still OFF in CR01.",
         "Keep as control metric."),
        ("N.PDSCH.InitTbDl.Rank3 / Rank4", "High DL rank share",
         "Rank boosting (MU_RANK_BOOSTING) not in CR01.",
         "Do not expect Rank3/4 jump from CR01."),
        ("User DL Average Throughput (DU)", "Loaded-hour DL tput",
         "Capacity KPI of Step4.",
         "CR01 iBeam may still lift average tput without multilayer."),
        ("N.PRB.DL.Used.Avg", "DL load",
         "Multilayer gain is a loaded-hour story.",
         "Compare BH only."),
        ("N.DL.SCH.*.ErrTB.Ibler", "DL IBLER",
         "More layers → more interference if isolation fails.",
         "Gate for any future multilayer CR."),
        ("N.CCE.Used.Avg", "PDCCH load",
         "More paired UEs need more DCI.",
         "iBeam agg-compress (CR01) helps this even before Step4."),
    ],
    "7. Step5 Beam Management": [
        ("N.User.OptimalSSBBeam.Avg", "Users on the optimal SSB beam",
         "Beam Mgmt FPD. SSB_BEAM_ADAPT / vertical coverage (CR01).",
         "Should rise vs control after CR01 SSB bits."),
        ("N.MAC.ThpVol.DL.OptimalSSB", "DL MAC volume on optimal SSB",
         "Traffic that sits on the chosen SSB beam.",
         "Pair with User DL tput."),
        ("N.MAC.ThpVol.UL.OptimalSSB", "UL MAC volume on optimal SSB",
         "UL on the serving SSB.",
         "HO/idle-connected consistency."),
        ("N.User.RRCConn.Avg", "RRC connected users",
         "Normalise OptimalSSB counters.",
         "Per cell."),
        ("User DL Average Throughput (DU)", "DL tput in mobility",
         "BEAM_TRACKING + INTELLIGENT_BEAM_SELECTION (CR01).",
         "Drive-route / mobility cells matter more than rooftop static."),
        ("N.ChMeas.PDSCH.MCS.k", "DL MCS",
         "Connected-mode BF tracking.",
         "MCS should hold on moving UEs."),
        ("N.UECntx.AbnormRel", "Abnormal release",
         "Bad SSB adapt → HO/drop.",
         "Rollback SSB bits first if drop rises."),
        ("N.RRC.ReEst.Att", "RRC re-establishment attempts",
         "Beam mismatch / HO ping-pong.",
         "Watch with HOSR (not in this FPD; use cluster KPI)."),
    ],
    "8. Step6 AHR": [
        ("N.ChMeas.CQI.SingleCW.k", "Wideband CQI histogram",
         "AHR Phase1 CSI multi-beam / multi-stream (already ON DHK).",
         "CQI mass should already be commercial-like. CU 2.0 not in CR01."),
        ("N.ChMeas.MIMO.DL.Pair.Layer", "DL MU pair layers",
         "Capacity Upgrade 2.0 target (still OFF).",
         "CR01 SRS_JOINT_PC is Turbo leftover without SRS_IC."),
        ("N.PRB.DL.Used.Avg / N.PRB.DL.DrbUsed.Avg", "DL PRB / DRB PRB",
         "AHR FPD load KPIs.",
         "Loaded-hour only."),
        ("N.QoS.DL.PktDelayAirInterface.Time / .Num", "DL air-interface delay",
         "Turbo / precise AMC experience.",
         "Should not worsen after SRS_JOINT_PC."),
        ("N.CCE.Used.Avg / N.CCE.DLDCI.Used.Avg", "PDCCH CCE",
         "AHR CU would add multi-dim joint sch (not in CR01).",
         "Baseline."),
        ("N.UL.RSSI.Avg", "UL RSSI",
         "SRS joint PC changes SRS power.",
         "Watch with N.SRS.NI.Avg."),
        ("N.UL.SCH.*.ErrTB.Ibler / .TB", "UL IBLER",
         "SRS_JOINT_PC + SRS quality.",
         "CR01 enables joint PC without SRS_IC — NI must stay flat."),
        ("User DL Average Throughput (DU)", "DL experience KPI",
         "AHR FPD Ch.4–6.",
         "Phase1+Turbo already ON; do not attribute CR01 tput only to AHR."),
    ],
    "9. Step7 iBeam": [
        ("N.CCE.DL.AllocReq.Num", "DL CCE allocation requests",
         "iBeam FPD. PDCCH demand.",
         "PDCCH_AGG_LVL_COMPR (CR01) should serve same traffic with fewer CCE."),
        ("N.CCE.DL.AggLvl1Num … AggLvl16Num", "DL PDCCH aggregation-level mix",
         "Compress high agg-levels when channel allows.",
         "Agg16/8 share should fall if RF is healthy; if it rises, channel got worse."),
        ("N.CCE.Avail.Equivalent / N.CCE.Used.Avg", "Equivalent CCE availability vs used",
         "Blocking proxy.",
         "Blocking up = rollback iBeam PDCCH bits."),
        ("N.ChMeas.MIMO.DL.Pair.Layer.Avg / Pair.PRB", "DL MU pairing under iBeam",
         "DL_MU_PRECISE_SCH + DL_MU_ANTI_INTRF_SCH (CR01).",
         "Pairing quality KPI of iBeam 1.0."),
        ("N.DL.SCH.256QAM.ErrTB.Ibler / .Rbler / .TB", "High-MCS DL BLER",
         "Interference-aware sch + tail-pkt MCS opt (CR01).",
         "TAIL_PKT_MCS_OPT targets tail users — watch low-MCS IBLER too."),
        ("N.UL.SRS.PreSINR.Index* / N.SRS.NI.Avg", "SRS PreSINR / SRS NI",
         "SRS_TIGHT_MULTIPLEXING + SRS_BLIND_IS_MEAS (CR01).",
         "Do not add SRS_IC the same night (sheet 14 rule; CR01 already omits IC)."),
        ("N.ThpVol.DL / N.Thp.DL.Samp.Indexk", "DL tput + sample bins",
         "iBeam FPD primary benefit.",
         "Main success criterion of CR01 on the 10 sites."),
        ("N.QoS.DL.PktDelayAirInterface.* / N.Traffic.DL.RlcFirstPktDelay.Time",
         "DL latency / first-packet delay",
         "RES_BASED_DL_ADAPT_SCH + RLC status merge (CR01).",
         "Tail experience. Should not rise."),
    ],
    "10. Step8 UL Boosting": [
        ("User UL Average Throughput (DU)", "UL user tput",
         "UL Boosting FPD primary KPI. UL_LOW_NOISE_SW (CR01).",
         "Busy hour vs control. Secondary for DL but feeds SRS→DL weights."),
        ("N.ChMeas.MIMO.UL.Pair.Layer / Pair.PRB", "UL MU pairing",
         "UL_MU_GRP_PAIR + DIFF_WAVEFORM_PAIR + UL_CORR_ACCELERATION (CR01).",
         "Pairing up must not explode UL IBLER."),
        ("N.ChMeas.PUSCH.MCS.k", "PUSCH MCS histogram",
         "PUSCH_CE_SINR_LEVEL_ENH (CR01, Step1 leftover on same CR).",
         "MCS upshift with stable UL IBLER = success."),
        ("N.UL.SCH.*.ErrTB.Ibler / .Rbler / .TB", "UL IBLER / RBLER",
         "Interference reduction claim of Phase1.",
         "Gate for UL_LOW_NOISE. Rollback master if IBLER jumps."),
        ("N.PUSCH.TbUl.Rank1 / Rank2", "UL TB rank mix",
         "With UL_RANK_FAST_DECREASE.",
         "Rank2 share may fall slightly; tput should still rise via MCS/IBLER."),
        ("N.CCE.UL.AggLvl* / N.CCE.ULDCI.Used.Avg", "UL DCI / UL CCE",
         "PDCCH symbol/CCE helpers in the FPD child set.",
         "CR01 does not send those PDCCH UL children — only agg-compress + UL boosting master."),
        ("N.UL.RSSI.Avg", "UL RSSI",
         "Low-noise path.",
         "Step change without traffic change = RF/PC issue."),
        ("N.PRB.PUSCH.Used.Avg", "PUSCH PRB used",
         "UL resource efficiency.",
         "Read with UL tput."),
    ],
    "11. Step9 DAS+Fusion Cell": [
        ("N.ThpVol.DL / N.ThpTime.DL.RmvLastSlot", "DL tput (DAS FPD)",
         "Distributed Massive MIMO coverage-fill KPI.",
         "N/A on this n41 single-TRP 32T cluster — do not enable DM_MIMO."),
        ("N.ThpVol.DL.LastSlot", "Last-slot DL volume",
         "DAS FPD counter set.",
         "Keep OFF. Listed so DAS trials have a pack later."),
        ("N.ChMeas.MIMO.DL.Pair.PRB", "DL MU PRB (Fusion Cell FPD)",
         "Virtual 128T MU.",
         "gNBCluster INTRA_CELL_MIMO not deployed."),
        ("N.UL.NI.Avg.PRB0 / PRB272", "UL NI per PRB",
         "Fusion Cell FPD. Multi-TRP NI.",
         "Use only if Fusion is introduced."),
        ("N.UL.PUSCH.RSRP.Index* / SINR.Index* / RSSI.Index*", "PUSCH RSRP/SINR/RSSI bins",
         "Fusion Cell UL quality.",
         "N/A now."),
        ("N.PRB.DL.Used.Avg", "DL PRB usage",
         "Load after combining TRPs.",
         "N/A now."),
        ("N.PowerSaving.SymbolShutdown.Dur", "Symbol shutdown duration",
         "Do not mix Fusion trial with aggressive symbol shutdown.",
         "Keep as exclusion check."),
        ("User DL Average Throughput (DU)", "User DL tput",
         "Architecture KPI if DAS/Fusion is ever built.",
         "Skip on CR01 sites."),
    ],
    "12. Step10 mmWave": [
        ("N.PDSCH.InitTbDl.Rank1–4", "FR2 DL rank mix",
         "mmWave Beam Mgmt / MIMO TDD Ch.8–9.",
         "N/A — live network is n41 FR1 TDD 40 MHz, not FR2."),
        ("N.PUSCH.InitTbUl.Rank1–2", "FR2 UL rank",
         "mmWave UL.",
         "Do not copy VOL_BASED_BEAM_MULTIPLEX onto n41."),
        ("N.ChMeas.PDSCH.MCS.0–28", "FR2 PDSCH MCS",
         "mmWave FPD.",
         "Not present on DHK dump (empty mmWave MOs)."),
        ("N.ChMeas.PUSCH.MCS.0–28", "FR2 PUSCH MCS",
         "mmWave FPD.",
         "N/A."),
        ("N.CCE.Used.Avg / N.CCE.Avail.Avg", "FR2 PDCCH load",
         "Control beams on mmWave.",
         "N/A."),
        ("User DL Average Throughput (DU)", "FR2 user tput",
         "If an FR2 layer is ever added.",
         "Out of CR01 / sheet 14 scope."),
        ("User UL Average Throughput (DU)", "FR2 UL tput",
         "FR2.",
         "Out of scope."),
        ("N.PRB.DL.Used.Avg", "FR2 PRB usage",
         "Load.",
         "Out of scope."),
    ],
    "13. Step11 Cable Sequence": [
        ("DSP ANTENNAPORTOPTDET result", "Inter-cell sequence result (not a PM counter)",
         "Table 10-2: CROSSED vs CORRECT / ERROR_*.",
         "Run on 4T4R NORMAL_CELL. Not used on 32T AAU CR01 sites."),
        ("N.UL.RSSI.Avg", "UL RSSI during detection",
         "FPD: NI+noise 15 min average must be ≤ −90 dBm or ERROR_EXCESSIVE_INTERFERENCE.",
         "Pre-check before STR ANTENNAPORTOPTDET."),
        ("N.UECntx.AbnormRel", "Abnormal release during STR",
         "ERROR_SERVICE_INTERRUPT if service drops mid-test.",
         "Do not run detection in the same window as a MIMO CR."),
        ("User DL Average Throughput (DU)", "DL tput after cable fix",
         "If result was CROSSED, tput/MCS should recover after jumper swap.",
         "32T AAU CR01 sites skip this step (not 4T4R)."),
        ("N.ChMeas.MIMO.DL.Transmission.Layer.Max", "DL layers after CORRECT",
         "Crossed feeders silently destroy MIMO layers.",
         "Commissioning gate for 4T4R, not for this 32T trial."),
        ("N.ChMeas.PUSCH.MCS.k", "UL MCS after CORRECT",
         "Rx-diversity restored.",
         "4T4R only."),
        ("VSWR / RET alarms (FM, not PM)", "VSWR",
         "FPD: results inaccurate with VSWR alarms.",
         "Clear FM before blaming MIMO KPI."),
        ("N.PRB.UL.Used.Avg", "UL load during detect",
         "ERROR_INSUFFICIENT_DATA if no RSSI samples.",
         "Need cells in service."),
    ],
    "14. MIMO Suggest + Incon": [
        ("N.ThpVol.DL / N.RLC.ThpTime.DL.Cell", "User DL Average Throughput (DU)",
         "North-star for Section 1 DL boxes and Section 3 trial.",
         "Same BH, trial vs neighbour 32T control."),
        ("N.ThpVol.UL / N.RLC.ThpTime.UL.Cell", "User UL Average Throughput (DU)",
         "UL boxes and SRS→DL weights.",
         "Must not pay for DL gain with UL collapse."),
        ("N.DL.SCH.*.ErrTB.Ibler / N.DL.SCH.*.TB", "DL IBLER",
         "iBeam / MU extras in Section 3.",
         "Hold within planned band vs control."),
        ("N.UL.SRS.PreSINR.* / N.SRS.NI.Avg", "SRS quality / NI",
         "Weights + tight MUX + joint PC (no SRS_IC).",
         "NI jump = stop."),
        ("N.CCE.DL.AllocReq.Num / AggLvl*", "PDCCH CCE / agg mix",
         "PDCCH_AGG_LVL_COMPR.",
         "Blocking up = fail."),
        ("N.ChMeas.MIMO.DL.Pair.Layer.Avg", "DL MU pair layers",
         "Precise/anti-intrf MU sch.",
         "Quality gated by IBLER."),
        ("N.User.OptimalSSBBeam.Avg", "Optimal SSB users",
         "SSB adapt + tracking.",
         "Drop/HO are rollback."),
        ("N.UECntx.AbnormRel / N.RRC.ReEst.Att", "Drop / re-establish",
         "Safety for combined W1+W3 night.",
         "Any rise vs control → rollback."),
    ],
    "15. MIMO Suggestions from Doc": [
        ("User DL Average Throughput (DU)", "Benefit KPI of every DL box",
         "Sheet 15 is FPD-based, not dump-based. Still measure this after any enable.",
         "CR01 realises a subset of boxes S15-01 / beam / iBeam 1.0 / UL Boosting."),
        ("User UL Average Throughput (DU)", "Benefit KPI of UL boxes",
         "UL Boosting / PUSCH CE / UL rank.",
         "CR01 lands those on 10 sites."),
        ("N.ChMeas.MIMO.DL.Transmission.Layer.Max", "SU layer counter",
         "S15 SU box. Already live LAYER_16 — Skip.",
         "Do not send FPD sample LAYER_8."),
        ("N.ChMeas.MIMO.DL.Pair.Layer.Avg", "MU / multilayer boxes",
         "Multilayer boxes still Hold (not in CR01).",
         "Use as baseline until a later CR."),
        ("N.User.OptimalSSBBeam.Avg", "Beam-management boxes",
         "SSB adapt now in CR01.",
         "Measure on the 10 sites vs sheet 15 remaining Hold."),
        ("N.CCE.Used.Avg / AggLvl*", "iBeam 1.0 boxes",
         "CR01 HighPrecisionBeamSwitch=ON + children.",
         "Main monitoring pack for suggestion family Step7."),
        ("N.ChMeas.CQI.SingleCW.k", "AHR boxes",
         "Phase1+Turbo already ON. Capacity Upgrade still Hold.",
         "Do not enable AHR CU in the CR01 window."),
        ("DSP ANTENNAPORTOPTDET", "Cable-sequence box",
         "Commissioning, not a throughput switch.",
         "Not in CR01 (32T AAU)."),
    ],
    "16. CR01 MIMO Exec Pack": [
        ("User DL Average Throughput (DU)", "CR01 success KPI",
         "10 gNB × 3 cells (101/102/103). Same BH vs neighbour 32T control.",
         "CR01 stacks beam-weight + iBeam 1.0 + UL Boosting in one CME file."),
        ("User UL Average Throughput (DU)", "UL success KPI",
         "UL_LOW_NOISE + UL MU group-pair + PUSCH CE + rank fast decrease.",
         "Fail if UL tput or UL IBLER moves against control."),
        ("DL IBLER (N.DL.SCH.*.ErrTB.Ibler/TB)", "DL residual BLER",
         "iBeam precise/anti-intrf + tail-pkt MCS + hybrid BWP random.",
         "Gate. Rollback HighPrecisionBeamSwitch if IBLER jumps."),
        ("N.CCE.DL.AllocReq.Num / N.CCE.DL.AggLvl*", "PDCCH CCE / agg-level",
         "PDCCH_AGG_LVL_COMPR ON 30/30 in CR01.",
         "Equivalent CCE used should fall or hold."),
        ("N.ChMeas.MIMO.DL.Pair.Layer.Avg / Pair.PRB", "DL MU pairing",
         "DL_MU_PRECISE_SCH + ANTI_INTRF ON 30/30.",
         "Pairing quality. Watch 256QAM IBLER together."),
        ("N.ChMeas.MIMO.UL.Pair.Layer / Pair.PRB", "UL MU pairing",
         "UL_MU_GRP_PAIR + DIFF_WAVEFORM + CORR_ACCEL ON 30/30.",
         "UL Boosting children."),
        ("N.UL.SRS.PreSINR.Index* / N.SRS.NI.Avg", "SRS quality",
         "SRS_WEIGHT_ESTIMATE + SRS_SINR_MEAS_OPT + TIGHT_MUX + BLIND_IS + JOINT_PC.",
         "CR01 omits SRS_IC — keep it OFF this window. NI must stay flat."),
        ("N.User.OptimalSSBBeam.Avg", "SSB adapt",
         "SSB_BEAM_ADAPT + VERTICAL_COV_IMP ON 30/30. Tracking ON.",
         "Drop/HO are the rollback trigger for these bits."),
        ("N.ChMeas.PDSCH.MCS.k / N.PDSCH.InitTbDl.Rank*", "DL MCS / rank",
         "SRS weights + beam tracking.",
         "Expect MCS upshift on large-packet / mobility samples."),
        ("N.UECntx.AbnormRel / N.RRC.ReEst.Att", "Drop / re-establish",
         "Safety for SSB + iBeam in one night.",
         "Any rise vs control → rollback CR01 on that gNB."),
    ],
}

KPI_ROWS = {
    "0. Cover & Index": [
        ("User DL Average Throughput (DU)", "N.ThpVol.DL / N.RLC.ThpTime.DL.Cell (MAE)", "Mbit/s",
         "Workbook north-star. Trial vs control, same BH."),
        ("User UL Average Throughput (DU)", "N.ThpVol.UL / N.RLC.ThpTime.UL.Cell (MAE)", "Mbit/s",
         "UL sister. SRS→DL weight dependency."),
        ("DL IBLER", "Σ N.DL.SCH.*.ErrTB.Ibler / Σ N.DL.SCH.*.TB", "%",
         "Must hold when adding MU/iBeam bits."),
        ("PDCCH CCE utilisation", "N.CCE.Used.Avg / N.CCE.Avail.*", "%",
         "iBeam agg-compress target."),
    ],
    "16. CR01 MIMO Exec Pack": [
        ("User DL Average Throughput (DU)", "MAE DU, BH, 10 CR01 gNB vs neighbour 32T", "Mbit/s",
         "Pass: trial ≥ control trend (no drop). Target: lift vs pre-CR week."),
        ("User UL Average Throughput (DU)", "MAE DU, same window", "Mbit/s",
         "Pass: no UL regression."),
        ("DL IBLER", "Σ ErrTB.Ibler / Σ TB", "%",
         "Pass: within planned band vs control."),
        ("UL IBLER", "Σ N.UL.SCH.*.ErrTB.Ibler / Σ TB", "%",
         "Pass: no jump after UL_LOW_NOISE / rank fast decrease."),
        ("MU pair-layer (DL)", "N.ChMeas.MIMO.DL.Pair.Layer.Avg", "layers",
         "Informational. Quality gated by IBLER."),
        ("SSB optimal-beam users", "N.User.OptimalSSBBeam.Avg", "users",
         "Should rise after SSB adapt. Watch drop."),
    ],
}

NOTES = {
    "0. Cover & Index":
        "Collect on MAE at 15-minute granularity, busy hour, same weekday. "
        "Always keep a neighbour 32T control (not in the trial CR). "
        "Indoor 2T2R DHTIAA1 / DHAPT11 / DHTEJ34 are not trial objects.",
    "12. Step10 mmWave":
        "This cluster is n41 FR1 — mmWave counters will be empty. Section kept so FR2 work has a pack.",
    "11. Step9 DAS+Fusion Cell":
        "DAS/Fusion not deployed. Do not enable those MOs to ‘make counters appear’.",
    "13. Step11 Cable Sequence":
        "32T AAU CR01 sites skip STR ANTENNAPORTOPTDET (4T4R NORMAL_CELL only).",
    "16. CR01 MIMO Exec Pack":
        "Pre-check: 7 days before CR. Post-check: D+1 and D+7 busy hour. "
        "Rollback = CME reverse of the same Proposed bits (set back to live). "
        "Do not add SRS_IC_SW or iBeam 2.0/3.0 in this window.",
}


def _sheet_has_section(ws):
    for row in ws.iter_rows(min_col=1, max_col=1, min_row=1, max_row=ws.max_row or 1):
        v = row[0].value
        if not v:
            continue
        s = str(v).strip()
        if s == SECTION_TITLE:
            return True
        if s.startswith("Section 3.  Performance counter and Monitoring KPI"):
            return True
        if s.startswith("Section 4.  Performance counter and Monitoring KPI"):
            return True
    return False


def add_perf_monitor(ws, r=None, cols=None, sheet_key=None):
    """Append the Performance and Monitoring Counter section. Idempotent."""
    if sheet_key is None:
        sheet_key = ws.title
    if cols is None:
        cols = max(10, min(ws.max_column or 10, 11))
        if cols < 10:
            cols = 10
    if _sheet_has_section(ws):
        return ws.max_row + 1
    if r is None:
        r = (ws.max_row or 1) + 2

    counters = COUNTERS.get(sheet_key) or COUNTERS.get("0. Cover & Index")
    kpis = KPI_ROWS.get(sheet_key)
    note = NOTES.get(sheet_key,
                     "Source: Huawei 5G RAN10.1 FPD Counter Changes / Network Impact chapters for this sheet’s feature. "
                     "Collect on MAE (DU), 15 min, busy hour. Compare trial vs neighbour control.")

    r = blank(ws, r, 12)
    r = section(ws, r, cols, SECTION_TITLE, fill_hex=TEAL, size=14)
    r = note_bar(ws, r, cols, note)

    # Counter table — pad to sheet width
    titles = ["SN", "Counter ID", "Counter Name", "Function / what it measures",
              "Use on this sheet / CR"]
    while len(titles) < min(cols, 10):
        titles.append("")
    if cols > 10:
        titles = titles[:10] + ["Granularity"] + [""] * (cols - 11)
        titles = titles[:cols]
    r = headers(ws, r, titles[:cols] if len(titles) >= cols else titles + [""] * (cols - len(titles)),
                fill_hex=TEAL)

    for i, rec in enumerate(counters, 1):
        cid, name, fn, use = rec
        vals = [i, cid, name, fn, use]
        while len(vals) < cols:
            vals.append("15 min / cell" if len(vals) == 5 and cols > 10 else "")
        vals = vals[:cols]
        fh = alt_fill(i)
        r = table_row(ws, r, vals, fills=[fh] * cols,
                      bolds=[False, True, True] + [False] * (cols - 3),
                      center_cols={1},
                      height=min(56, 28 + max(len(fn), len(use)) // 90 * 10))
        if cols >= 10:
            # merge description across trailing empties when using 10-col layout
            pass

    if kpis:
        r = blank(ws, r, 8)
        r = subsection(ws, r, cols, "MAE KPI (derived from the counters above)", fill_hex=NAVY2)
        r = add_kpi_header(ws, r)
        # add_kpi writes 10 columns; if sheet is 11-col, extra col stays empty
        for i, rec in enumerate(kpis, 1):
            r = add_kpi(ws, r, i, *rec)
            if cols > 10:
                for c in range(11, cols + 1):
                    ws.cell(r - 1, c).fill = fill(alt_fill(i))
                    ws.cell(r - 1, c).border = thin
    return r


def append_counters_to_all_sheets(wb):
    """Walk every worksheet and append the section if missing."""
    for ws in wb.worksheets:
        key = ws.title
        cols = 14 if "Suggest + Incon" in key else (
            11 if key.startswith("14.") or key.startswith("15.") or key.startswith("16.") else 10)
        print("  counters:", key)
        add_perf_monitor(ws, r=None, cols=cols, sheet_key=key)
