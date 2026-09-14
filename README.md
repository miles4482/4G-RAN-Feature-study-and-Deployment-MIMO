# 4G-RAN-Feature-study-and-Deployment-MIMO

Combined **5G MIMO deployment workbook** extracted from the eight Huawei 5G RAN10.1 Feature Parameter Descriptions in this repository (same document-style Excel as the Carrier Aggregation workbook).

## Download (do not use GitHub Raw for `.xlsx`)

- **v10.0 (current):** [MIMO_Deployment_v10.0.zip](MIMO_Deployment_v10.0.zip) → **Download raw file**, then unzip.
  - New sheet 20 **Pairing + DL Tput Plan** — two equal targets: raise `N.ChMeas.MIMO.DL.MuPairing` AND User DL Average Throughput. P0 recover pairing (`SRS_IC_SW-0`) → P1 land the rejected DL-tput levers (`SRS_WEIGHT_ESTIMATE` / `BEAM_TRACKING` with `SUPER_COVERAGE` first) → P2 spend LAYER_16 (`MMIMO_MULTILAYER_ENHANCE` + `PAIRING_PREFERRED`) → P3 hold pairs. A night that raises one KPI by lowering the other is a fail.
- Or [MIMO_Deployment_v10.0.xlsx](MIMO_Deployment_v10.0.xlsx) via **Download raw file**
- v9.0: [MIMO_Deployment_v9.0.zip](MIMO_Deployment_v9.0.zip) (pairing-only lift plan; DL-tput levers parked)
- v8.0: [MIMO_Deployment_v8.0.zip](MIMO_Deployment_v8.0.zip) (pairing rollback pack)
- v7.0: [MIMO_Deployment_v7.0.zip](MIMO_Deployment_v7.0.zip) (pre-requisite columns + Master Findings + Action Plan)
  - Sheet 14 Section 1: every MML row now carries **Pre-requisite switch / parameter (must be ON first)** and **Pre-requisite MML (run BEFORE this line)**, so the enable order is on the row itself.
  - Sheet 16 **Master Findings** — the Ph1 trial of 10–11 Sep 2026.
  - Sheet 17 **Action Plan Ph1-Ph7**.
- v6.0: [MIMO_Deployment_v6.0.zip](MIMO_Deployment_v6.0.zip) (dump status on each suggestion MML, no pre-requisite columns)
- v5.0: [MIMO_Deployment_v5.0.zip](MIMO_Deployment_v5.0.zip) (dump as a separate Section 2)
- v4.0: [MIMO_Deployment_v4.0.zip](MIMO_Deployment_v4.0.zip)
- v3.0: [MIMO_Deployment_v3.0.zip](MIMO_Deployment_v3.0.zip)
- v2.0 (incon report only): [MIMO_Deployment_v2.0.xlsx](MIMO_Deployment_v2.0.xlsx)
- v1 (deployment steps only): [MIMO_Deployment.xlsx](MIMO_Deployment.xlsx)

## Mandatory sequence

Deploy **Step1 → Step11** in order. Later Massive MIMO features list earlier MIMO / MU-MIMO / Beam Management as prerequisites. FR2 mmWave (Step10) is a parallel track, not a substitute for FR1 Steps 1–9. Step11 is **Inter-Cell Cable Sequence Detection** (commissioning).

| Seq | Sheet | Feature |
| --- | --- | --- |
| 1 | Step1 Basic MIMO | FBFD-010003 |
| 2 | Step2 SU-MIMO | FOFD-010020 |
| 3 | Step3 MU-MIMO | FOFD-010010 |
| 4 | Step4 MM Multi-Layer | MIMO TDD Ch.7 (FR1) |
| 5 | Step5 Beam Management | FBFD-010015 / FOFD-010100 |
| 6 | Step6 AHR | Phase1 → Turbo 2.0 → Capacity Upgrade 2.0 |
| 7 | Step7 iBeam | iBeam → 2.0 → 3.0 |
| 8 | Step8 UL Boosting | Phase1 → 2.0 |
| 9 | Step9 DAS + Fusion Cell | FOFD-050202 / 071211 / FBFD-091101 |
| 10 | Step10 mmWave | FR2 beams + MU-MIMO + multi-beam FDM |
| 11 | Step11 Cable Sequence | FBFD-010025 `STR ANTENNAPORTOPTDET` |
| 14 | MIMO Suggest + Incon (v7.0) | Section 1 FPD boxes (MML + **pre-requisite switch / pre-requisite MML** + dump Live/status/Enabled?/Action + Counter monitor / Impact on KPI / short Notes / Jump to Basic) · Section 2 Enable/Fix proposal · Section 3 counters/KPI |
| 15 | CR01 MIMO Exec Pack | 10 gNB × 3 cells CME Proposed vs live + site-specific MML (101/102/103) |
| 16 | Master Findings | Ph1 trial 10–11 Sep 2026: executed vs rejected switches, DL MU pairing per layer for 5/9/10/11/12 Sep, findings F-01…F-08, verdict per switch |
| 17 | Action Plan Ph1-Ph7 | Ph1 complete · Ph2 corrective MML sequence · Ph3…Ph7 with pre-requisites, exit KPI gates and rollback |
| 18 | Pairing Rollback Pack | RB-1 `SRS_IC_SW-0` tonight · RB-2/3 DL MU gates only if pairing stays down · keep-ON list · restore MML |
| 19 | Pairing Lift Plan | Pairing-only variant · P0 recover → P1 multilayer → P2 pairing-preferred → P3 BackToSu=0 → P4 AHR CU |
| 20 | Pairing + DL Tput Plan | Dual target · P0 recover pairing → P1 land rejected tput levers → P2 multilayer pairing lift → P3 hold pairs |

Every MML table uses a **Seq** column (activation `n.1, n.2, …` then deactivation `n.D.*`), matching the CA workbook. **v6.0** puts dump enabled-or-not on each suggestion MML. **v7.0** adds the pre-requisite columns. **v8.0** adds the pairing rollback pack. **v9.0** is the pairing-only lift plan. **v10.0** is the current action plan: raise `N.ChMeas.MIMO.DL.MuPairing` **and** DL user throughput. Every step sheet still ends with **Performance and Monitoring Counter**.

## Why the pre-requisite columns exist

The 10 Sep 2026 trial sent four beam switches in one command:

```
MOD NRDUCELLBEAMALGO: NRDUCELLID=102, WEIGHTALGOSWITCH=SRS_WEIGHT_ESTIMATE_SW-1,
  CHANNELOPTALGOSWITCH=BEAM_SELECT_OPT_SW-1,
  BEAMOPTALGOSWITCH=INTELLIGENT_BEAM_SELECTION_SW-1&BEAM_TRACKING_SW-1;
RETCODE = 2147616329
```

`BEAM_TRACKING_SW` may only be selected while `DlCoverageAlgoSwitch = SUPER_COVERAGE_SW-1` in the matching `NRDUCellChnCovAlgo` MO. That line was never sent, and because `MOD` is atomic all four switches were lost — including `SRS_WEIGHT_ESTIMATE_SW`, the main DL throughput lever. Sheet 14 now states that dependency on the `BEAM_TRACKING_SW` row, and sheet 17 Ph2 sends the seven corrective lines one switch at a time.

## How to regenerate

```bash
python3 tools/build_mimo_workbook.py          # v1 + v2.0 + v3.0 + v4.0 + v6.0 + v7.0 + v8.0 + v9.0 + v10.0
python3 tools/build_incon_report.py           # v2.0 only (needs v1 xlsx)
python3 tools/build_suggestions_sheet.py      # v3.0 only (needs v2.0 xlsx)
python3 tools/build_cr01_sheet.py             # v4.0 only (needs v3.0 xlsx)
python3 tools/build_combined_sheet.py         # v6.0 only (needs v4.0 xlsx)
python3 tools/build_v7_findings.py            # v7.0 only (needs v4.0 xlsx)
python3 tools/build_v8_rollback.py            # v8.0 only (needs v4.0 xlsx)
python3 tools/build_v9_pairing_plan.py        # v9.0 only (needs v4.0 xlsx)
python3 tools/build_v10_dual_plan.py          # v10.0 only (needs v4.0 xlsx)
```
