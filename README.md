# 4G-RAN-Feature-study-and-Deployment-MIMO

Combined **5G MIMO deployment workbook** extracted from the eight Huawei 5G RAN10.1 Feature Parameter Descriptions in this repository (same document-style Excel as the Carrier Aggregation workbook).

## Download (do not use GitHub Raw for `.xlsx`)

- **v7.0 (current):** [MIMO_Deployment_v7.0.zip](MIMO_Deployment_v7.0.zip) → **Download raw file**, then unzip.
  - Sheet 14 Section 1: every MML row now carries **Pre-requisite switch / parameter (must be ON first)** and **Pre-requisite MML (run BEFORE this line)**, so the enable order is on the row itself.
  - New sheet 16 **Master Findings** — the Ph1 trial of 10–11 Sep 2026: 18 switches accepted, 4 rejected, DL MU pairing evidence per layer, 8 findings with root cause, and a keep / roll-back / re-run verdict per switch.
  - New sheet 17 **Action Plan Ph1-Ph7** — Ph1 complete, Ph2 corrective sequence as 7 numbered single-switch MML lines, then Ph3…Ph7 with their own pre-requisites and KPI exit gates.
- Or [MIMO_Deployment_v7.0.xlsx](MIMO_Deployment_v7.0.xlsx) via **Download raw file**
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

Every MML table uses a **Seq** column (activation `n.1, n.2, …` then deactivation `n.D.*`), matching the CA workbook. **v6.0** puts dump enabled-or-not on each suggestion MML. **v7.0** adds the pre-requisite columns, so a dependent switch is never sent before the switch it depends on. Every step sheet still ends with **Performance and Monitoring Counter**.

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
python3 tools/build_mimo_workbook.py          # v1 + v2.0 + v3.0 + v4.0 + v6.0 + v7.0
python3 tools/build_incon_report.py           # v2.0 only (needs v1 xlsx)
python3 tools/build_suggestions_sheet.py      # v3.0 only (needs v2.0 xlsx)
python3 tools/build_cr01_sheet.py             # v4.0 only (needs v3.0 xlsx)
python3 tools/build_combined_sheet.py         # v6.0 only (needs v4.0 xlsx)
python3 tools/build_v7_findings.py            # v7.0 only (needs v4.0 xlsx)
```
