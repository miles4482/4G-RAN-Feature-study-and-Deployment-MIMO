# 4G-RAN-Feature-study-and-Deployment-MIMO

Combined **5G MIMO deployment workbook** extracted from the eight Huawei 5G RAN10.1 Feature Parameter Descriptions in this repository (same document-style Excel as the Carrier Aggregation workbook).

## Download (do not use GitHub Raw for `.xlsx`)

- **v5.0 (current):** [MIMO_Deployment_v5.0.zip](MIMO_Deployment_v5.0.zip) → **Download raw file**, then unzip.
  - Sheet 14 combines Suggestions + Incon (4 sections). MML columns: Counter monitor, Impact on KPI, short Notes; Jump to Basic.
  - Sheet 15 = CR01 10-site CME pack
- Or [MIMO_Deployment_v5.0.xlsx](MIMO_Deployment_v5.0.xlsx) via **Download raw file**
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
| 14 | MIMO Suggest + Incon (v5.0) | Section 1 FPD boxes (MML + Counter monitor / Impact on KPI / short Notes / Jump to Basic) · Section 2 dump enabled-or-not · Section 3 Enable/Fix proposal · Section 4 counters/KPI |
| 15 | CR01 MIMO Exec Pack | 10 gNB × 3 cells CME Proposed vs live + site-specific MML (101/102/103) |

Every MML table uses a **Seq** column (activation `n.1, n.2, …` then deactivation `n.D.*`), matching the CA workbook. **v5.0** combines former sheets 14+15. Every step sheet still ends with **Performance and Monitoring Counter**.

## How to regenerate

```bash
python3 tools/build_mimo_workbook.py          # v1 + v2.0 + v3.0 + v4.0 + v5.0
python3 tools/build_incon_report.py           # v2.0 only (needs v1 xlsx)
python3 tools/build_suggestions_sheet.py      # v3.0 only (needs v2.0 xlsx)
python3 tools/build_cr01_sheet.py             # v4.0 only (needs v3.0 xlsx)
python3 tools/build_combined_sheet.py         # v5.0 only (needs v4.0 xlsx)
```
