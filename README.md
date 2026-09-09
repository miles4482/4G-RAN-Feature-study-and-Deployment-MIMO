# 4G-RAN-Feature-study-and-Deployment-MIMO

Combined **5G MIMO deployment workbook** extracted from the eight Huawei 5G RAN10.1 Feature Parameter Descriptions in this repository (same document-style Excel as the Carrier Aggregation workbook).

## Download (do not use GitHub Raw for `.xlsx`)

- **v2.0 (current):** [MIMO_Deployment_v2.0.zip](MIMO_Deployment_v2.0.zip) → **Download raw file**, then unzip. Last sheet is the dump vs commercial MIMO inconsistency report.
- Or [MIMO_Deployment_v2.0.xlsx](MIMO_Deployment_v2.0.xlsx) via **Download raw file**
- v1 (deployment steps only, no dump sheet): [MIMO_Deployment.xlsx](MIMO_Deployment.xlsx)

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
| 14 | MIMO Incon Report (v2.0 last sheet) | Live DHK 5G CME dump 8 Sep 2026 vs commercial 32T good-DL-tput |

Every MML table uses a **Seq** column (activation `n.1, n.2, …` then deactivation `n.D.*`), matching the CA workbook. The v2.0 last sheet uses the 11-column inconsistency MML layout (SN / RAT / Doc / Action / MML / MO / Parameter / Live dump / proposed / notes / License).

## How to regenerate

```bash
python3 tools/build_mimo_workbook.py          # v1 + v2.0
python3 tools/build_incon_report.py           # v2.0 only (needs v1 xlsx)
```
