#!/usr/bin/env python3
"""Build the combined 5G MIMO deployment workbook (document-style, sequenced).

Source FPDs in this repository (5G RAN10.1). Sheet/MML order is the mandatory
live-network sequence (same approach as the Carrier Aggregation workbook).
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook
from build_cover_overview import build_cover, build_overview, build_types
from build_steps_foundation import (
    build_step1, build_step2, build_step3, build_step4, build_step5, build_step6,
)
from build_steps_advanced import (
    build_step7, build_step8, build_step9, build_step10, build_step11,
)

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(ROOT, "MIMO_Deployment.xlsx")


def main():
    wb = Workbook()
    default = wb.active
    default.title = "_tmp"

    print("cover...")
    build_cover(wb)
    print("overview...")
    build_overview(wb)
    print("types...")
    build_types(wb)
    print("step1 basic...")
    build_step1(wb)
    print("step2 su...")
    build_step2(wb)
    print("step3 mu...")
    build_step3(wb)
    print("step4 multilayer...")
    build_step4(wb)
    print("step5 beam...")
    build_step5(wb)
    print("step6 ahr...")
    build_step6(wb)
    print("step7 ibeam...")
    build_step7(wb)
    print("step8 ul boosting...")
    build_step8(wb)
    print("step9 das+fusion...")
    build_step9(wb)
    print("step10 mmwave...")
    build_step10(wb)
    print("step11 cable sequence...")
    build_step11(wb)

    if "_tmp" in wb.sheetnames:
        del wb["_tmp"]

    colors = ["1F4E79", "2E75B6", "0D7377", "C00000", "C65911", "548235",
              "7030A0", "1F4E79", "2E75B6", "0D7377", "C00000", "C65911",
              "548235", "7030A0"]
    for i, ws in enumerate(wb.worksheets):
        ws.sheet_properties.tabColor = colors[i % len(colors)]
        ws.sheet_view.showGridLines = False

    print("saving", OUT)
    wb.save(OUT)
    print("ok", os.path.getsize(OUT), "sheets", len(wb.worksheets))

    # v2.0 = v1 + dump inconsistency; v3.0 = v2 + document suggestion boxes
    # v4.0 = v3 + CR01 10-site pack + Performance and Monitoring Counter on every sheet
    from build_incon_report import main as build_v2
    from build_suggestions_sheet import main as build_v3
    from build_cr01_sheet import main as build_v4
    # v5.0 is frozen (dump as a separate Section 2). v6.0 merges dump onto each MML.
    # v7.0 = v6 + pre-requisite columns on every MML + Master Findings + Action Plan.
    # v8.0 = v7 + sequenced pairing rollback pack (SRS_IC first).
    # v9.0 = pairing lift plan (N.ChMeas.MIMO.DL.MuPairing as north-star).
    # v10.0 = dual-target action plan (MuPairing AND DL user throughput).
    from build_combined_sheet import main as build_v6
    from build_v7_findings import main as build_v7
    from build_v8_rollback import main as build_v8
    from build_v9_pairing_plan import main as build_v9
    from build_v10_dual_plan import main as build_v10
    build_v2()
    build_v3()
    build_v4()
    build_v6()
    build_v7()
    build_v8()
    build_v9()
    build_v10()


if __name__ == "__main__":
    main()
