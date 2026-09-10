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
    build_v2()
    build_v3()
    build_v4()


if __name__ == "__main__":
    main()
