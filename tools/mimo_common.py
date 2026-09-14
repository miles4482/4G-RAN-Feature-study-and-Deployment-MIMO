# Shared table helpers for MIMO deployment sheets (same column layout as the CA workbook).
import os
from mimo_excel_style import *

COLS = 10
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
FIG = os.path.join(ROOT, "mimo_figures")
STEP_WIDTHS = [8, 28, 32, 22, 28, 55, 42, 12, 12, 12]


def fig(*parts):
    return os.path.join(FIG, *parts)


def add_param_header(ws, r):
    return headers(ws, r, ["SN", "MO Name", "Parameter name", "Default value", "Recommended value",
                           "Parameter description", "Notes"] + [""] * 3)


def add_param(ws, r, sn, mo, name, default, reco, desc, notes, related=False):
    fh = PALE_GOLD if related else alt_fill(sn)
    vals = [sn, mo, name, default, reco, desc, notes, "", "", ""]
    r = table_row(ws, r, vals, fills=[fh] * 10, bolds=[False, True, True, False, True, False, False],
                  center_cols={1}, height=min(78, 28 + len(desc) // 90 * 12))
    merge(ws, r - 1, 7, r - 1, COLS)
    return r


def add_mml_header(ws, r):
    return headers(ws, r, ["SN", "Seq", "Mode", "RAT", "Phase", "MML command (verbatim from document)", "Remarks"] + [""] * 3)


def add_mml(ws, r, sn, seq, mode, rat, phase, cmd, remark):
    vals = [sn, seq, mode, rat, phase, cmd, remark, "", "", ""]
    r = table_row(ws, r, vals, fills=[alt_fill(sn)] * 10, bolds=[False, False, False, False, False, True, False],
                  center_cols={1, 2}, height=min(70, 24 + len(cmd) // 80 * 12))
    merge(ws, r - 1, 7, r - 1, COLS)
    return r


def add_mmls(ws, r, rows):
    """rows: (sn, seq, mode, rat, phase, cmd, remark)"""
    r = add_mml_header(ws, r)
    for rec in rows:
        r = add_mml(ws, r, *rec)
    return r


def add_ctr_header(ws, r):
    return headers(ws, r, ["SN", "Counter ID", "Counter Name", "Function / what it measures", "Use"] + [""] * 5)


def add_ctr(ws, r, sn, cid, name, fn, use):
    vals = [sn, cid, name, fn, use, "", "", "", "", ""]
    r = table_row(ws, r, vals, fills=[alt_fill(sn)] * 10, bolds=[False, False, True, False, False],
                  center_cols={1}, height=32)
    merge(ws, r - 1, 5, r - 1, COLS)
    return r


def add_kpi_header(ws, r):
    return headers(ws, r, ["SN", "KPI", "Formula (from document)", "Unit", "Notes"] + [""] * 5)


def add_kpi(ws, r, sn, kpi, formula, unit, notes):
    vals = [sn, kpi, formula, unit, notes, "", "", "", "", ""]
    r = table_row(ws, r, vals, fills=[alt_fill(sn)] * 10, bolds=[False, True, False, False, False],
                  center_cols={1}, height=36)
    merge(ws, r - 1, 5, r - 1, COLS)
    return r


def add_lic_header(ws, r):
    return headers(ws, r, ["SN", "RAT", "Feature ID", "Feature name", "Model", "Sales unit", "When consumed"] + [""] * 3)


def add_lic(ws, r, sn, rat, fid, fname, model, unit, when):
    vals = [sn, rat, fid, fname, model, unit, when, "", "", ""]
    r = table_row(ws, r, vals, fills=[alt_fill(sn)] * 10, bolds=[False, False, True, True, False, False, False],
                  center_cols={1}, height=36)
    merge(ws, r - 1, 7, r - 1, COLS)
    return r


def impact_table(ws, r, rows):
    r = headers(ws, r, ["SN", "Relation", "RAT", "Feature", "Parameter(s) — one row one feature", "Required action", "Impact / note"] + [""] * 3)
    for rec in rows:
        vals = list(rec) + [""] * 3
        rel = rec[1]
        fh = PALE_GREEN if rel == "Prerequisite" else (
            PALE_RED if rel == "Exclusive" else (
                PALE_GOLD if rel in ("Constraint", "Inherit") else alt_fill(int(str(rec[0]).split(".")[0] or 1))
            )
        )
        r = table_row(ws, r, vals, fills=[fh] * 10, height=44)
        merge(ws, r - 1, 7, r - 1, COLS)
    return r


def start_step(wb, sheet_name, banner_text, note):
    ws = wb.create_sheet(sheet_name)
    setup_sheet(ws, sheet_name)
    set_widths(ws, STEP_WIDTHS)
    r = 1
    r = banner(ws, r, COLS, banner_text)
    r = note_bar(ws, r, COLS, note)
    return ws, r
