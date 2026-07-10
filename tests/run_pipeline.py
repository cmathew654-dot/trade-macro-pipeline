"""Headless pipeline harness for CDS Trade Assistant.

Drives a real Excel instance over a synthetic fixture CSV:
  import vba/* into a macro workbook -> ProcessCDSHoldings ->
  classify unknowns -> SaveUnknownsAndRefresh -> AddRaiseCashScenarios ->
  BuildSellWorkbench -> AddBuyPlans -> AuditActiveCDSMath

MsgBox/InputBox are shadowed by tests/TestShims.bas so nothing blocks.
Exit code 0 = all steps + audit assertions passed.

Usage:
  python tests/run_pipeline.py [--fixture sample-data/<name>.csv] [--visible]
"""

import argparse
import atexit
import ctypes
import json
import os
import shutil
import sys
import tempfile
import traceback

import pythoncom
import win32com.client

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VBA_DIR = os.path.join(REPO, "vba")
SHIMS = os.path.join(REPO, "tests", "TestShims.bas")
DEFAULT_FIXTURE = os.path.join(
    REPO, "sample-data", "cds_holdings_raw_actual_export_shape.csv"
)

# Document module (must not be imported as a component) and binary-paired form
SKIP_IMPORT = {"ThisWorkbook.cls"}

RESULTS = {"steps": [], "ok": True}
_excel_pid = None


def step(name, ok, detail=""):
    RESULTS["steps"].append({"step": name, "ok": bool(ok), "detail": str(detail)[:400]})
    if not ok:
        RESULTS["ok"] = False
    print(("PASS " if ok else "FAIL ") + name + (" - " + str(detail)[:200] if detail else ""))


def _kill_excel_pid():
    """Terminate only the Excel instance this harness started."""
    global _excel_pid
    if _excel_pid:
        handle = ctypes.windll.kernel32.OpenProcess(1, False, _excel_pid)
        if handle:
            ctypes.windll.kernel32.TerminateProcess(handle, 0)
            ctypes.windll.kernel32.CloseHandle(handle)
        _excel_pid = None


def excel_pid(app):
    import win32process

    hwnd = app.Hwnd
    return win32process.GetWindowThreadProcessId(hwnd)[1]


def _staged_crlf_copy(src, stage_dir):
    """VBE's .frm parser requires CRLF; stage a normalized copy of any source."""
    dst = os.path.join(stage_dir, os.path.basename(src))
    with open(src, "rb") as f:
        b = f.read()
    if os.path.splitext(src)[1].lower() != ".frx":
        b = b.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    with open(dst, "wb") as f:
        f.write(b)
    return dst


def import_vba(wb, stage_dir):
    comps = wb.VBProject.VBComponents
    frx = os.path.join(VBA_DIR, "frmCDSTradeAssistant.frx")
    if os.path.exists(frx):
        _staged_crlf_copy(frx, stage_dir)  # sibling needed by the .frm import
    for fname in sorted(os.listdir(VBA_DIR)):
        if fname in SKIP_IMPORT:
            continue
        if os.path.splitext(fname)[1].lower() not in (".bas", ".cls", ".frm"):
            continue
        comps.Import(_staged_crlf_copy(os.path.join(VBA_DIR, fname), stage_dir))
    comps.Import(_staged_crlf_copy(SHIMS, stage_dir))


def find_processed_sheet(wb):
    """Processed report marker: header row 2 col A is literally 'ASSET CLASS'."""
    for ws in wb.Worksheets:
        if str(ws.Cells(2, 1).Value or "").strip().upper() == "ASSET CLASS":
            return ws
    return wb.Worksheets(1)


def used_range_dump(ws, max_rows=80, max_cols=30):
    """Small debug dump helper (only printed on failure)."""
    out = []
    for r in range(1, max_rows + 1):
        row = []
        for c in range(1, max_cols + 1):
            v = ws.Cells(r, c).Value
            if v is not None:
                row.append(f"{r}:{c}={v}")
        if row:
            out.append(" | ".join(row))
    return "\n".join(out)


def run(fixture, visible=False, keep=False):
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = visible
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_harness_")
    csv_copy = os.path.join(workdir, os.path.basename(fixture))
    shutil.copy2(fixture, csv_copy)

    wb = None
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "harness_run.xlsm")
        wb.SaveAs(xlsm, FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
        step("open+saveas", True, xlsm)

        import_vba(wb, workdir)
        step("import_vba", True, f"{wb.VBProject.VBComponents.Count} components")

        qual = f"'{wb.Name}'!"

        app.Run(qual + "ProcessCDSHoldings")
        ws = find_processed_sheet(wb)
        step("process_holdings", True, ws.Name)

        # --- classify unknowns if the review box appeared (layout patched in
        # --- after module mapping; see classify_unknowns)
        n = classify_unknowns(app, wb, ws, qual)
        step("classify_unknowns", True, f"{n} unknowns classified")

        app.Run(qual + "AddRaiseCashScenarios")
        step("add_scenarios", True)

        app.Run(qual + "BuildSellWorkbench")
        step("build_workbench", True)

        set_target_raise(ws, 20000.0)
        app.Calculate()
        step("set_target_raise", True, "20000")

        app.Run(qual + "AddBuyPlans")
        step("add_buy_plans", True)

        ws.Activate()
        app.Run(qual + "AuditActiveCDSMath")
        ok, detail = read_audit_result(app, wb)
        step("math_audit", ok, detail)

        if keep:
            wb.Save()
            print("kept:", xlsm)
    except Exception as e:
        print("FULL TRACEBACK:\n" + traceback.format_exc())
        print("EXC REPR:", repr(e)[:2000])
        step("exception", False, repr(e)[:1500])
        try:
            print(used_range_dump(find_processed_sheet(wb)))
        except Exception:
            pass
    finally:
        try:
            if wb is not None and not keep:
                wb.Close(False)
            app.Quit()
        except Exception:
            pass
        _kill_excel_pid()
        if not keep:
            shutil.rmtree(workdir, ignore_errors=True)

    print(json.dumps(RESULTS))
    return 0 if RESULTS["ok"] else 1


# --------------------------------------------------------------------------
# Layout-dependent helpers (facts from the module layout map):
#   report: header row 2, data row 3+, FMV col E, class col A, qty col K
#   unknowns box: col M title row, headers +1, data +2 (ticker M, class O)
#   workbench target input: row 2 cell carrying the "total cash amount" comment
#   audit: new workbook, sheet CDS_MATH_AUDIT, statuses col A rows 6+
# --------------------------------------------------------------------------

UNKNOWN_COL = 13  # M
UNKNOWN_CLASS_COL = 15  # O
DEFAULT_UNKNOWN_CLASS = "STOCK"


def classify_unknowns(app, wb, ws, qual):
    """Fill the unknown-ticker review box (if present) and save. Returns count."""
    title_row = None
    for r in range(1, 300):
        v = str(ws.Cells(r, UNKNOWN_COL).Value or "")
        if "UNKNOWN" in v.upper():
            title_row = r
            break
    if title_row is None:
        return 0
    n = 0
    r = title_row + 2
    while str(ws.Cells(r, UNKNOWN_COL).Value or "").strip():
        if not str(ws.Cells(r, UNKNOWN_CLASS_COL).Value or "").strip():
            ws.Cells(r, UNKNOWN_CLASS_COL).Value = DEFAULT_UNKNOWN_CLASS
            n += 1
        r += 1
    if n:
        app.Run(qual + "SaveUnknownsAndRefresh")
    return n


def set_target_raise(ws, amount):
    """The workbench target input is the row-2 cell whose comment explains it."""
    for cm in ws.Comments:
        if "total cash amount" in str(cm.Text() or "").lower():
            cm.Parent.Value = amount
            return
    raise RuntimeError("workbench target input cell not found (no comment match)")


def read_audit_result(app, wb):
    """AuditActiveCDSMath writes a new workbook w/ sheet CDS_MATH_AUDIT."""
    for awb in app.Workbooks:
        for aws in awb.Worksheets:
            if aws.Name == "CDS_MATH_AUDIT":
                statuses, fails, warns = [], [], 0
                r = 6
                while str(aws.Cells(r, 1).Value or "").strip():
                    status = str(aws.Cells(r, 1).Value).strip().upper()
                    statuses.append(status)
                    if status == "FAIL":
                        fails.append(
                            f"{aws.Cells(r, 2).Value}: {aws.Cells(r, 3).Value}"
                        )
                    elif status == "WARN":
                        warns += 1
                    r += 1
                awb.Close(False)
                ok = bool(statuses) and not fails
                detail = (
                    f"{statuses.count('PASS')} pass / {warns} warn / "
                    f"{len(fails)} fail"
                    + ("; " + " || ".join(fails[:5]) if fails else "")
                )
                return ok, detail
    return False, "CDS_MATH_AUDIT sheet not found"


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--fixture", default=DEFAULT_FIXTURE)
    p.add_argument("--visible", action="store_true")
    p.add_argument("--keep", action="store_true", help="keep the temp xlsm for inspection")
    a = p.parse_args()
    sys.exit(run(a.fixture, a.visible, a.keep))
