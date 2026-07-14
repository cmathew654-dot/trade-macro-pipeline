"""
THROWAWAY verification script for the W8 wash-sale flag audit (AuditWashFlag)
plus a spot-check that AuditSellSpecs / AuditRouting still produce PASS lines.

Flow: ProcessCDSHoldings -> AddRaiseCashScenarios -> BuildSellWorkbench ->
force FCAVX (known loss position, G/L=-1036.12) to Manual/ALL -> put FCAVX
into S2's buy plan ticker/amount -> AddBuyPlans -> AuditActiveCDSMath ->
assert a WARN "Wash-sale risk: ..." exists and the buy-plan ticker cell
interior is orange RGB(255,192,96); assert routing/spec audits PASS.
"""

import atexit
import ctypes
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
FIXTURE = os.path.join(REPO, "sample-data", "cds_holdings_raw_actual_export_shape.csv")

SKIP_IMPORT = {"ThisWorkbook.cls"}

RESULTS = []
_excel_pid = None


def assert_true(name, condition, detail=""):
    RESULTS.append((name, bool(condition), detail))
    print(("PASS " if condition else "FAIL ") + f"{name} - {detail}")
    return bool(condition)


def _kill_excel_pid():
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
        _staged_crlf_copy(frx, stage_dir)
    for fname in sorted(os.listdir(VBA_DIR)):
        if fname in SKIP_IMPORT:
            continue
        if os.path.splitext(fname)[1].lower() not in (".bas", ".cls", ".frm"):
            continue
        comps.Import(_staged_crlf_copy(os.path.join(VBA_DIR, fname), stage_dir))
    comps.Import(_staged_crlf_copy(SHIMS, stage_dir))


def find_processed_sheet(wb):
    for ws in wb.Worksheets:
        if str(ws.Cells(2, 1).Value or "").strip().upper() == "ASSET CLASS":
            return ws
    return wb.Worksheets(1)


def find_col_in_row(ws, row, text, max_col=400):
    for c in range(1, max_col):
        if str(ws.Cells(row, c).Value or "").strip() == text:
            return c
    return 0


def find_header_row(ws):
    for r in range(1, 11):
        if str(ws.Cells(r, 1).Value or "").strip().upper() == "ASSET CLASS":
            return r
    raise RuntimeError("header row not found")


def find_total_row(ws, data_start):
    for r in range(data_start, data_start + 500):
        a = str(ws.Cells(r, 1).Value or "").strip()
        e = str(ws.Cells(r, 5).Value or "").strip()
        if a == "" and e != "":
            return r
    raise RuntimeError("total row not found")


def find_ticker_row(ws, data_start, data_end, ticker):
    for r in range(data_start, data_end + 1):
        if str(ws.Cells(r, 3).Value or "").strip().upper() == ticker:
            return r
    return 0


def find_grand_total_row(ws, start_col, tot_row):
    for r in range(tot_row + 2, tot_row + 150):
        if str(ws.Cells(r, start_col).Value or "").strip().upper() == "GRAND TOTAL":
            return r
    return 0


def find_buy_plan_header_row(ws, start_col, pivot_end):
    for r in range(pivot_end + 1, pivot_end + 120):
        if str(ws.Cells(r, start_col).Value or "").strip().upper() == "BUY PLAN":
            return r
    return 0


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_wash_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    wb = None
    ok_overall = True
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "wash_verify.xlsm")
        wb.SaveAs(xlsm, FileFormat=52)
        import_vba(wb, workdir)
        qual = f"'{wb.Name}'!"

        app.Run(qual + "ProcessCDSHoldings")
        ws = find_processed_sheet(wb)

        app.Run(qual + "AddRaiseCashScenarios")
        app.Run(qual + "BuildSellWorkbench")
        app.Calculate()

        header_row = find_header_row(ws)
        data_start = header_row + 1
        total_row = find_total_row(ws, data_start)
        data_end = total_row - 1

        mode_col = find_col_in_row(ws, header_row, "Sell Mode")
        spec_type_col = find_col_in_row(ws, header_row, "Amt Type")
        spec_amt_col = find_col_in_row(ws, header_row, "Amount")
        target_label_col = find_col_in_row(ws, header_row, "Plan Target Raise")
        assert_true("workbench columns located",
                    all([mode_col, spec_type_col, spec_amt_col, target_label_col]),
                    f"mode={mode_col} type={spec_type_col} amt={spec_amt_col} target_label={target_label_col}")

        fcavx_row = find_ticker_row(ws, data_start, data_end, "FCAVX")
        assert_true("FCAVX found in holdings", fcavx_row > 0, f"row={fcavx_row}")
        gl = float(ws.Cells(fcavx_row, 6).Value or 0)
        assert_true("FCAVX has a loss (G/L < 0)", gl < 0, f"G/L={gl}")

        # Force FCAVX to Manual/ALL -> full-position loss sale.
        ws.Cells(fcavx_row, mode_col).Value = "Manual"
        ws.Cells(fcavx_row, spec_type_col).Value = "ALL"
        ws.Cells(fcavx_row, spec_amt_col).Value = 0

        target_input_col = target_label_col + 1
        ws.Cells(header_row, target_input_col).Value = 50000.0
        app.Calculate()

        # Put FCAVX into S2's buy plan ticker/amount inputs.
        S2_COL = 16 + (2 - 1) * 6  # ScenStartCol()=16, ScenStride()=6 defaults
        grand_row = find_grand_total_row(ws, S2_COL, total_row)
        assert_true("S2 pivot Grand Total row found", grand_row > 0, f"row={grand_row}")

        app.Run(qual + "AddBuyPlans")
        app.Calculate()

        buy_header_row = find_buy_plan_header_row(ws, S2_COL, grand_row)
        assert_true("S2 BUY PLAN block found", buy_header_row > 0, f"row={buy_header_row}")
        buy_input_row = buy_header_row + 2  # first buy-plan input row
        ws.Cells(buy_input_row, S2_COL).Value = "FCAVX"
        ws.Cells(buy_input_row, S2_COL + 1).Value = 10000
        app.Calculate()

        # --- Run the audit ---
        ws.Activate()
        app.Run(qual + "AuditActiveCDSMath")

        wash_warns = []
        spec_pass = False
        routing_pass_count = 0
        fails = []
        found_audit = False
        for awb in app.Workbooks:
            for aws in awb.Worksheets:
                if aws.Name == "CDS_MATH_AUDIT":
                    found_audit = True
                    r = 6
                    while str(aws.Cells(r, 1).Value or "").strip():
                        status = str(aws.Cells(r, 1).Value).strip().upper()
                        check = str(aws.Cells(r, 2).Value or "")
                        detail = str(aws.Cells(r, 3).Value or "")
                        if status == "FAIL":
                            fails.append(f"{check}: {detail}")
                        if status == "WARN" and "Wash-sale risk" in detail:
                            wash_warns.append(detail)
                        if status == "PASS" and "Sell spec Manual Sell $" in check:
                            spec_pass = True
                        if status == "PASS" and check.startswith("Routing"):
                            routing_pass_count += 1
                        r += 1
                    awb.Close(False)

        assert_true("CDS_MATH_AUDIT sheet produced", found_audit, "")
        assert_true("audit produced 0 FAIL", len(fails) == 0, "; ".join(fails[:8]))
        assert_true("WARN 'Wash-sale risk' present", any("Wash-sale risk" in w for w in wash_warns),
                    f"wash_warns={wash_warns}")
        assert_true("WARN mentions FCAVX and S2 buy plan",
                    any("FCAVX" in w and "S2" in w for w in wash_warns), f"wash_warns={wash_warns}")

        cell_color = ws.Cells(buy_input_row, S2_COL).Interior.Color
        expected_color = 96 * 65536 + 192 * 256 + 255  # RGB(255,192,96) as VBA long (BGR-packed)
        assert_true("buy-plan ticker cell interior is orange RGB(255,192,96)",
                    int(cell_color) == expected_color, f"actual_color={cell_color} expected={expected_color}")

        assert_true("AuditSellSpecs produced a PASS line", spec_pass, "")
        assert_true("AuditRouting produced PASS line(s)", routing_pass_count > 0, f"count={routing_pass_count}")

    except Exception:
        print("FULL TRACEBACK:\n" + traceback.format_exc())
        ok_overall = False
    finally:
        try:
            if wb is not None:
                wb.Close(False)
            app.Quit()
        except Exception:
            pass
        _kill_excel_pid()
        shutil.rmtree(workdir, ignore_errors=True)

    ok_overall = ok_overall and all(r[1] for r in RESULTS)
    print("\n=== SUMMARY ===")
    for name, ok, detail in RESULTS:
        print(f"{'PASS' if ok else 'FAIL'} | {name} | {detail}")
    print("OVERALL:", "PASS" if ok_overall else "FAIL")
    return 0 if ok_overall else 1


if __name__ == "__main__":
    sys.exit(main())
