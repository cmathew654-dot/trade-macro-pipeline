"""
PERMANENT regression test for the CDS Trade Assistant guard features:

  1. Multi-account guard  (CDS_Holdings_Processor.ResolveAccountFilter)
  2. Illiquid CUSIP default Sell Mode  (CDS_Sell_Workbench.NormalizeSellMode)
  3. Cash/MM-available advisory  (CDS_Sell_Workbench.ApplyWorkbenchStatus, row 8)
  4. Target shortfall alarm      (CDS_Sell_Workbench.ApplyWorkbenchStatus, row 3)
  5. AuditActiveCDSMath stays clean (0 FAIL) across all of the above

Drives the same COM flow as tests/run_pipeline.py (import_vba, header/total-row
finders, TestShims-answered dialogs) against sample-data/cds_holdings_raw_two_accounts.csv,
which contains two accounts: XXXX-2026 (IWM, SPYV, ...) and XXXX-3031 (an 8-digit
numeric CUSIP bond + AAPL). TestShims.bas pops InputBox answers off a TestInputs
sheet (col A, pointer B1); that sheet is created and populated BEFORE running any
macro that prompts.

Run:
  python tests/verify_guards.py
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

REPO = r"C:\Users\Cyril\Projects\cds-trade-assistant"
TESTS_DIR = os.path.join(REPO, "tests")
if TESTS_DIR not in sys.path:
    sys.path.insert(0, TESTS_DIR)

import run_pipeline as rp  # reuse import_vba / find_processed_sheet / classify_unknowns / excel_pid

FIXTURE = os.path.join(REPO, "sample-data", "cds_holdings_raw_two_accounts.csv")

RESULTS = []
_excel_pid = None


def assert_true(name, condition, detail=""):
    ok = bool(condition)
    RESULTS.append((name, ok, detail))
    print(("PASS " if ok else "FAIL ") + name + (f" - {detail}" if detail else ""))
    return ok


def _kill_excel_pid():
    """Terminate only the Excel instance this harness started."""
    global _excel_pid
    if _excel_pid:
        handle = ctypes.windll.kernel32.OpenProcess(1, False, _excel_pid)
        if handle:
            ctypes.windll.kernel32.TerminateProcess(handle, 0)
            ctypes.windll.kernel32.CloseHandle(handle)
        _excel_pid = None


def find_header_row(ws):
    for r in range(1, 11):
        if str(ws.Cells(r, 1).Value or "").strip().upper() == "ASSET CLASS":
            return r
    raise RuntimeError("header row (ASSET CLASS) not found")


def find_total_row(ws, data_start):
    for r in range(data_start, data_start + 500):
        a = str(ws.Cells(r, 1).Value or "").strip()
        e = str(ws.Cells(r, 5).Value or "").strip()
        if a == "" and e != "":
            return r
    raise RuntimeError("total row not found")


def find_col_in_row(ws, row, text, max_col=400):
    for c in range(1, max_col):
        if str(ws.Cells(row, c).Value or "").strip() == text:
            return c
    return 0


def cell_text(v):
    """COM hands back all-digit tickers (e.g. numeric CUSIPs) as Python floats
    (38141508.0); VBA's CStr() on the same cell yields "38141508". Normalize
    the same way here so ticker comparisons match what the VBA code sees."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def add_test_inputs_sheet(wb, answers):
    """TestShims.InputBox pops answers off TestInputs col A, pointer in B1."""
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "TestInputs"
    for i, v in enumerate(answers, start=1):
        ws.Cells(i, 1).Value = v
    ws.Range("B1").Value = 1
    return ws


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = rp.excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_guards_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    wb = None
    ok_overall = True
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "guards_verify.xlsm")
        wb.SaveAs(xlsm, FileFormat=52)
        rp.import_vba(wb, workdir)
        qual = f"'{wb.Name}'!"

        # Queue the multi-account picker's answer BEFORE running ProcessCDSHoldings.
        # "2" selects the second distinct account found (XXXX-3031).
        data_ws = wb.ActiveSheet
        add_test_inputs_sheet(wb, ["2"])
        # Worksheets.Add() makes the new sheet active; ProcessCDSHoldings runs
        # against ActiveSheet, so re-activate the raw CSV data sheet.
        data_ws.Activate()

        app.Run(qual + "ProcessCDSHoldings")
        ws = rp.find_processed_sheet(wb)
        rp.classify_unknowns(app, wb, ws, qual)

        # --- Assertion 1: multi-account guard ---
        a1_stamp = str(ws.Cells(1, 1).Value or "").strip()
        assert_true("Multi-account guard: A1 stamped with the picked account",
                     a1_stamp == "XXXX-3031", f"A1={a1_stamp!r} (expected XXXX-3031)")

        header_row = find_header_row(ws)
        data_start = header_row + 1
        total_row = find_total_row(ws, data_start)
        data_end = total_row - 1

        tickers = [cell_text(ws.Cells(r, 3).Value) for r in range(data_start, data_end + 1)]
        assert_true("Multi-account guard: first account's tickers excluded",
                     "IWM" not in tickers and "SPYV" not in tickers,
                     f"tickers={tickers}")
        assert_true("Multi-account guard: second account's rows present",
                     "AAPL" in tickers and "38141508" in tickers,
                     f"tickers={tickers}")

        app.Run(qual + "AddRaiseCashScenarios")
        app.Run(qual + "BuildSellWorkbench")
        app.Calculate()

        mode_col = find_col_in_row(ws, header_row, "Sell Mode")
        assert_true("Sell Workbench located (Sell Mode column found)", mode_col > 0, f"mode_col={mode_col}")

        # Column-order contract from CDS_Sell_Workbench.bas BuildSellWorkbenchOnSheet:
        # modeCol, +1 Amt Type, +2 Amount, +3 Manual Sell $, +4 Proposed Sell $,
        # +5 Manual Used $, +7 Plan Status; target input sits at modeCol-2.
        target_input_col = mode_col - 2
        status_col = mode_col + 7

        rows = []
        for r in range(data_start, data_end + 1):
            cls = str(ws.Cells(r, 1).Value or "").strip().upper()
            ticker = cell_text(ws.Cells(r, 3).Value)
            fmv = float(ws.Cells(r, 5).Value or 0)
            mode = str(ws.Cells(r, mode_col).Value or "").strip()
            rows.append({"row": r, "class": cls, "ticker": ticker, "fmv": fmv, "mode": mode})

        bond_row = next((x for x in rows if x["ticker"] == "38141508"), None)
        aapl_row = next((x for x in rows if x["ticker"] == "AAPL"), None)
        assert_true("CUSIP bond row and AAPL row both present in workbench",
                     bond_row is not None and aapl_row is not None, f"rows={rows}")

        # --- Assertion 2: illiquid CUSIP default ---
        assert_true("Illiquid CUSIP bond row defaults Sell Mode to Exclude",
                     bond_row["mode"] == "Exclude", f"mode={bond_row['mode']!r}")
        assert_true("Equity (AAPL) row defaults Sell Mode to Pool",
                     aapl_row["mode"] == "Pool", f"mode={aapl_row['mode']!r}")

        # --- Assertion 3: cash/MM-available advisory (status row 8) ---
        cash_short_total = sum(x["fmv"] for x in rows if x["class"] in ("CASH", "SHORT"))
        assert_true("account has a CASH/SHORT balance to test the advisory against",
                     cash_short_total > 0, f"cash_short_total={cash_short_total}")

        below_target = round(cash_short_total / 2.0, 2)
        ws.Cells(header_row, target_input_col).Value = below_target
        app.Calculate()
        advisory_below = str(ws.Cells(8, status_col).Value or "").strip()
        assert_true("Cash-available advisory fires when target <= CASH+SHORT",
                     "redemption may cover" in advisory_below,
                     f"target={below_target} cash_short_total={cash_short_total} advisory={advisory_below!r}")

        above_target = round(cash_short_total * 4.0, 2)
        ws.Cells(header_row, target_input_col).Value = above_target
        app.Calculate()
        advisory_above = str(ws.Cells(8, status_col).Value or "").strip()
        assert_true("Cash-available advisory is blank when target > CASH+SHORT",
                     advisory_above == "",
                     f"target={above_target} cash_short_total={cash_short_total} advisory={advisory_above!r}")

        # --- Assertion 4: shortfall alarm (status row 3) ---
        sellable_fmv = sum(x["fmv"] for x in rows if x["mode"] in ("Pool", "Manual"))
        assert_true("account has sellable (Pool/Manual) FMV to test the shortfall alarm against",
                     sellable_fmv > 0, f"sellable_fmv={sellable_fmv}")

        shortfall_target = round(sellable_fmv * 10.0, 2)
        ws.Cells(header_row, target_input_col).Value = shortfall_target
        app.Calculate()
        status_shortfall = str(ws.Cells(3, status_col).Value or "").strip()
        assert_true("Shortfall alarm fires with a $ amount when target far exceeds sellable FMV",
                     status_shortfall.startswith("SHORTFALL") and "$" in status_shortfall,
                     f"target={shortfall_target} status={status_shortfall!r}")

        achievable_target = round(sellable_fmv * 0.5, 2)
        ws.Cells(header_row, target_input_col).Value = achievable_target
        app.Calculate()
        status_ok = str(ws.Cells(3, status_col).Value or "").strip()
        assert_true("Status returns to the OK message once target is achievable",
                     status_ok == "OK: proposed sells match target.",
                     f"target={achievable_target} status={status_ok!r}")

        # --- Assertion 5: math audit clean ---
        app.Run(qual + "AddBuyPlans")
        app.Calculate()
        ws.Activate()
        app.Run(qual + "AuditActiveCDSMath")

        fails = []
        warns = 0
        passes = 0
        found_audit = False
        for awb in app.Workbooks:
            for aws in awb.Worksheets:
                if aws.Name == "CDS_MATH_AUDIT":
                    found_audit = True
                    r = 6
                    while str(aws.Cells(r, 1).Value or "").strip():
                        status = str(aws.Cells(r, 1).Value).strip().upper()
                        if status == "FAIL":
                            fails.append(f"{aws.Cells(r, 2).Value}: {aws.Cells(r, 3).Value}")
                        elif status == "WARN":
                            warns += 1
                        elif status == "PASS":
                            passes += 1
                        r += 1
                    awb.Close(False)

        assert_true("CDS_MATH_AUDIT sheet produced", found_audit, "")
        assert_true("AuditActiveCDSMath: 0 FAIL", len(fails) == 0,
                    f"{passes} pass / {warns} warn / {len(fails)} fail" +
                    (("; " + " || ".join(fails[:8])) if fails else ""))

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
        print(f"{'PASS' if ok else 'FAIL'} | {name}" + (f" | {detail}" if detail else ""))
    print("OVERALL:", "PASS" if ok_overall else "FAIL")
    return 0 if ok_overall else 1


if __name__ == "__main__":
    sys.exit(main())
