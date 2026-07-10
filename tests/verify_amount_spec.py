"""
THROWAWAY verification script for the W1 amount-spec engine.
Not part of the repo test suite -- lives in the scratchpad only.

Drives the same COM flow as tests/run_pipeline.py, but after
BuildSellWorkbench it sets row-level amount specs (Shares / % Pos / ALL),
recalculates, and asserts:
  - Manual Sell $ (a formula now) matches the expected value per spec type
  - Manual Used $ / Proposed Sell $ dispatch still caps + pools correctly
  - Pool rows fill the remainder up to the target raise
  - Sell Mode / Amt Type / Amount stay editable (Locked=False) after
    AddBuyPlans reruns ApplyScenarioUXRulesToSheet (wholesale re-lock)
  - Manual Sell $ stays locked (it's a formula column now)
  - AuditActiveCDSMath reports 0 FAIL
"""

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

REPO = r"C:\Users\Cyril\Projects\cds-trade-assistant"
VBA_DIR = os.path.join(REPO, "vba")
SHIMS = os.path.join(REPO, "tests", "TestShims.bas")
FIXTURE = os.path.join(REPO, "sample-data", "cds_holdings_raw_actual_export_shape.csv")

SKIP_IMPORT = {"ThisWorkbook.cls"}

RESULTS = []
_excel_pid = None


def assert_near(name, actual, expected, tol=0.05):
    ok = abs(float(actual) - float(expected)) <= tol
    RESULTS.append((name, ok, actual, expected))
    print(("PASS " if ok else "FAIL ") + f"{name}: actual={actual!r} expected={expected!r} (tol={tol})")
    return ok


def assert_true(name, condition, detail=""):
    RESULTS.append((name, bool(condition), detail, True))
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


UNKNOWN_COL = 13
UNKNOWN_CLASS_COL = 15
DEFAULT_UNKNOWN_CLASS = "STOCK"


def classify_unknowns(app, wb, ws, qual):
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


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_w1_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    wb = None
    ok_overall = True
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "w1_verify.xlsm")
        wb.SaveAs(xlsm, FileFormat=52)
        import_vba(wb, workdir)
        qual = f"'{wb.Name}'!"

        app.Run(qual + "ProcessCDSHoldings")
        ws = find_processed_sheet(wb)
        classify_unknowns(app, wb, ws, qual)

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
        manual_col = find_col_in_row(ws, header_row, "Manual Sell $")
        proposed_col = find_col_in_row(ws, header_row, "Proposed Sell $")
        used_header_col = find_col_in_row(ws, header_row, "Manual Used $")
        status_col = find_col_in_row(ws, header_row, "Plan Status")
        target_label_col = find_col_in_row(ws, header_row, "Plan Target Raise")

        assert_true("workbench columns located",
                     all([mode_col, spec_type_col, spec_amt_col, manual_col,
                          proposed_col, used_header_col, status_col, target_label_col]),
                     f"mode={mode_col} type={spec_type_col} amt={spec_amt_col} "
                     f"manual={manual_col} proposed={proposed_col} used={used_header_col} "
                     f"status={status_col} target_label={target_label_col}")

        # Column order contract check: type=mode+1, amt=mode+2, manual=mode+3
        assert_true("column order Mode,AmtType,Amount,Manual contiguous",
                     spec_type_col == mode_col + 1 and spec_amt_col == mode_col + 2
                     and manual_col == mode_col + 3,
                     f"mode={mode_col} type={spec_type_col} amt={spec_amt_col} manual={manual_col}")

        target_input_col = target_label_col + 1

        # Gather row data: ticker(C=3), class(A=1), FMV(E=5), Qty(K=11)
        rows = []
        for r in range(data_start, data_end + 1):
            cls = str(ws.Cells(r, 1).Value or "").strip().upper()
            ticker = str(ws.Cells(r, 3).Value or "").strip()
            fmv = float(ws.Cells(r, 5).Value or 0)
            qty = ws.Cells(r, 11).Value
            qty = float(qty) if qty not in (None, "") else 0.0
            rows.append({"row": r, "class": cls, "ticker": ticker, "fmv": fmv, "qty": qty})

        equity_rows = sorted(
            [x for x in rows if x["class"] not in ("CASH", "SHORT", "BOND") and x["fmv"] > 0],
            key=lambda x: -x["fmv"],
        )
        assert_true("at least one equity-like row found", len(equity_rows) >= 1, str(equity_rows[:3]))
        shares_row = equity_rows[0]

        remaining = sorted(
            [x for x in rows if x["row"] != shares_row["row"] and x["class"] != "CASH" and x["fmv"] > 0],
            key=lambda x: -x["fmv"],
        )
        assert_true("2 more distinct rows available for %%Pos/ALL", len(remaining) >= 2, str(remaining[:3]))
        pctpos_row = remaining[0]
        all_row = remaining[1]

        print(f"shares_row={shares_row['ticker']} fmv={shares_row['fmv']} qty={shares_row['qty']}")
        print(f"pctpos_row={pctpos_row['ticker']} fmv={pctpos_row['fmv']}")
        print(f"all_row={all_row['ticker']} fmv={all_row['fmv']}")

        def set_spec(rowinfo, spec_type, amount):
            r = rowinfo["row"]
            ws.Cells(r, mode_col).Value = "Manual"
            ws.Cells(r, spec_type_col).Value = spec_type
            ws.Cells(r, spec_amt_col).Value = amount

        set_spec(shares_row, "Shares", 50)
        set_spec(pctpos_row, "% Pos", 25)
        set_spec(all_row, "ALL", 0)

        TARGET_RAISE = 60000.0
        ws.Cells(header_row, target_input_col).Value = TARGET_RAISE

        app.Calculate()

        # --- Assertion 1: Manual Sell $ formula results ---
        price = shares_row["fmv"] / shares_row["qty"] if shares_row["qty"] else 0.0
        expected_shares_manual = 50 * price
        actual_shares_manual = float(ws.Cells(shares_row["row"], manual_col).Value or 0)
        assert_near("Manual Sell $ (Shares spec)", actual_shares_manual, expected_shares_manual, tol=0.05)

        expected_pctpos_manual = 0.25 * pctpos_row["fmv"]
        actual_pctpos_manual = float(ws.Cells(pctpos_row["row"], manual_col).Value or 0)
        assert_near("Manual Sell $ (% Pos spec)", actual_pctpos_manual, expected_pctpos_manual, tol=0.05)

        expected_all_manual = all_row["fmv"]
        actual_all_manual = float(ws.Cells(all_row["row"], manual_col).Value or 0)
        assert_near("Manual Sell $ (ALL spec)", actual_all_manual, expected_all_manual, tol=0.05)

        # --- Assertion 2: used/proposed dispatch ---
        used_col = mode_col  # placeholder, real used col located below
        used_col = find_col_in_row(ws, header_row, "Manual Used $")
        for rowinfo, label in ((shares_row, "shares"), (pctpos_row, "%pos"), (all_row, "ALL")):
            manual_v = float(ws.Cells(rowinfo["row"], manual_col).Value or 0)
            used_v = float(ws.Cells(rowinfo["row"], used_col).Value or 0)
            proposed_v = float(ws.Cells(rowinfo["row"], proposed_col).Value or 0)
            expected_used = min(max(0.0, manual_v), rowinfo["fmv"])
            assert_near(f"Manual Used $ cap ({label})", used_v, expected_used, tol=0.05)
            assert_near(f"Proposed Sell $ dispatch ({label})", proposed_v, used_v, tol=0.05)

        manual_used_total = sum(
            min(max(0.0, float(ws.Cells(x["row"], manual_col).Value or 0)), x["fmv"])
            for x in (shares_row, pctpos_row, all_row)
        )
        residual_expected = max(0.0, TARGET_RAISE - manual_used_total)

        pool_rows = [x for x in rows if x["row"] not in
                     (shares_row["row"], pctpos_row["row"], all_row["row"]) and x["class"] != "CASH"]
        pool_proposed_total = sum(float(ws.Cells(x["row"], proposed_col).Value or 0) for x in pool_rows)
        assert_near("Pool fills remainder after manual specs", pool_proposed_total, residual_expected, tol=1.0)

        total_proposed = float(ws.Cells(total_row, proposed_col).Value or 0)
        assert_near("Total Proposed matches target (pool has capacity)", total_proposed, TARGET_RAISE, tol=1.0)

        # --- Assertion 3: ScenarioUX unlock survives AddBuyPlans ---
        app.Run(qual + "AddBuyPlans")
        app.Calculate()

        mode_locked = bool(ws.Cells(shares_row["row"], mode_col).Locked)
        type_locked = bool(ws.Cells(shares_row["row"], spec_type_col).Locked)
        amt_locked = bool(ws.Cells(shares_row["row"], spec_amt_col).Locked)
        manual_locked = bool(ws.Cells(shares_row["row"], manual_col).Locked)

        assert_true("Sell Mode stays editable after AddBuyPlans", not mode_locked, f"Locked={mode_locked}")
        assert_true("Amt Type stays editable after AddBuyPlans", not type_locked, f"Locked={type_locked}")
        assert_true("Amount stays editable after AddBuyPlans", not amt_locked, f"Locked={amt_locked}")
        assert_true("Manual Sell $ stays locked (formula) after AddBuyPlans", manual_locked, f"Locked={manual_locked}")

        # --- Assertion 4: math audit clean ---
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
    for name, ok, actual, expected in RESULTS:
        print(f"{'PASS' if ok else 'FAIL'} | {name} | actual={actual} expected={expected}")
    print("OVERALL:", "PASS" if ok_overall else "FAIL")
    return 0 if ok_overall else 1


if __name__ == "__main__":
    sys.exit(main())
