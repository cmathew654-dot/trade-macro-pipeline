"""
THROWAWAY verification script for W10 (trade email narrates the plan
faithfully). Not part of the repo test suite -- lives in the scratchpad
only.

Flow: ProcessCDSHoldings -> classify unknowns -> AddRaiseCashScenarios ->
BuildSellWorkbench (auto-creates S2) -> set one Shares spec + one ALL spec
(all other rows Excluded so proceeds are deterministic) -> configure the
PROCEEDS ROUTING block (Buy Plan $ / Money Market $ / Transfer Out
Residual) -> AddBuyPlans + fill one buy row -> set Settings!AutomationMode
= 1 directly on the CDS_Settings sheet -> queue TestInputs "2" so the
scenario-pick InputBox resolves to S2 -> run GenerateTradeEmail -> assert
the "CDS Email Preview" sheet's body narrates the shares spec, the ALL
spec, and every routing destination.

Then push the routing block into an over-allocated state and rerun:
assert GenerateTradeEmail refuses (no preview update, refusal MsgBox
logged to TestLog).
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


def bulk_find_all(ws, text):
    """Fast text search over UsedRange via a single Value2 bulk read."""
    used = ws.UsedRange
    r0, c0 = used.Row, used.Column
    vals = used.Value2
    out = []
    for ri, row in enumerate(vals):
        for ci, v in enumerate(row):
            if v is not None and str(v).strip() == text:
                out.append((r0 + ri, c0 + ci))
    return out


def bulk_find_first(ws, text):
    found = bulk_find_all(ws, text)
    return found[0] if found else (0, 0)


def find_settings_sheet(wb):
    for ws in wb.Worksheets:
        if ws.Name == "CDS_Settings":
            return ws
    return None


def set_setting(wb, key, value):
    """Directly write key/value into the CDS_Settings sheet's A:B area
    (the sheet already exists by this point - GetSetting/GetOrCreateSettingsSheet
    has been called many times by the prior macro runs)."""
    ws = find_settings_sheet(wb)
    if ws is None:
        raise RuntimeError("CDS_Settings sheet not found - expected it to be seeded already")
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp
    for r in range(2, last_row + 1):
        if str(ws.Cells(r, 1).Value or "").strip() == key:
            ws.Cells(r, 2).Value = value
            return True
    # Not found (shouldn't happen) - append it.
    ws.Cells(last_row + 1, 1).Value = key
    ws.Cells(last_row + 1, 2).Value = value
    return False


def queue_test_input(wb, values):
    ws = None
    for w in wb.Worksheets:
        if w.Name == "TestInputs":
            ws = w
            break
    if ws is None:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "TestInputs"
    ws.Cells.Clear()
    for i, v in enumerate(values):
        ws.Cells(i + 1, 1).Value = v
    ws.Range("B1").Value = 1


def test_log_text(wb):
    for w in wb.Worksheets:
        if w.Name == "TestLog":
            used = w.UsedRange
            r0, c0 = used.Row, used.Column
            vals = used.Value2
            lines = []
            for row in vals:
                for v in row:
                    if v:
                        lines.append(str(v))
            return "\n".join(lines)
    return ""


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_w10_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    wb = None
    ok_overall = True
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "w10_verify.xlsm")
        wb.SaveAs(xlsm, FileFormat=52)
        import_vba(wb, workdir)
        qual = f"'{wb.Name}'!"

        app.Run(qual + "ProcessCDSHoldings")
        ws = find_processed_sheet(wb)
        classify_unknowns(app, wb, ws, qual)

        app.Run(qual + "AddRaiseCashScenarios")   # S1
        app.Run(qual + "BuildSellWorkbench")      # auto-creates S2 (EnsureSellPlanScenario)
        app.Calculate()

        header_row = find_header_row(ws)
        data_start = header_row + 1
        total_row = find_total_row(ws, data_start)
        data_end = total_row - 1

        mode_col = find_col_in_row(ws, header_row, "Sell Mode")
        spec_type_col = find_col_in_row(ws, header_row, "Amt Type")
        spec_amt_col = find_col_in_row(ws, header_row, "Amount")
        status_col = find_col_in_row(ws, header_row, "Plan Status")
        assert_true("workbench columns located",
                     all([mode_col, spec_type_col, spec_amt_col, status_col]),
                     f"mode={mode_col} type={spec_type_col} amt={spec_amt_col} status={status_col}")

        # Gather row data: ticker(C=3), class(A=1), FMV(E=5), Qty(K=11)
        rows = []
        for r in range(data_start, data_end + 1):
            cls = str(ws.Cells(r, 1).Value or "").strip().upper()
            ticker = str(ws.Cells(r, 3).Value or "").strip()
            fmv = float(ws.Cells(r, 5).Value or 0)
            qty = ws.Cells(r, 11).Value
            qty = float(qty) if qty not in (None, "") else 0.0
            rows.append({"row": r, "class": cls, "ticker": ticker, "fmv": fmv, "qty": qty})

        # Deterministic proceeds: exclude everything, then pick two distinct
        # equity-like rows for the Shares and ALL specs.
        for x in rows:
            ws.Cells(x["row"], mode_col).Value = "Exclude"

        equity_rows = sorted(
            [x for x in rows if x["class"] not in ("CASH", "SHORT") and x["fmv"] > 0 and x["qty"] > 0],
            key=lambda x: -x["fmv"],
        )
        assert_true("2+ equity-like rows found for specs", len(equity_rows) >= 2, str(equity_rows[:3]))
        shares_row = equity_rows[0]
        all_row = equity_rows[1]

        def set_spec(rowinfo, spec_type, amount):
            r = rowinfo["row"]
            ws.Cells(r, mode_col).Value = "Manual"
            ws.Cells(r, spec_type_col).Value = spec_type
            ws.Cells(r, spec_amt_col).Value = amount

        SHARES_QTY = min(50, shares_row["qty"])
        set_spec(shares_row, "Shares", SHARES_QTY)
        set_spec(all_row, "ALL", 0)
        app.Calculate()

        # --- Total proceeds (deterministic: no pool rows contribute) ---
        total_proposed = float(ws.Cells(6, status_col + 1).Value or 0)
        assert_true("Total Proposed > 0", total_proposed > 0, str(total_proposed))

        # --- Configure PROCEEDS ROUTING: Buy Plan $ / Money Market $ / Transfer Out Residual ---
        title_row, dest_col = bulk_find_first(ws, "PROCEEDS ROUTING")
        assert_true("PROCEEDS ROUTING block found", title_row > 0 and dest_col > 0,
                    f"row={title_row} col={dest_col}")
        assert_true("routing dest_col == workbench status_col", dest_col == status_col,
                    f"dest_col={dest_col} status_col={status_col}")

        detail_col = dest_col + 1
        spec_col = dest_col + 2
        amt_col_r = dest_col + 3
        routed_col = dest_col + 4
        data_start_r = title_row + 2
        status_row = data_start_r + 5

        buy_row = data_start_r          # row 12: Buy Plan (default)
        mm_row = data_start_r + 1       # row 13: Money Market (default)
        xfer_row = data_start_r + 2     # row 14: Transfer Out (blank by default)

        BUY_AMT = round(total_proposed * 0.5)
        MM_AMT = round(total_proposed * 0.3)
        XFER_DETAIL = "Community Bank checking x5678"

        ws.Cells(buy_row, spec_col).Value = "$"
        ws.Cells(buy_row, amt_col_r).Value = BUY_AMT
        ws.Cells(mm_row, spec_col).Value = "$"
        ws.Cells(mm_row, amt_col_r).Value = MM_AMT
        ws.Cells(xfer_row, dest_col).Value = "Transfer Out"
        ws.Cells(xfer_row, detail_col).Value = XFER_DETAIL
        ws.Cells(xfer_row, spec_col).Value = "Residual"
        app.Calculate()

        mm_detail = str(ws.Cells(mm_row, detail_col).Value or "").strip()

        status_text = str(ws.Cells(status_row, dest_col).Value or "")
        assert_true("routing status = Routing OK before email", status_text == "Routing OK", status_text)

        # --- Buy plan: build the grid, then fund one row from the routed Buy Plan $ ---
        app.Run(qual + "AddBuyPlans")
        app.Calculate()

        S2_COL = 16 + (2 - 1) * 6  # ScenStartCol()=16, ScenStride()=6 defaults
        bp_hits = [c for (r, c) in bulk_find_all(ws, "BUY PLAN") if c == S2_COL]
        assert_true("S2 BUY PLAN header found", len(bp_hits) >= 0, str(bp_hits))
        bp_header_row = None
        for r, c in bulk_find_all(ws, "BUY PLAN"):
            if c == S2_COL:
                bp_header_row = r
                break
        assert_true("S2 BUY PLAN header row located", bp_header_row is not None, str(bp_header_row))
        bp_input_start = bp_header_row + 2

        ws.Cells(bp_input_start, S2_COL).Value = "AAPL"
        ws.Cells(bp_input_start, S2_COL + 1).Value = BUY_AMT
        app.Calculate()

        # --- AutomationMode = 1 (direct settings-sheet write) ---
        set_setting(wb, "AutomationMode", "1")

        # --- Queue the scenario-pick InputBox answer ("2") for both runs ---
        queue_test_input(wb, ["2", "2"])

        # --- Run 1: expect a PASS -- preview sheet populated ---
        ws.Activate()
        app.Run(qual + "GenerateTradeEmail")

        preview_ws = None
        for w in wb.Worksheets:
            if w.Name == "CDS Email Preview":
                preview_ws = w
                break
        assert_true("CDS Email Preview sheet created", preview_ws is not None, "")

        body1 = str(preview_ws.Cells(2, 1).Value or "") if preview_ws else ""
        subject1 = str(preview_ws.Cells(1, 1).Value or "") if preview_ws else ""
        assert_true("subject non-empty", subject1 != "", subject1)

        assert_true('body contains "shares of"', "shares of" in body1, body1[:400])
        assert_true('body contains SELLING ALL marker', "SELLING ALL" in body1, body1[:400])
        assert_true('body contains "Of the ~$"', "Of the ~$" in body1, body1[:400])
        assert_true(f'body contains "remains in {mm_detail} (money market)"',
                    f"remains in {mm_detail} (money market)" in body1, body1[:400])
        assert_true('body contains "to be transferred out"', "to be transferred out" in body1, body1[:400])
        assert_true(f'body contains transfer detail "{XFER_DETAIL}"', XFER_DETAIL in body1, body1[:400])

        # --- Push routing into an over-allocated state ---
        ws.Cells(buy_row, amt_col_r).Value = total_proposed * 3
        app.Calculate()
        bad_status = str(ws.Cells(status_row, dest_col).Value or "")
        assert_true("routing status now over-allocated", "over-allocat" in bad_status.lower(), bad_status)

        # --- Run 2: expect a refusal -- no preview update, refusal logged ---
        ws.Activate()
        app.Run(qual + "GenerateTradeEmail")

        body2 = str(preview_ws.Cells(2, 1).Value or "") if preview_ws else ""
        assert_true("preview body unchanged after refusal", body2 == body1,
                    f"len(body1)={len(body1)} len(body2)={len(body2)}")

        log_text = test_log_text(wb)
        assert_true('TestLog contains refusal MsgBox text',
                    "does not reconcile" in log_text and "PROCEEDS ROUTING" in log_text,
                    log_text[-600:])

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
