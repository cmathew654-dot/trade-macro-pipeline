"""
THROWAWAY verification script for the W2 proceeds routing block.
Not part of the repo test suite -- lives in the scratchpad only.

Flow: ProcessCDSHoldings -> classify unknowns -> AddRaiseCashScenarios ->
SpawnScenario (S3, done BEFORE BuildSellWorkbench since spawning clears/
repositions the whole scenario+summary+workbench area -- pre-existing
behavior, not part of W2) -> BuildSellWorkbench -> set S3 manual sells ->
set Plan Target Raise (S2) -> exercise the PROCEEDS ROUTING block (default
seed, custom split, over-allocation, flip-back) -> AddBuyPlans -> assert S2's
funding switches to the routing "Buy Plan" Routed $ cell while S3 keeps
legacy (own Raise $ total) funding -> AuditActiveCDSMath and report the
full PASS/WARN/FAIL breakdown.
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


def assert_near(name, actual, expected, tol=0.5):
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


def read_audit_result(app, wb):
    for awb in app.Workbooks:
        for aws in awb.Worksheets:
            if aws.Name == "CDS_MATH_AUDIT":
                fails, warns, passes = [], 0, 0
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
                return passes, warns, fails
    return 0, 0, ["CDS_MATH_AUDIT sheet not found"]


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_w2_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    wb = None
    ok_overall = True
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "w2_verify.xlsm")
        wb.SaveAs(xlsm, FileFormat=52)
        import_vba(wb, workdir)
        qual = f"'{wb.Name}'!"

        app.Run(qual + "ProcessCDSHoldings")
        ws = find_processed_sheet(wb)
        classify_unknowns(app, wb, ws, qual)

        # AddRaiseCashScenarios creates S1 only (count=1). Spawn twice to
        # reach S1+S2+S3 (count=3) BEFORE building the sell workbench:
        # SpawnScenario clears/repositions the whole scenario+summary+
        # workbench area as part of shifting layout for each new scenario
        # column block (pre-existing behavior, unrelated to routing).
        # BuildSellWorkbench (and the routing block it re-creates) must run
        # after the scenario count is final, same as any other workbench
        # rebuild trigger. With count already >= 2 by the time it runs,
        # BuildSellWorkbench's own EnsureSellPlanScenario is a no-op and S3
        # survives.
        app.Run(qual + "AddRaiseCashScenarios")
        app.Run(qual + "SpawnScenario")  # -> S2
        app.Run(qual + "SpawnScenario")  # -> S3

        app.Run(qual + "BuildSellWorkbench")
        app.Calculate()

        # --- Give S3 (manual scenario) a couple of real sells so its legacy
        # funding assertion is non-trivial (not just 0 == 0). Pick non-CASH
        # rows with enough FMV to comfortably hold a sell amount. ---
        S3_COL = 16 + (3 - 1) * 6  # ScenStartCol()=16, ScenStride()=6 defaults
        header_row, _ = bulk_find_first(ws, "ASSET CLASS")
        data_start = header_row + 1
        total_row_s3 = None
        for r in range(data_start, data_start + 200):
            if str(ws.Cells(r, 1).Value or "").strip() == "" and str(ws.Cells(r, 5).Value or "").strip() != "":
                total_row_s3 = r
                break
        data_end_s3 = total_row_s3 - 1

        candidate_rows = []
        for r in range(data_start, data_end_s3 + 1):
            cls = str(ws.Cells(r, 1).Value or "").strip().upper()
            fmv = float(ws.Cells(r, 5).Value or 0)
            if cls not in ("CASH", "SHORT") and fmv > 5000:
                candidate_rows.append(r)
        assert_true("found 2+ non-cash rows with FMV>5000 for S3 sells", len(candidate_rows) >= 2, str(candidate_rows[:3]))
        s3_row1, s3_row2 = candidate_rows[0], candidate_rows[1]
        ws.Cells(s3_row1, S3_COL).Value = 4000
        ws.Cells(s3_row2, S3_COL).Value = 2500
        app.Calculate()

        target_label_row, target_label_col = bulk_find_first(ws, "Plan Target Raise")
        target_input_col = target_label_col + 1

        TARGET_RAISE = 30000.0
        ws.Cells(target_label_row, target_input_col).Value = TARGET_RAISE
        app.Calculate()

        # --- Locate the routing block ---
        title_row, dest_col = bulk_find_first(ws, "PROCEEDS ROUTING")
        assert_true("PROCEEDS ROUTING block auto-seeded by BuildSellWorkbench",
                    title_row > 0 and dest_col > 0, f"row={title_row} col={dest_col}")

        detail_col = dest_col + 1
        spec_col = dest_col + 2
        amt_col = dest_col + 3
        routed_col = dest_col + 4
        data_start_r = title_row + 2
        status_row = data_start_r + 5

        # Sanity: default seed = Buy Plan/Residual, Money Market/$0, blanks
        assert_true("default row12 = Buy Plan/Residual",
                    str(ws.Cells(data_start_r, dest_col).Value) == "Buy Plan"
                    and str(ws.Cells(data_start_r, spec_col).Value) == "Residual",
                    f"dest={ws.Cells(data_start_r, dest_col).Value} spec={ws.Cells(data_start_r, spec_col).Value}")
        assert_true("default row13 = Money Market/$/DefaultMM",
                    str(ws.Cells(data_start_r + 1, dest_col).Value) == "Money Market"
                    and str(ws.Cells(data_start_r + 1, spec_col).Value) == "$",
                    f"dest={ws.Cells(data_start_r + 1, dest_col).Value} spec={ws.Cells(data_start_r + 1, spec_col).Value} "
                    f"detail={ws.Cells(data_start_r + 1, detail_col).Value}")
        assert_true("default status = Routing OK (100% to Buy Plan)",
                    str(ws.Cells(status_row, dest_col).Value) == "Routing OK",
                    str(ws.Cells(status_row, dest_col).Value))

        # --- Configure the 18000 / 7000 / 5000 split ---
        buy_row = data_start_r          # row 12: Buy Plan
        mm_row = data_start_r + 1       # row 13: Money Market
        xfer_row = data_start_r + 2     # row 14: Transfer Out (was blank)

        ws.Cells(buy_row, spec_col).Value = "$"
        ws.Cells(buy_row, amt_col).Value = 18000
        ws.Cells(mm_row, spec_col).Value = "$"
        ws.Cells(mm_row, amt_col).Value = 7000
        ws.Cells(xfer_row, dest_col).Value = "Transfer Out"
        ws.Cells(xfer_row, detail_col).Value = "to Schwab checking"
        ws.Cells(xfer_row, spec_col).Value = "Residual"
        app.Calculate()

        assert_near("Routed $ Buy Plan = 18000", ws.Cells(buy_row, routed_col).Value, 18000, tol=0.5)
        assert_near("Routed $ Money Market = 7000", ws.Cells(mm_row, routed_col).Value, 7000, tol=0.5)
        assert_near("Routed $ Transfer Out (residual) = 5000", ws.Cells(xfer_row, routed_col).Value, 5000, tol=0.5)
        assert_true("status = Routing OK", str(ws.Cells(status_row, dest_col).Value) == "Routing OK",
                    str(ws.Cells(status_row, dest_col).Value))

        # --- Over-allocation ---
        ws.Cells(buy_row, amt_col).Value = 30000
        app.Calculate()
        status_text = str(ws.Cells(status_row, dest_col).Value)
        assert_true("status shows over-allocation", "over-allocat" in status_text.lower(), status_text)

        # --- Flip back ---
        ws.Cells(buy_row, amt_col).Value = 18000
        app.Calculate()
        assert_true("status back to Routing OK after flip-back",
                    str(ws.Cells(status_row, dest_col).Value) == "Routing OK",
                    str(ws.Cells(status_row, dest_col).Value))

        # --- Buy plans ---
        app.Run(qual + "AddBuyPlans")
        app.Calculate()

        def find_available_to_buy(scenCol):
            for r, c in bulk_find_all(ws, "Available to Buy"):
                if c == scenCol:
                    return ws.Cells(r, c + 2).Value
            return None

        def find_scenario_raise_total(scenCol):
            # report totals row: col A blank, col E non-blank
            for r in range(1, 200):
                if str(ws.Cells(r, 1).Value or "").strip() == "" and str(ws.Cells(r, 5).Value or "").strip() != "":
                    return ws.Cells(r, scenCol).Value
            return None

        S2_COL = 16 + (2 - 1) * 6

        s2_available = find_available_to_buy(S2_COL)
        s3_available = find_available_to_buy(S3_COL)
        s3_raise_total = find_scenario_raise_total(S3_COL)

        assert_near("S2 Available to Buy = 18000 (routing funding, not 30000)", s2_available, 18000, tol=0.5)
        assert_true("S3 Available to Buy present", s3_available is not None, str(s3_available))
        assert_true("S3 raise total is non-trivial (manual sells entered)",
                    s3_raise_total is not None and float(s3_raise_total) > 100, str(s3_raise_total))
        if s3_available is not None and s3_raise_total is not None:
            assert_near("S3 Available to Buy keeps legacy funding (== S3 Raise $ total)",
                        s3_available, s3_raise_total, tol=0.5)

        # --- Audit ---
        ws.Activate()
        app.Run(qual + "AuditActiveCDSMath")
        passes, warns, fails = read_audit_result(app, wb)
        print(f"\nAUDIT: {passes} pass / {warns} warn / {len(fails)} fail")
        for f in fails:
            print("  FAIL:", f)

        # W11: MathAudit.bas is now routing-aware -- S2's "Available to Buy" /
        # "Diff vs Raise" expectation is resolved the same way the sheet
        # itself resolves S2's funding (routing block's "Buy Plan" Routed $
        # when a PROCEEDS ROUTING block exists), so the non-default split
        # (18000 funding vs 30000 raised) should no longer surface any FAIL.
        assert_true("no FAILs (routing-aware audit)", len(fails) == 0,
                    "; ".join(fails) if fails else "(none)")

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
