"""
Regression verification for the W3 price staleness guard (CDS_PriceGuard.bas).

Drives the same COM flow as tests/run_pipeline.py through BuildSellWorkbench,
then runs RefreshLivePrices and asserts:
  - "CDS Live Prices" helper sheet exists with an A1 "Checked:" stamp and the
    row-2 headers Ticker | Class | Implied Price | Live Price | Drift % | Note
  - CASH and CJTXX (SHORT / money market) rows are skipped
  - Implied Price = report FMV(E)/Qty(K) within 0.01 for IWM and SPYV
  - report FMV column E is byte-identical before/after (source data immutable)
  - each Live Price row is either a resolved numeric quote or blank with note
    "n/a - not resolved" / "linked data types unavailable" -- both outcomes
    are acceptable; which one occurs depends on this machine's M365 licensing
    and connectivity, and the script prints which branch actually ran
  - workbench status line (row 9, Plan Status col): "Prices checked ..." when
    quotes flowed, or legitimately blank when the linked-data-types call was
    unavailable (graceful path leaves the report completely untouched)

Then exercises the ApplyDriftAlerts seam standalone (no re-fetch) by seeding
Live Price values directly into the helper sheet:
  - +5% drift -> report ticker cell (col C) turns red RGB(255,120,120)
  - +0.2% drift -> ticker cell stays uncolored (threshold DriftAlertPct=1.0)
  - re-seeding the drifted row back under threshold clears the red interior
    on rerun (idempotent), and the status line flips to "no drift"
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


def snapshot_col_e(ws, data_start, total_row):
    return [ws.Cells(r, 5).Value for r in range(data_start, total_row + 1)]


def find_price_sheet(wb):
    for ws in wb.Worksheets:
        if ws.Name == "CDS Live Prices":
            return ws
    return None


def price_sheet_rows(price_ws):
    """One dict per data row (row 3 down to first blank Ticker)."""
    rows = []
    r = 3
    while str(price_ws.Cells(r, 1).Value or "").strip():
        rows.append({
            "row": r,
            "ticker": str(price_ws.Cells(r, 1).Value or "").strip(),
            "cls": str(price_ws.Cells(r, 2).Value or "").strip(),
            "implied": price_ws.Cells(r, 3).Value,
            "live": price_ws.Cells(r, 4).Value,
            "drift": price_ws.Cells(r, 5).Value,
            "note": str(price_ws.Cells(r, 6).Value or "").strip(),
        })
        r += 1
    return rows


def find_report_row(ws, data_start, data_end, ticker):
    for r in range(data_start, data_end + 1):
        if str(ws.Cells(r, 3).Value or "").strip() == ticker:
            return r
    return None


RED = 255 + 120 * 256 + 120 * 65536  # RGB(255,120,120) as OLE color int


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_w3_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    wb = None
    ok_overall = True
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "w3_verify.xlsm")
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

        fmv_before = snapshot_col_e(ws, data_start, total_row)

        # --- Main flow: RefreshLivePrices = FetchLivePrices + ApplyDriftAlerts ---
        app.Run(qual + "RefreshLivePrices")
        app.Calculate()

        fmv_after_refresh = snapshot_col_e(ws, data_start, total_row)
        assert_true("report FMV col E unchanged after RefreshLivePrices",
                    fmv_after_refresh == fmv_before,
                    "diffs at offsets: " + str([i for i, (a, b) in enumerate(zip(fmv_before, fmv_after_refresh)) if a != b][:5]))

        price_ws = find_price_sheet(wb)
        assert_true("CDS Live Prices sheet exists", price_ws is not None)
        if price_ws is None:
            raise RuntimeError("helper sheet missing; cannot continue")

        stamp = str(price_ws.Cells(1, 1).Value or "")
        assert_true("A1 has a Checked stamp", stamp.startswith("Checked:"), stamp)

        headers = [str(price_ws.Cells(2, c).Value or "") for c in range(1, 7)]
        assert_true("header row 2 = Ticker|Class|Implied Price|Live Price|Drift %|Note",
                    headers == ["Ticker", "Class", "Implied Price", "Live Price", "Drift %", "Note"],
                    str(headers))

        rows = price_sheet_rows(price_ws)
        tickers_present = {r["ticker"] for r in rows}
        assert_true("CJTXX (SHORT) skipped", "CJTXX" not in tickers_present, str(sorted(tickers_present))[:200])
        assert_true("CASH skipped", "CASH" not in tickers_present, str(sorted(tickers_present))[:200])

        by_ticker = {r["ticker"]: r for r in rows}
        for tk in ("IWM", "SPYV"):
            assert_true(f"{tk} present in CDS Live Prices", tk in by_ticker, str(sorted(tickers_present))[:200])

        # Implied Price = report FMV/Qty, within 0.01
        for tk in ("IWM", "SPYV"):
            if tk not in by_ticker:
                continue
            rrow = find_report_row(ws, data_start, data_end, tk)
            assert_true(f"{tk} found on report sheet", rrow is not None)
            if rrow is not None:
                fmv = float(ws.Cells(rrow, 5).Value)
                qty = float(ws.Cells(rrow, 11).Value)
                assert_near(f"{tk} Implied Price == FMV/Qty", by_ticker[tk]["implied"], fmv / qty, tol=0.01)

        # Live Price: either a resolved numeric quote, or blank with an
        # explanatory note. Both are acceptable (network/licensing-dependent).
        resolved = 0
        unresolved = 0
        for r in rows:
            live = r["live"]
            note = r["note"]
            is_numeric = isinstance(live, (int, float))
            is_blank_with_note = (live in (None, "")) and note in (
                "n/a - not resolved", "linked data types unavailable")
            assert_true(f"{r['ticker']} Live Price numeric or blank+note",
                        is_numeric or is_blank_with_note, f"live={live!r} note={note!r}")
            if is_numeric:
                resolved += 1
            else:
                unresolved += 1
        print(f"\nLIVE QUOTE RESOLUTION on this machine: {resolved} resolved / "
              f"{unresolved} unresolved (of {len(rows)})\n")

        # Workbench status line (row 9, same column as "Plan Status").
        # If ConvertToLinkedDataType itself is unavailable (older Excel build /
        # not licensed for Data Types), the module's documented graceful path
        # skips ApplyDriftAlerts entirely -- report completely untouched -- so
        # row 9 legitimately stays blank in that branch.
        all_unavailable = bool(rows) and all(
            r["note"] == "linked data types unavailable" for r in rows)
        status_col = find_col_in_row(ws, header_row, "Plan Status")
        assert_true("Plan Status header found (Sell Workbench present)", status_col > 0)
        status_text = str(ws.Cells(9, status_col).Value or "")
        if all_unavailable:
            print("CONNECTIVITY: Stocks linked data types unavailable on this machine "
                  "-- graceful-degradation branch exercised.")
            assert_true("row 9 left blank (report untouched - unavailable branch)",
                        status_text == "", status_text)
        else:
            assert_true("row 9 status contains 'Prices checked'",
                        "Prices checked" in status_text, status_text)

        # --- ApplyDriftAlerts seam: seed Live Price directly (no re-fetch) ---
        assert_true("IWM/SPYV rows available for seeded drift test",
                    "IWM" in by_ticker and "SPYV" in by_ticker)
        iwm_row = by_ticker["IWM"]["row"]
        spyv_row = by_ticker["SPYV"]["row"]
        iwm_implied = float(price_ws.Cells(iwm_row, 3).Value)
        spyv_implied = float(price_ws.Cells(spyv_row, 3).Value)

        # IWM: +5% drift (over 1.0% threshold); SPYV: +0.2% (under threshold)
        price_ws.Cells(iwm_row, 4).Value = iwm_implied * 1.05
        price_ws.Cells(spyv_row, 4).Value = spyv_implied * 1.002
        price_ws.Cells(iwm_row, 6).Value = ""
        price_ws.Cells(spyv_row, 6).Value = ""

        ws.Activate()
        app.Run(qual + "ApplyDriftAlerts")
        app.Calculate()

        iwm_report_row = find_report_row(ws, data_start, data_end, "IWM")
        spyv_report_row = find_report_row(ws, data_start, data_end, "SPYV")

        iwm_color = ws.Cells(iwm_report_row, 3).Interior.Color
        spyv_color = ws.Cells(spyv_report_row, 3).Interior.Color
        assert_true("IWM ticker cell red (drift +5% > 1% threshold)",
                    iwm_color == RED, f"color={iwm_color}")
        assert_true("SPYV ticker cell NOT red (drift +0.2% <= 1% threshold)",
                    spyv_color != RED, f"color={spyv_color}")

        status_text2 = str(ws.Cells(9, status_col).Value or "")
        assert_true("status line updated after seeded ApplyDriftAlerts",
                    "Prices checked" in status_text2 and "IWM" in status_text2, status_text2)

        fmv_after_apply = snapshot_col_e(ws, data_start, total_row)
        assert_true("report FMV col E still unchanged after ApplyDriftAlerts",
                    fmv_after_apply == fmv_before,
                    "diffs at offsets: " + str([i for i, (a, b) in enumerate(zip(fmv_before, fmv_after_apply)) if a != b][:5]))

        # --- Idempotent clear: re-seed IWM under threshold, rerun, expect clear ---
        price_ws.Cells(iwm_row, 4).Value = iwm_implied * 1.002
        price_ws.Cells(iwm_row, 6).Value = ""
        ws.Activate()
        app.Run(qual + "ApplyDriftAlerts")
        app.Calculate()

        iwm_color_after = ws.Cells(iwm_report_row, 3).Interior.Color
        assert_true("IWM ticker cell cleared once drift back under threshold (idempotent)",
                    iwm_color_after != RED, f"color={iwm_color_after}")

        status_text3 = str(ws.Cells(9, status_col).Value or "")
        assert_true("status line shows no drift after rerun",
                    "no drift" in status_text3.lower(), status_text3)

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
