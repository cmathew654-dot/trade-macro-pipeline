"""
THROWAWAY verification script for W9 (snapshot system). Not part of the
repo test suite -- lives in the scratchpad only.

Flow: ProcessCDSHoldings -> classify unknowns -> AddRaiseCashScenarios (S1)
-> BuildSellWorkbench (auto-creates S2 + default PROCEEDS ROUTING block) ->
set target raise + one Manual/$ spec row + a Money Market routing amount ->
SaveCDSSnapshot -> assert SNAP sheet properties (A2 marker, no formulas on
previously-formula cells, validations gone, tab gray, index row w/
hyperlink + Target Raise + Total Proposed + Allocation) -> assert source
sheet still active + AuditActiveCDSMath still 0 FAIL -> ExportSnapshotToFile
-> assert xlsx on disk -> SaveCDSSnapshot again while a SNAP sheet is
active -> assert refusal (no new sheet, TestLog has the refusal text).
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
    raise RuntimeError("header row not found")


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


def set_target_raise(ws, amount):
    for cm in ws.Comments:
        if "total cash amount" in str(cm.Text() or "").lower():
            cm.Parent.Value = amount
            return
    raise RuntimeError("workbench target input cell not found")


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


def read_audit_result(app, wb):
    for awb in app.Workbooks:
        for aws in awb.Worksheets:
            if aws.Name == "CDS_MATH_AUDIT":
                statuses, fails, warns = [], [], 0
                r = 6
                while str(aws.Cells(r, 1).Value or "").strip():
                    status = str(aws.Cells(r, 1).Value).strip().upper()
                    statuses.append(status)
                    if status == "FAIL":
                        fails.append(f"{aws.Cells(r, 2).Value}: {aws.Cells(r, 3).Value}")
                    elif status == "WARN":
                        warns += 1
                    r += 1
                awb.Close(False)
                ok = bool(statuses) and not fails
                detail = (f"{statuses.count('PASS')} pass / {warns} warn / {len(fails)} fail"
                          + ("; " + " || ".join(fails[:5]) if fails else ""))
                return ok, detail
    return False, "CDS_MATH_AUDIT sheet not found"


def rgb_tuple(color_long):
    color_long = int(color_long)
    r = color_long & 0xFF
    g = (color_long >> 8) & 0xFF
    b = (color_long >> 16) & 0xFF
    return (r, g, b)


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_w9_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    wb = None
    ok_overall = True
    try:
        wb = app.Workbooks.Open(csv_copy)
        xlsm = os.path.join(workdir, "w9_verify.xlsm")
        wb.SaveAs(xlsm, FileFormat=52)
        import_vba(wb, workdir)
        qual = f"'{wb.Name}'!"

        app.Run(qual + "ProcessCDSHoldings")
        ws = find_processed_sheet(wb)
        classify_unknowns(app, wb, ws, qual)

        app.Run(qual + "AddRaiseCashScenarios")   # S1
        app.Run(qual + "BuildSellWorkbench")      # auto-creates S2 + default routing block
        app.Calculate()

        header_row = find_header_row(ws)
        data_start = header_row + 1
        total_row = find_total_row(ws, data_start)
        data_end = total_row - 1

        mode_col = find_col_in_row(ws, header_row, "Sell Mode")
        spec_type_col = find_col_in_row(ws, header_row, "Amt Type")
        spec_amt_col = find_col_in_row(ws, header_row, "Amount")
        proposed_col = find_col_in_row(ws, header_row, "Proposed Sell $")
        status_col = find_col_in_row(ws, header_row, "Plan Status")
        assert_true("workbench columns located",
                     all([mode_col, spec_type_col, spec_amt_col, proposed_col, status_col]),
                     f"mode={mode_col} type={spec_type_col} amt={spec_amt_col} proposed={proposed_col} status={status_col}")

        set_target_raise(ws, 20000.0)

        # Set one manual $ spec on a non-cash equity row.
        manual_row = None
        for r in range(data_start, data_end + 1):
            cls = str(ws.Cells(r, 1).Value or "").strip().upper()
            fmv = float(ws.Cells(r, 5).Value or 0)
            if cls not in ("CASH", "SHORT") and fmv > 0:
                manual_row = r
                break
        assert_true("manual-spec row found", manual_row is not None, str(manual_row))
        ws.Cells(manual_row, mode_col).Value = "Manual"
        ws.Cells(manual_row, spec_type_col).Value = "$"
        ws.Cells(manual_row, spec_amt_col).Value = 2000
        app.Calculate()

        # Give the routing block's Money Market row a nonzero $ so its
        # Routed $ formula cell resolves to a nonzero value.
        title_row, dest_col = bulk_find_first(ws, "PROCEEDS ROUTING")
        assert_true("PROCEEDS ROUTING block found", title_row > 0 and dest_col > 0,
                    f"row={title_row} col={dest_col}")
        detail_col = dest_col + 1
        spec_col = dest_col + 2
        amt_col = dest_col + 3
        routed_col = dest_col + 4
        data_start_r = title_row + 2
        mm_row = data_start_r + 1
        ws.Cells(mm_row, spec_col).Value = "$"
        ws.Cells(mm_row, amt_col).Value = 5000
        app.Calculate()

        # --- "Before" formula strings for cells that must be frozen ---
        S1_COL = 16  # ScenStartCol() default
        s1_raise_before = str(ws.Cells(manual_row, S1_COL).Formula)
        proposed_before = str(ws.Cells(manual_row, proposed_col).Formula)
        routed_before = str(ws.Cells(mm_row, routed_col).Formula)
        assert_true("S1 raise cell is a formula before snapshot", s1_raise_before.startswith("="), s1_raise_before)
        assert_true("workbench proposed cell is a formula before snapshot", proposed_before.startswith("="), proposed_before)
        assert_true("routing Routed $ cell is a formula before snapshot", routed_before.startswith("="), routed_before)

        sheet_names_before = {s.Name for s in wb.Worksheets}

        # --- SaveCDSSnapshot ---
        ws.Activate()
        app.Run(qual + "SaveCDSSnapshot")

        sheet_names_after = {s.Name for s in wb.Worksheets}
        new_sheets = sheet_names_after - sheet_names_before
        snap_sheets = [n for n in new_sheets if n.startswith("SNAP ")]
        assert_true("exactly one new SNAP sheet created", len(snap_sheets) == 1, str(new_sheets))
        snap_name = snap_sheets[0] if snap_sheets else None
        snap_ws = wb.Worksheets(snap_name) if snap_name else None

        assert_true("SNAP sheet A2 marker",
                    snap_ws is not None and str(snap_ws.Cells(2, 1).Value) == "ASSET CLASS (SNAPSHOT)",
                    str(snap_ws.Cells(2, 1).Value) if snap_ws else "n/a")

        if snap_ws is not None:
            s1_after = str(snap_ws.Cells(manual_row, S1_COL).Formula)
            proposed_after = str(snap_ws.Cells(manual_row, proposed_col).Formula)
            routed_after = str(snap_ws.Cells(mm_row, routed_col).Formula)
            assert_true("S1 raise cell has no formula on snapshot", not s1_after.startswith("="), s1_after)
            assert_true("workbench proposed cell has no formula on snapshot", not proposed_after.startswith("="), proposed_after)
            assert_true("routing Routed $ cell has no formula on snapshot", not routed_after.startswith("="), routed_after)

            # --- Validations gone ---
            set_ok = True
            try:
                snap_ws.Cells(manual_row, mode_col).Value = "NotAValidMode"
            except Exception as e:
                set_ok = False
                print("unexpected exception writing invalid Sell Mode value:", e)
            assert_true("writing an invalid Sell Mode value to the snapshot does not raise", set_ok, "")

            validation_present = True
            try:
                _ = snap_ws.Cells(manual_row, mode_col).Validation.Type
            except Exception:
                validation_present = False
            assert_true("Sell Mode validation removed from snapshot (Validation.Type absent)", not validation_present, "")

            assert_true("SNAP sheet tab is gray", rgb_tuple(snap_ws.Tab.Color) == (128, 128, 128), str(snap_ws.Tab.Color))

        # --- Index sheet ---
        idx_ws = None
        for w in wb.Worksheets:
            if w.Name == "CDS Snapshots":
                idx_ws = w
                break
        assert_true("CDS Snapshots index sheet exists", idx_ws is not None, "")

        if idx_ws is not None:
            last_row = idx_ws.Cells(idx_ws.Rows.Count, 1).End(-4162).Row  # xlUp
            assert_true("index has at least one data row", last_row >= 3, str(last_row))
            hlink_count = idx_ws.Cells(last_row, 2).Hyperlinks.Count
            assert_true("index row has a hyperlink in Sheet column", hlink_count > 0, str(hlink_count))
            target_raise_val = idx_ws.Cells(last_row, 4).Value
            total_proposed_val = idx_ws.Cells(last_row, 5).Value
            allocation_val = str(idx_ws.Cells(last_row, 7).Value or "")
            assert_true("index Target Raise populated (20000)",
                        target_raise_val is not None and abs(float(target_raise_val) - 20000) < 0.01,
                        str(target_raise_val))
            assert_true("index Total Proposed populated (nonzero)",
                        total_proposed_val is not None and float(total_proposed_val) != 0,
                        str(total_proposed_val))
            assert_true("index Allocation string non-empty", allocation_val.strip() != "", allocation_val)

        # --- Source sheet still active + fully functional ---
        active_name = wb.ActiveSheet.Name
        assert_true("source sheet reactivated after snapshot", active_name == ws.Name, active_name)

        ws.Activate()
        app.Run(qual + "AuditActiveCDSMath")
        ok, detail = read_audit_result(app, wb)
        assert_true("source sheet audit still 0 FAIL after snapshot", ok, detail)

        # --- ExportSnapshotToFile ---
        if snap_ws is not None:
            snap_ws.Activate()
            app.Run(qual + "ExportSnapshotToFile")
            snapshots_dir = os.path.join(workdir, "Snapshots")
            found_files = os.listdir(snapshots_dir) if os.path.isdir(snapshots_dir) else []
            xlsx_files = [f for f in found_files if f.lower().endswith(".xlsx")]
            assert_true("exported xlsx exists in Snapshots subfolder", len(xlsx_files) >= 1, str(found_files))

        # --- Refusal: SaveCDSSnapshot again while a SNAP sheet is active ---
        if snap_ws is not None:
            snap_ws.Activate()
            sheet_count_before_refusal = wb.Worksheets.Count
            app.Run(qual + "SaveCDSSnapshot")
            sheet_count_after_refusal = wb.Worksheets.Count
            assert_true("no new sheet created on refusal",
                        sheet_count_after_refusal == sheet_count_before_refusal,
                        f"{sheet_count_before_refusal} -> {sheet_count_after_refusal}")
            log_text = test_log_text(wb)
            assert_true("TestLog contains refusal text", "already a snapshot" in log_text, log_text[-400:])

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
