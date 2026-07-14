"""
Verification script for "New Session" (per-client workbook rollover).

Unlike the other verify_*.py scripts, this one uses TWO workbooks: a
"macro host" that holds the imported VBA project (standing in for
PERSONAL.XLSB / an add-in) and a separate "client" workbook that is the
one StartNewSession actually operates on (ActiveWorkbook). This matters
specifically for CanRunMacroById, whose ActiveClientSheet guard checks
`ActiveWorkbook Is ThisWorkbook` - a check that is only meaningful, and
only exercised here for the first time in this suite, when the running
macro's own workbook and the client workbook are genuinely different
objects (the real production topology).

Flow: start from a fresh, empty client workbook (no live report yet - the
way a client file is born) -> StartNewSession <fixture csv> (imports +
processes, no rollover since there is nothing live yet) -> StartNewSession
<fixture csv> again (freezes the first session to a SNAP sheet, removes it,
imports + processes a new session sheet) -> StartNewSession a third time
(same rollover again) -> assert SNAP/index accumulation and that exactly
one live report ever exists -> CanRunMacroById("new_session") is True on
the client workbook.

Uses the single-account fixture run_pipeline.py defaults to.
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


def live_report_sheets(wb):
    """Exact A2 idiom: UCase(Trim(...)) = 'ASSET CLASS'."""
    return [ws for ws in wb.Worksheets if str(ws.Cells(2, 1).Value or "").strip().upper() == "ASSET CLASS"]


def snap_sheets(wb):
    return [ws for ws in wb.Worksheets if ws.Name.startswith("SNAP ")]


def index_notes(wb):
    """Notes column (H) values from every data row of the CDS Snapshots index, or None if absent."""
    for w in wb.Worksheets:
        if w.Name == "CDS Snapshots":
            last_row = w.Cells(w.Rows.Count, 1).End(-4162).Row  # xlUp
            notes = []
            for r in range(3, last_row + 1):
                v = w.Cells(r, 8).Value
                if v:
                    notes.append(str(v))
            return notes
    return None


def main():
    global _excel_pid
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    _excel_pid = excel_pid(app)
    atexit.register(_kill_excel_pid)
    app.Visible = False
    app.DisplayAlerts = False

    workdir = tempfile.mkdtemp(prefix="cds_new_session_verify_")
    csv_copy = os.path.join(workdir, os.path.basename(FIXTURE))
    shutil.copy2(FIXTURE, csv_copy)

    macro_wb = None
    client_wb = None
    ok_overall = True
    try:
        # --- Macro host workbook: holds the VBA project only, never the data ---
        macro_wb = app.Workbooks.Add()
        macro_xlsm = os.path.join(workdir, "macro_host.xlsm")
        macro_wb.SaveAs(macro_xlsm, FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
        import_vba(macro_wb, workdir)
        macro_qual = f"'{macro_wb.Name}'!"

        # --- Client workbook: fresh, empty - no live report yet. This is how
        # --- a client file is born; StartNewSession must handle it without a
        # --- rollover. Never receives the VBA project itself.
        client_wb = app.Workbooks.Add()
        client_xlsm = os.path.join(workdir, "client_file.xlsm")
        client_wb.SaveAs(client_xlsm, FileFormat=52)

        def run_new_session():
            client_wb.Activate()
            app.Run(macro_qual + "StartNewSession", csv_copy)
            app.Calculate()

        # --- Call 1: fresh client workbook, nothing to freeze ---
        run_new_session()

        live1 = live_report_sheets(client_wb)
        assert_true("call1: exactly one live report", len(live1) == 1, [w.Name for w in live1])
        session1_name = live1[0].Name if live1 else None
        assert_true(
            "call1: session sheet name starts 'Session '",
            bool(session1_name) and session1_name.startswith("Session "),
            session1_name,
        )
        assert_true(
            "call1: no SNAP sheets yet", len(snap_sheets(client_wb)) == 0, [w.Name for w in snap_sheets(client_wb)]
        )

        # --- Call 2: must freeze + remove session1, import + process a new one ---
        run_new_session()

        live2 = live_report_sheets(client_wb)
        assert_true("call2: exactly one live report", len(live2) == 1, [w.Name for w in live2])
        session2_name = live2[0].Name if live2 else None
        assert_true(
            "call2: session sheet name starts 'Session '",
            bool(session2_name) and session2_name.startswith("Session "),
            session2_name,
        )
        # ("exactly one live report" above already proves session1 is gone -
        # otherwise there would be two. Sheet names can legitimately repeat
        # across rapid calls since the naming timestamp is minute-grained
        # and the old sheet is deleted before the new one is created.)

        snaps2 = snap_sheets(client_wb)
        assert_true("call2: exactly one SNAP sheet", len(snaps2) == 1, [w.Name for w in snaps2])
        if snaps2:
            snap_ws = snaps2[0]
            assert_true(
                "call2: SNAP sheet A2 marker",
                str(snap_ws.Cells(2, 1).Value) == "ASSET CLASS (SNAPSHOT)",
                str(snap_ws.Cells(2, 1).Value),
            )
            protected = False
            try:
                protected = bool(snap_ws.ProtectContents)
            except Exception:
                pass
            assert_true("call2: SNAP sheet protected", protected, str(protected))

        idx_notes2 = index_notes(client_wb)
        assert_true("call2: CDS Snapshots index sheet exists", idx_notes2 is not None, "")
        assert_true(
            "call2: index notes column contains 'session rollover'",
            idx_notes2 is not None and any("session rollover" in n for n in idx_notes2),
            idx_notes2,
        )

        # --- Call 3: must freeze + remove session2, import + process a third ---
        run_new_session()

        live3 = live_report_sheets(client_wb)
        assert_true("call3: exactly one live report", len(live3) == 1, [w.Name for w in live3])

        snaps3 = snap_sheets(client_wb)
        assert_true("call3: exactly two SNAP sheets", len(snaps3) == 2, [w.Name for w in snaps3])

        idx_notes3 = index_notes(client_wb)
        assert_true(
            "call3: index has two 'session rollover' rows",
            idx_notes3 is not None and sum(1 for n in idx_notes3 if "session rollover" in n) == 2,
            idx_notes3,
        )

        # --- CanRunMacroById on the client workbook (must be runnable, no state restriction) ---
        client_wb.Activate()
        client_wb.Worksheets(1).Activate()
        can_run = app.Run(macro_qual + "CanRunMacroById", "new_session", "")
        assert_true("CanRunMacroById('new_session') is True on client workbook", bool(can_run), str(can_run))

    except Exception:
        print("FULL TRACEBACK:\n" + traceback.format_exc())
        ok_overall = False
    finally:
        try:
            if client_wb is not None:
                client_wb.Close(False)
            if macro_wb is not None:
                macro_wb.Close(False)
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
