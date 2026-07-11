Attribute VB_Name = "CDS_Session"
Option Explicit

' ============================================================
' CDS Session (New Session)
'
' Purpose:
'   This workbook is a per-client history file, not a scratch pad: one
'   file per client, holding a "CDS Snapshots" index sheet, a growing
'   stack of frozen SNAP session sheets, and exactly ONE live report at
'   a time. StartNewSession is the single action that rolls a workbook
'   from "last meeting's live report" to "next meeting's live report":
'   freeze whatever is currently live into a dated SNAP sheet, remove
'   the now-stale live report, import a fresh Client Center CSV as a
'   new sheet, and process it. Filename discipline - which client this
'   workbook belongs to - stays the advisor's job; this macro never
'   touches the workbook's own file name.
'
'   Deleting the old live report is gated on its snapshot having been
'   verified to exist (the SNAP sheet count must actually go up after
'   SaveCDSSnapshot runs). If freezing fails or refuses for any reason,
'   the old report is left exactly as it was and the macro aborts - a
'   client's only copy of a plan is never removed on the strength of an
'   assumption that freezing "probably worked."
'
' Public entries:
'   StartNewSession - freeze + remove the current live report (if any),
'                      import a fresh Client Center CSV, process it.
' ============================================================

Private Const SNAP_PREFIX_SESSION As String = "SNAP "

Public Sub StartNewSession(Optional ByVal csvPath As String = "")
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevAlerts As Boolean

    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation
    prevAlerts = Application.DisplayAlerts

    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If LCase(wb.Name) = "personal.xlsb" Then
        MsgBox "Cannot run New Session on PERSONAL.XLSB. Open the client's workbook first, then run.", vbCritical, "CDS Trade Assistant"
        GoTo Done
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- Find the current live report, if any (exact A2 = "ASSET CLASS") ---
    Dim oldWs As Worksheet
    Dim liveCount As Long
    liveCount = FindLiveReportsSession(wb, oldWs)

    If liveCount > 1 Then
        MsgBox "This workbook has more than one live report sheet. Freeze or remove the extra report(s) manually, then run New Session again.", _
               vbExclamation, "CDS Trade Assistant"
        GoTo Done
    End If

    ' --- Resolve the CSV to import BEFORE freezing anything, so a cancelled
    '     or bad pick genuinely leaves the session unchanged ---
    If Trim(csvPath) = "" Then
        Dim picked As Variant
        picked = Application.GetOpenFilename( _
            FileFilter:="Client Center CSV (*.csv), *.csv", _
            Title:="Select Client Center CSV to import")

        If picked = False Then
            MsgBox "Session unchanged.", vbInformation, "CDS Trade Assistant"
            GoTo Done
        End If

        csvPath = CStr(picked)
    Else
        If Dir(csvPath) = "" Then
            MsgBox "CSV file not found: " & csvPath, vbExclamation, "CDS Trade Assistant"
            GoTo Done
        End If
    End If

    Dim didRollover As Boolean
    Dim snapNameFrozen As String
    didRollover = False

    If liveCount = 1 Then
        ' --- Freeze the current live report before ever deleting it ---
        Dim namesBefore As Collection
        Set namesBefore = CollectSheetNamesSession(wb)

        Dim snapCountBefore As Long
        snapCountBefore = CountSnapSheetsSession(wb)

        oldWs.Activate
        SaveCDSSnapshot noteText:="session rollover"

        Dim snapCountAfter As Long
        snapCountAfter = CountSnapSheetsSession(wb)

        If snapCountAfter <= snapCountBefore Then
            ' Snapshot was not verifiably created - never delete the report.
            MsgBox "Could not freeze the current report - nothing was changed.", vbExclamation, "CDS Trade Assistant"
            GoTo Done
        End If

        snapNameFrozen = FindNewSnapSheetNameSession(wb, namesBefore)

        Application.DisplayAlerts = False
        oldWs.Delete
        Application.DisplayAlerts = prevAlerts

        didRollover = True
    End If

    ' --- Import the CSV as a new sheet, after the last sheet in the workbook ---
    Dim csvWb As Workbook
    Set csvWb = Workbooks.Open(csvPath, ReadOnly:=True)

    csvWb.Worksheets(1).Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Application.CutCopyMode = False

    Dim newWs As Worksheet
    Set newWs = ActiveSheet   ' .Copy activates the new copy, now in wb

    csvWb.Close SaveChanges:=False

    Dim baseName As String
    baseName = "Session " & Format(Now, "yyyy-mm-dd hhnn")
    newWs.Name = UniqueSessionSheetName(wb, baseName)

    ' --- Process the freshly imported sheet in place ---
    newWs.Activate
    ProcessCDSHoldings_Lite

    Dim msg As String
    msg = "New session started: " & newWs.Name
    If didRollover Then
        msg = msg & vbCrLf & "Previous report frozen to: " & snapNameFrozen
    End If
    MsgBox msg, vbInformation, "CDS Trade Assistant"

Done:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "StartNewSession failed: " & Err.Number & " - " & Err.Description, vbExclamation, "CDS Trade Assistant"
    Resume Done
End Sub

' ============================================================
' Helpers
' ============================================================

' Counts sheets whose A2 is exactly "ASSET CLASS" (the same idiom used by
' IsProcessedReportLocal in CDS_MacroLauncher.bas and the guard in
' SaveCDSSnapshot). Returns the count; if exactly one is found, singleWs is
' set to it.
Private Function FindLiveReportsSession(wb As Workbook, ByRef singleWs As Worksheet) As Long
    Dim ws As Worksheet
    Dim n As Long
    n = 0

    For Each ws In wb.Worksheets
        If UCase(Trim(CStr(ws.Cells(2, 1).Value))) = "ASSET CLASS" Then
            n = n + 1
            Set singleWs = ws
        End If
    Next ws

    FindLiveReportsSession = n
End Function

Private Function CountSnapSheetsSession(wb As Workbook) As Long
    Dim ws As Worksheet
    Dim n As Long
    n = 0

    For Each ws In wb.Worksheets
        If Left$(ws.Name, Len(SNAP_PREFIX_SESSION)) = SNAP_PREFIX_SESSION Then n = n + 1
    Next ws

    CountSnapSheetsSession = n
End Function

Private Function CollectSheetNamesSession(wb As Workbook) As Collection
    Dim col As New Collection
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        col.Add ws.Name, ws.Name
    Next ws

    Set CollectSheetNamesSession = col
End Function

Private Function NameInCollectionSession(col As Collection, ByVal nameIn As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = col(nameIn)
    NameInCollectionSession = (Err.Number = 0)
    On Error GoTo 0
End Function

' Diff-based lookup (not position-based) so it is unaffected by exactly
' where SaveCDSSnapshot happens to place the new SNAP sheet.
Private Function FindNewSnapSheetNameSession(wb As Workbook, namesBefore As Collection) As String
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If Left$(ws.Name, Len(SNAP_PREFIX_SESSION)) = SNAP_PREFIX_SESSION Then
            If Not NameInCollectionSession(namesBefore, ws.Name) Then
                FindNewSnapSheetNameSession = ws.Name
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function UniqueSessionSheetName(wb As Workbook, baseName As String) As String
    Dim candidate As String
    Dim n As Long

    candidate = baseName
    n = 1

    Do While SheetExistsInWbSession(wb, candidate)
        n = n + 1
        Dim suffix As String
        suffix = " (" & n & ")"
        Dim maxBase As Long
        maxBase = 31 - Len(suffix)
        If maxBase < 1 Then maxBase = 1
        candidate = Left$(baseName, maxBase) & suffix
    Loop

    UniqueSessionSheetName = candidate
End Function

Private Function SheetExistsInWbSession(wb As Workbook, ByVal nameIn As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = wb.Sheets(nameIn)
    On Error GoTo 0
    SheetExistsInWbSession = Not s Is Nothing
End Function
