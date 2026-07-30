Attribute VB_Name = "CDS_Snapshots"
Option Explicit

' ============================================================
' CDS Snapshots (W9)
'
' Purpose:
'   Freezes the active processed report into a dated, values-only
'   archive sheet so the workbook still means something months
'   later ("how did we raise the 200k last time", "what was the
'   allocation drift in March"). Formulas and workbook links make the
'   working sheet mutable; a snapshot is a deliberate, non-live copy.
'
' Public entries:
'   SaveCDSSnapshot     - copy the active processed report, freeze
'                          it to values+formats, log it to the
'                          "CDS Snapshots" index sheet.
'   ExportSnapshotToFile - copy the active SNAP sheet out to its
'                          own standalone .xlsx file on disk.
'
' A snapshot sheet is deliberately NOT a processed report as far as
' every other macro (and the launcher's state detection) is
' concerned: its A2 marker reads "ASSET CLASS (SNAPSHOT)" instead
' of "ASSET CLASS", so IsProcessedReport-style exact-match checks
' elsewhere in the codebase ignore it without needing to be taught
' about snapshots explicitly.
'
' Called from CDS_Trade_Email.GenerateTradeEmail after a successful
' send (Outlook or AutomationMode preview): "Save a snapshot of
' this plan for the record?"
' ============================================================

Private Const SNAPSHOT_PREFIX As String = "SNAP "
Private Const SNAPSHOT_MARKER As String = "ASSET CLASS (SNAPSHOT)"
Private Const INDEX_SHEET_NAME As String = "CDS Snapshots"
Private Const INDEX_TITLE As String = "CDS SNAPSHOT INDEX"

' ============================================================
' SAVE SNAPSHOT
' ============================================================

Public Sub SaveCDSSnapshot(Optional ByVal noteText As String = "")
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation

    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim srcWs As Worksheet
    Set srcWs = ActiveSheet

    ' No snapshot-of-snapshot: a SNAP sheet's own A2 marker differs from
    ' "ASSET CLASS", which would already fail the guard below, but this
    ' check gives a clearer, dedicated refusal message.
    If Left$(srcWs.Name, Len(SNAPSHOT_PREFIX)) = SNAPSHOT_PREFIX Then
        MsgBox "This sheet is already a snapshot. Activate the live report sheet first.", vbExclamation, "CDS Snapshot"
        Exit Sub
    End If

    If UCase(Trim(CStr(srcWs.Cells(2, 1).Value))) <> "ASSET CLASS" Then
        MsgBox "Activate a processed CDS report sheet first.", vbExclamation, "CDS Snapshot"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- Gather index data from the LIVE source sheet before freezing ---
    Dim acctNum As String, acctName As String
    acctNum = Trim(CStr(srcWs.Cells(1, 1).Value))
    acctName = Trim(CStr(srcWs.Cells(1, 2).Value))

    Dim targetRaise As Variant
    targetRaise = FindPlanTargetRaiseSnap(srcWs)

    Dim totalProposed As Variant
    totalProposed = FindTotalProposedSnap(srcWs)

    Dim scenarioCount As Long
    scenarioCount = GetScenarioCount(srcWs)

    Dim allocationStr As String
    allocationStr = BuildAllocationString(srcWs)

    ' --- Name + copy ---
    Dim snapName As String
    snapName = UniqueSheetName(srcWs.Parent, BuildSnapshotName(acctNum, acctName))

    srcWs.Copy After:=srcWs
    Application.CutCopyMode = False

    Dim snapWs As Worksheet
    Set snapWs = ActiveSheet   ' .Copy activates the new copy
    snapWs.Name = snapName

    FreezeSheetToValues snapWs
    RemoveValidationsFromSheet snapWs

    ' Distinguish snapshot from processed report so no other macro (and the
    ' launcher's WorkflowStateCode / ThisWorkbook's selection-change context
    ' handler, both of which do exact A2="ASSET CLASS" matches) treats this
    ' sheet as a live report.
    snapWs.Cells(2, 1).Value = SNAPSHOT_MARKER

    Dim stampCell As Range
    Set stampCell = FirstEmptyRow1Cell(snapWs)
    stampCell.Value = "SNAPSHOT " & Format(Now, "mm/dd/yyyy hh:nn") & " (frozen copy - do not trade from this sheet)"

    snapWs.Tab.Color = RGB(128, 128, 128)

    ' View-only intent: lock every cell (including any workbench/routing
    ' inputs that were unlocked on the live sheet) before protecting.
    snapWs.Cells.Locked = True
    On Error Resume Next
    snapWs.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0

    AppendSnapshotIndexRow srcWs.Parent, snapWs, acctName, acctNum, targetRaise, totalProposed, scenarioCount, allocationStr, noteText

    srcWs.Activate

    MsgBox "Snapshot saved: " & snapName, vbInformation, "CDS Snapshot"

Done:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "SaveCDSSnapshot failed: " & Err.Number & " - " & Err.Description, vbExclamation, "CDS Snapshot"
    Resume Done
End Sub

Private Function BuildSnapshotName(acctNum As String, acctName As String) As String
    Dim stamp As String
    stamp = SNAPSHOT_PREFIX & Format(Now, "mmdd-hhnn")

    Dim acctFragment As String
    acctFragment = Trim(acctNum)
    If acctFragment = "" Then acctFragment = Trim(acctName)

    Dim nameOut As String
    If acctFragment <> "" Then
        nameOut = stamp & " " & acctFragment
    Else
        nameOut = stamp
    End If

    If Len(nameOut) > 31 Then nameOut = Left$(nameOut, 31)

    Do While Len(nameOut) > 0 And Right$(nameOut, 1) = " "
        nameOut = Left$(nameOut, Len(nameOut) - 1)
    Loop

    BuildSnapshotName = nameOut
End Function

Private Function UniqueSheetName(wb As Workbook, baseName As String) As String
    Dim candidate As String
    Dim n As Long

    candidate = baseName
    n = 1

    Do While SheetExistsSnap(wb, candidate)
        n = n + 1
        Dim suffix As String
        suffix = "-" & n
        Dim maxBase As Long
        maxBase = 31 - Len(suffix)
        If maxBase < 1 Then maxBase = 1
        candidate = Left$(baseName, maxBase) & suffix
    Loop

    UniqueSheetName = candidate
End Function

Private Function SheetExistsSnap(wb As Workbook, nameIn As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = wb.Sheets(nameIn)
    On Error GoTo 0
    SheetExistsSnap = Not s Is Nothing
End Function

' Freezes every formula and workbook-linked value to a plain value.
' Run twice so copied cells are fully detached from mutable workbook state.
Private Sub FreezeSheetToValues(ws As Worksheet)
    ' A live PivotTable anywhere in the UsedRange blocks a bulk
    ' Range.Value=Range.Value re-write for the WHOLE range (not just the
    ' pivot's own cells), and a copied sheet's pivot stays bound to the
    ' SAME cache/source data as the original - which then breaks later
    ' refresh/audit on the source sheet with a "protected sheet contains
    ' another PivotTable report" error. Disconnect every pivot on this
    ' sheet first (copy its own displayed grid back onto itself as
    ' values-only, which keeps the existing cell formatting intact but
    ' un-binds it from the PivotTable object) before the generic freeze.
    DisconnectPivotTables ws

    Dim ur As Range
    Set ur = ws.UsedRange
    ur.Value = ur.Value

    Set ur = ws.UsedRange
    ur.Value = ur.Value
End Sub

Private Sub DisconnectPivotTables(ws As Worksheet)
    Dim pivotCount As Long
    On Error Resume Next
    pivotCount = ws.PivotTables.Count
    On Error GoTo 0

    If pivotCount = 0 Then Exit Sub

    ' Capture every TableRange2 first (a second pass over a live
    ' PivotTables collection, after the first entry has already been
    ' disconnected, is not reliable).
    Dim pivotRanges() As Range
    ReDim pivotRanges(1 To pivotCount)

    Dim i As Long
    For i = 1 To pivotCount
        Set pivotRanges(i) = ws.PivotTables(i).TableRange2
    Next i

    For i = 1 To pivotCount
        pivotRanges(i).Copy
        pivotRanges(i).PasteSpecial Paste:=xlPasteValues
    Next i
    Application.CutCopyMode = False
End Sub

Private Sub RemoveValidationsFromSheet(ws As Worksheet)
    On Error Resume Next
    ws.Cells.Validation.Delete
    On Error GoTo 0
End Sub

Private Function FirstEmptyRow1Cell(ws As Worksheet) As Range
    Dim c As Long
    For c = 3 To 60
        If Trim(CStr(ws.Cells(1, c).Value)) = "" Then
            Set FirstEmptyRow1Cell = ws.Cells(1, c)
            Exit Function
        End If
    Next c
    Set FirstEmptyRow1Cell = ws.Cells(1, 60)
End Function

Private Function FindPlanTargetRaiseSnap(ws As Worksheet) As Variant
    Dim c As Range
    Set c = FindCellExactSnap(ws, "Plan Target Raise")

    If c Is Nothing Then
        FindPlanTargetRaiseSnap = Empty
    Else
        FindPlanTargetRaiseSnap = ws.Cells(c.Row, c.Column + 1).Value
    End If
End Function

Private Function FindTotalProposedSnap(ws As Worksheet) As Variant
    Dim c As Range
    Set c = FindCellExactSnap(ws, "Total Proposed")

    If c Is Nothing Then
        FindTotalProposedSnap = Empty
    Else
        FindTotalProposedSnap = ws.Cells(c.Row, c.Column + 1).Value
    End If
End Function

' Best-effort compact allocation string from the base pivot at M2
' ("AllocPivot"): "CLASS pct / CLASS pct / ...". Blank if the pivot area
' isn't there (e.g. sheet built before a sell workbench/pivot existed).
Private Function BuildAllocationString(ws As Worksheet) As String
    On Error GoTo Fallback

    Dim r As Long, cls As String, pct As Variant, parts As String
    r = 3   ' M2 = pivot header row; data starts M3

    Do While r <= 200
        cls = Trim(CStr(ws.Cells(r, 13).Value))   ' col M
        If cls = "" Or UCase(cls) = "GRAND TOTAL" Then Exit Do

        pct = ws.Cells(r, 14).Value               ' col N
        If parts <> "" Then parts = parts & " / "
        If IsNumeric(pct) Then
            parts = parts & cls & " " & Format(CDbl(pct), "0%")
        Else
            parts = parts & cls
        End If

        r = r + 1
    Loop

    BuildAllocationString = parts
    Exit Function

Fallback:
    BuildAllocationString = ""
End Function

Private Function FindCellExactSnap(ws As Worksheet, textValue As String) As Range
    On Error Resume Next
    Set FindCellExactSnap = ws.Cells.Find(What:=textValue, _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          SearchOrder:=xlByRows, _
                                          SearchDirection:=xlNext, _
                                          MatchCase:=False)
    On Error GoTo 0
End Function

' ============================================================
' SNAPSHOT INDEX
' ============================================================

Private Sub AppendSnapshotIndexRow(wb As Workbook, snapWs As Worksheet, acctName As String, acctNum As String, _
                                   targetRaise As Variant, totalProposed As Variant, scenarioCount As Long, _
                                   allocationStr As String, Optional ByVal noteText As String = "")
    Dim idxWs As Worksheet
    Set idxWs = GetOrCreateSnapshotIndexSheet(wb)

    Dim r As Long
    r = idxWs.Cells(idxWs.Rows.Count, 1).End(xlUp).Row + 1
    If r < 3 Then r = 3

    With idxWs.Cells(r, 1)
        .Value = Now
        .NumberFormat = "mm/dd/yyyy hh:mm"
    End With

    idxWs.Hyperlinks.Add Anchor:=idxWs.Cells(r, 2), Address:="", _
                         SubAddress:="'" & snapWs.Name & "'!A1", TextToDisplay:=snapWs.Name

    idxWs.Cells(r, 3).Value = Trim(acctName & " " & acctNum)

    If IsNumeric(targetRaise) Then
        idxWs.Cells(r, 4).Value = CDbl(targetRaise)
        idxWs.Cells(r, 4).NumberFormat = "$#,##0"
    End If

    If IsNumeric(totalProposed) Then
        idxWs.Cells(r, 5).Value = CDbl(totalProposed)
        idxWs.Cells(r, 5).NumberFormat = "$#,##0"
    End If

    idxWs.Cells(r, 6).Value = scenarioCount
    idxWs.Cells(r, 7).Value = allocationStr

    ' Column 8 (Notes): caller-supplied note (e.g. "session rollover") when
    ' given, otherwise left blank for the advisor to fill in by hand.
    If Trim(noteText) <> "" Then idxWs.Cells(r, 8).Value = noteText

    idxWs.Range(idxWs.Cells(r, 1), idxWs.Cells(r, 8)).Borders.LineStyle = xlContinuous
End Sub

Private Function GetOrCreateSnapshotIndexSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(INDEX_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = INDEX_SHEET_NAME
        BuildSnapshotIndexHeader ws
    End If

    Set GetOrCreateSnapshotIndexSheet = ws
End Function

Private Sub BuildSnapshotIndexHeader(ws As Worksheet)
    With ws.Cells(1, 1)
        .Value = INDEX_TITLE
        .Font.Bold = True
        .Font.Size = 14
    End With

    Dim headers As Variant
    headers = Array("Saved", "Sheet", "Account", "Target Raise", "Total Proposed", "Scenarios", "Allocation", "Notes")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Interior.Color = RGB(189, 215, 238)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next i

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 26
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 44
    ws.Columns("H").ColumnWidth = 32
End Sub

' ============================================================
' EXPORT SNAPSHOT TO FILE
' ============================================================

Public Sub ExportSnapshotToFile()
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts

    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim srcWs As Worksheet
    Set srcWs = ActiveSheet

    If Left$(srcWs.Name, Len(SNAPSHOT_PREFIX)) <> SNAPSHOT_PREFIX Then
        MsgBox "Activate a SNAP snapshot sheet first.", vbExclamation, "CDS Snapshot"
        Exit Sub
    End If

    Dim folderPath As String
    folderPath = SnapshotExportFolder(srcWs.Parent)
    EnsureFolderExists folderPath

    Dim filePath As String
    filePath = UniqueFilePath(folderPath, SanitizeFileName(srcWs.Name), ".xlsx")

    Application.DisplayAlerts = False

    srcWs.Copy   ' new standalone workbook containing just this sheet
    Application.CutCopyMode = False

    Dim newWb As Workbook
    Set newWb = ActiveWorkbook
    newWb.SaveAs Filename:=filePath, FileFormat:=51   ' xlOpenXMLWorkbook (.xlsx)
    newWb.Close False

    Application.DisplayAlerts = prevAlerts

    MsgBox "Snapshot exported to:" & vbCrLf & filePath, vbInformation, "CDS Snapshot"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = prevAlerts
    MsgBox "ExportSnapshotToFile failed: " & Err.Number & " - " & Err.Description, vbExclamation, "CDS Snapshot"
End Sub

Private Function SnapshotExportFolder(wb As Workbook) As String
    Dim basePath As String
    If wb.Path = "" Then
        basePath = Environ$("USERPROFILE") & "\Documents"
    Else
        basePath = wb.Path
    End If
    SnapshotExportFolder = basePath & "\Snapshots"
End Function

Private Sub EnsureFolderExists(folderPath As String)
    On Error Resume Next
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function SanitizeFileName(nameIn As String) As String
    Dim badChars As String
    badChars = "\/:*?""<>|"

    Dim outStr As String
    Dim i As Long
    outStr = nameIn
    For i = 1 To Len(badChars)
        outStr = Replace(outStr, Mid$(badChars, i, 1), "_")
    Next i

    SanitizeFileName = outStr
End Function

Private Function UniqueFilePath(folderPath As String, baseName As String, ext As String) As String
    Dim candidate As String
    Dim n As Long

    candidate = folderPath & "\" & baseName & ext
    n = 1

    Do While Dir(candidate) <> ""
        n = n + 1
        candidate = folderPath & "\" & baseName & "-" & n & ext
    Loop

    UniqueFilePath = candidate
End Function
