Attribute VB_Name = "CDS_Raise_Cash_Scenarios"
Option Explicit

' ============================================================
' CDS Raise Cash Scenarios v6.4
'
' v6.4 changes vs v6.3:
'   - Default CG rate reads from CDS_Settings ("DefaultCGRate")
'   - Scenario start col + stride read from CDS_Settings
'   - Pivot column renamed "Post-Sell %" (was "Allocation %"), title
'     "S{N} Post-Sell" (was "S{N} Allocation"). Distinguishes from
'     Post-Reb % column added by Buy Plans v1.2.
'   - PERSONAL.XLSB safety guard on all three public subs.
'
' Three subs:
'   AddRaiseCashScenarios -> initial setup
'   SpawnScenario          -> add S2, S3, S4...
'   RemoveScenario         -> delete a scenario by number
' ============================================================

Private Const SUMMARY_GAP As Long = 1

' Constants kept as fallbacks; runtime uses settings values
Private Const SCEN_START_COL_DEFAULT As Long = 16
Private Const SCEN_BLOCK_WIDTH As Long = 5
Private Const SCEN_BLOCK_STRIDE_DEFAULT As Long = 6

' ============================================================
' PUBLIC SUBS
' ============================================================

Sub AddRaiseCashScenarios()
Attribute AddRaiseCashScenarios.VB_ProcData.VB_Invoke_Func = "A\n14"
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB. Open the client's workbook first.", vbCritical
        GoTo Done
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim headerRow As Long, dataStart As Long, dataEnd As Long, totRow As Long
    If Not GetSheetState(ws, headerRow, dataStart, dataEnd, totRow) Then
        MsgBox "Run ProcessCDSHoldings first.", vbExclamation: GoTo Done
    End If

    PrepareCDSWorksheetForMacro ws

    If GetScenarioCount(ws) > 0 Then
        If MsgBox("Scenarios already exist. Reset and rebuild?", vbYesNo + vbQuestion) = vbNo Then GoTo Done
        ClearAllScenarios ws
    End If

    BuildScenarioBlock ws, 1, ScenStartCol(), headerRow, dataStart, dataEnd, totRow
    BuildSummary ws, headerRow, dataStart, dataEnd, totRow

    Application.Calculate

    GoTo Done
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
Done:
    On Error Resume Next
    If Not ws Is Nothing Then ApplyScenarioUXRulesToSheet ws
    On Error GoTo 0
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
End Sub

Sub SpawnScenario()
Attribute SpawnScenario.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB.", vbCritical
        GoTo Done
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim headerRow As Long, dataStart As Long, dataEnd As Long, totRow As Long
    If Not GetSheetState(ws, headerRow, dataStart, dataEnd, totRow) Then
        MsgBox "Run ProcessCDSHoldings first.", vbExclamation: GoTo Done
    End If

    Dim N As Long
    N = GetScenarioCount(ws)
    If N = 0 Then
        MsgBox "Run AddRaiseCashScenarios first.", vbExclamation: GoTo Done
    End If

    PrepareCDSWorksheetForMacro ws

    DeleteSummary ws

    Dim newScenNum As Long, newScenCol As Long
    newScenNum = N + 1
    newScenCol = ScenStartCol() + N * ScenStride()
    BuildScenarioBlock ws, newScenNum, newScenCol, headerRow, dataStart, dataEnd, totRow

    BuildSummary ws, headerRow, dataStart, dataEnd, totRow

    Application.Calculate

    GoTo Done
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
Done:
    On Error Resume Next
    If Not ws Is Nothing Then ApplyScenarioUXRulesToSheet ws
    On Error GoTo 0
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
End Sub

Sub RemoveScenario()
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    On Error GoTo ErrHandler

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB.", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim N As Long
    N = GetScenarioCount(ws)
    If N <= 1 Then
        MsgBox "Need at least one scenario remaining. Cannot remove.", vbExclamation
        Exit Sub
    End If

    Dim sn As String
    sn = InputBox("Which scenario to remove? (1 to " & N & ")", "Remove Scenario")
    If sn = "" Then Exit Sub
    If Not IsNumeric(sn) Then MsgBox "Not a number.", vbExclamation: Exit Sub

    Dim removeNum As Long
    removeNum = CLng(sn)
    If removeNum < 1 Or removeNum > N Then
        MsgBox "Invalid scenario number.", vbExclamation: Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim headerRow As Long, dataStart As Long, dataEnd As Long, totRow As Long
    If Not GetSheetState(ws, headerRow, dataStart, dataEnd, totRow) Then GoTo Done

    PrepareCDSWorksheetForMacro ws

    DeleteSummary ws

    Dim removeStartCol As Long, removeEndCol As Long
    removeStartCol = ScenStartCol() + (removeNum - 1) * ScenStride()
    removeEndCol = removeStartCol + ScenStride() - 1

    ws.Range(ws.Columns(removeStartCol), ws.Columns(removeEndCol)).Delete Shift:=xlShiftToLeft

    Dim i As Long, col As Long, titleRow As Long
    titleRow = headerRow - 1
    For i = removeNum To N - 1
        col = ScenStartCol() + (i - 1) * ScenStride()
        UpdateScenarioTitle ws, col, i, titleRow
        BuildScenarioPivot ws, i, col, dataStart, dataEnd, totRow
    Next i

    BuildSummary ws, headerRow, dataStart, dataEnd, totRow

    Application.Calculate

    GoTo Done
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
Done:
    On Error Resume Next
    If Not ws Is Nothing Then ApplyScenarioUXRulesToSheet ws
    On Error GoTo 0
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
End Sub

' ============================================================
' SETTINGS ACCESSORS (with fallbacks)
' ============================================================

Public Function ScenStartCol() As Long
    ScenStartCol = CLng(GetSettingNum("ScenarioStartCol", SCEN_START_COL_DEFAULT))
End Function

Public Function ScenStride() As Long
    ScenStride = CLng(GetSettingNum("ScenarioStride", SCEN_BLOCK_STRIDE_DEFAULT))
End Function

' ============================================================
' SCENARIO BLOCK BUILDER
' ============================================================

Private Sub BuildScenarioBlock(ws As Worksheet, scenNum As Long, startCol As Long, _
                                headerRow As Long, dataStart As Long, dataEnd As Long, totRow As Long)

    Dim titleRow As Long
    titleRow = headerRow - 1

    Dim rcCol As Long, gcCol As Long, nfcCol As Long, npcCol As Long, ilcCol As Long
    rcCol = startCol
    gcCol = startCol + 1
    nfcCol = startCol + 2
    npcCol = startCol + 3
    ilcCol = startCol + 4

    Dim rcL As String, gcL As String, nfcL As String, ilcL As String
    rcL = ColLetter(rcCol)
    gcL = ColLetter(gcCol)
    nfcL = ColLetter(nfcCol)
    ilcL = ColLetter(ilcCol)

    Dim hdrColor As Long
    hdrColor = GetScenarioHeaderColor(scenNum)

    UpdateScenarioTitle ws, startCol, scenNum, titleRow

    Dim subHeaders As Variant, hi As Long
    subHeaders = Array("Raise $", "Realized Gains", "New FMV", "New %", "Income Lost")
    For hi = 0 To 4
        With ws.Cells(headerRow, startCol + hi)
            .Value = subHeaders(hi)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .WrapText = True
            .Interior.Color = hdrColor
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next hi

    Dim r As Long
    For r = dataStart To dataEnd
        With ws.Cells(r, rcCol)
            .Value = 0
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
            .Interior.Color = RGB(197, 217, 241)
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        With ws.Cells(r, gcCol)
            .Formula = "=IF(E" & r & "=0,0," & rcL & r & "*(F" & r & "/E" & r & "))"
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        With ws.Cells(r, nfcCol)
            .Formula = "=E" & r & "-" & rcL & r
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        With ws.Cells(r, npcCol)
            .Formula = "=" & nfcL & r & "/" & nfcL & "$" & totRow
            .NumberFormat = "0%"
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        With ws.Cells(r, ilcCol)
            .Formula = "=IF(E" & r & "=0,0,(" & rcL & r & "/E" & r & ")*I" & r & ")"
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next r

    With ws.Cells(totRow, rcCol)
        .Formula = "=SUM(" & rcL & dataStart & ":" & rcL & dataEnd & ")"
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(totRow, gcCol)
        .Formula = "=SUM(" & gcL & dataStart & ":" & gcL & dataEnd & ")"
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(totRow, nfcCol)
        .Formula = "=SUM(" & nfcL & dataStart & ":" & nfcL & dataEnd & ")"
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(totRow, npcCol)
        .Value = 1
        .Font.Bold = True
        .NumberFormat = "0%"
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(totRow, ilcCol)
        .Formula = "=SUM(" & ilcL & dataStart & ":" & ilcL & dataEnd & ")"
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ws.Columns(ColLetter(rcCol)).ColumnWidth = 13
    ws.Columns(ColLetter(gcCol)).ColumnWidth = 14
    ws.Columns(ColLetter(nfcCol)).ColumnWidth = 13
    ws.Columns(ColLetter(npcCol)).ColumnWidth = 8
    ws.Columns(ColLetter(ilcCol)).ColumnWidth = 12
    ws.Columns(ColLetter(ilcCol + 1)).ColumnWidth = 2

    BuildScenarioPivot ws, scenNum, startCol, dataStart, dataEnd, totRow
End Sub

Private Sub UpdateScenarioTitle(ws As Worksheet, startCol As Long, scenNum As Long, titleRow As Long)
    Dim scenColor As Long
    scenColor = GetScenarioColor(scenNum)

    Dim baseTitle As String
    If scenNum = 1 Then
        baseTitle = "S1 PR"
    Else
        baseTitle = "S" & scenNum
    End If

    On Error Resume Next
    ws.Range(ws.Cells(titleRow, startCol), ws.Cells(titleRow, startCol + SCEN_BLOCK_WIDTH - 1)).UnMerge
    On Error GoTo 0

    ws.Range(ws.Cells(titleRow, startCol), ws.Cells(titleRow, startCol + SCEN_BLOCK_WIDTH - 1)).Merge

    Dim Q As String
    Q = Chr(34)
    Dim formulaStr As String
    formulaStr = "=" & Q & baseTitle & " | " & Q & _
        "&INDEX(C:C,$L$1)" & _
        "&" & Q & " | FMV $" & Q & _
        "&TEXT(INDEX(E:E,$L$1)/1000," & Q & "#,##0" & Q & ")" & _
        "&" & Q & "K | CB $" & Q & _
        "&TEXT(INDEX(H:H,$L$1)/1000," & Q & "#,##0" & Q & ")" & _
        "&" & Q & "K | " & Q & _
        "&TEXT(INDEX(G:G,$L$1)," & Q & "0%" & Q & ")"

    With ws.Cells(titleRow, startCol)
        .Formula = formulaStr
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Interior.Color = scenColor
    End With
End Sub

' ============================================================
' SCENARIO PIVOT
' Title: "S{N} Post-Sell"
' Columns: Asset Class | Post-Sell %
' (Buy Plans v1.2 augments with Post-Reb % as a 3rd column)
' ============================================================

Private Sub BuildScenarioPivot(ws As Worksheet, scenNum As Long, startCol As Long, _
                                dataStart As Long, dataEnd As Long, totRow As Long)

    Dim pivotStartRow As Long
    pivotStartRow = totRow + 2

    Dim npcCol As Long
    npcCol = startCol + 3
    Dim npcL As String
    npcL = ColLetter(npcCol)

    ws.Range(ws.Cells(pivotStartRow, startCol), ws.Cells(pivotStartRow + 30, startCol + 2)).Clear

    Dim classes As Object
    Set classes = CreateObject("Scripting.Dictionary")
    classes.CompareMode = vbTextCompare

    Dim r As Long, cls As String
    Dim classArr() As String, ci As Long
    ReDim classArr(0 To 50)
    ci = 0
    For r = dataStart To dataEnd
        cls = Trim(CStr(ws.Cells(r, 1).Value))
        If cls <> "" And cls <> "???" Then
            If Not classes.Exists(cls) Then
                classes.Add cls, cls
                classArr(ci) = cls
                ci = ci + 1
            End If
        End If
    Next r
    If ci = 0 Then Exit Sub
    ReDim Preserve classArr(0 To ci - 1)

    Dim classCount As Long
    classCount = ci

    On Error Resume Next
    ws.Range(ws.Cells(pivotStartRow, startCol), ws.Cells(pivotStartRow, startCol + 1)).UnMerge
    On Error GoTo 0
    ws.Range(ws.Cells(pivotStartRow, startCol), ws.Cells(pivotStartRow, startCol + 1)).Merge
    With ws.Cells(pivotStartRow, startCol)
        .Value = "S" & scenNum & " Post-Sell"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    Dim hdrRow As Long
    hdrRow = pivotStartRow + 1
    With ws.Cells(hdrRow, startCol)
        .Value = "Asset Class"
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(hdrRow, startCol + 1)
        .Value = "Post-Sell %"
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight
    End With

    Dim i As Long, dataR As Long
    Dim classRange As String, npcRange As String
    classRange = "$A$" & dataStart & ":$A$" & dataEnd
    npcRange = "$" & npcL & "$" & dataStart & ":$" & npcL & "$" & dataEnd

    For i = 0 To classCount - 1
        dataR = hdrRow + 1 + i
        With ws.Cells(dataR, startCol)
            .Value = classArr(i)
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        With ws.Cells(dataR, startCol + 1)
            .Formula = "=SUMIF(" & classRange & ",""" & classArr(i) & """," & npcRange & ")"
            .NumberFormat = "0%"
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlRight
        End With
    Next i

    Dim totRow2 As Long
    totRow2 = hdrRow + classCount + 1
    With ws.Cells(totRow2, startCol)
        .Value = "Grand Total"
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(totRow2, startCol + 1)
        .Formula = "=SUM(" & ColLetter(startCol + 1) & (hdrRow + 1) & ":" & ColLetter(startCol + 1) & (totRow2 - 1) & ")"
        .NumberFormat = "0%"
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight
    End With
End Sub

' ============================================================
' SUMMARY BUILDER
' ============================================================

Private Sub BuildSummary(ws As Worksheet, headerRow As Long, dataStart As Long, dataEnd As Long, totRow As Long)

    Dim N As Long
    N = GetScenarioCount(ws)
    If N = 0 Then Exit Sub

    Dim sumLabelCol As Long
    sumLabelCol = ScenStartCol() + N * ScenStride() + SUMMARY_GAP

    Dim titleRow As Long
    titleRow = headerRow - 1

    Dim mergeEndCol As Long
    mergeEndCol = sumLabelCol + N
    On Error Resume Next
    ws.Range(ws.Cells(titleRow, sumLabelCol), ws.Cells(titleRow, mergeEndCol)).UnMerge
    On Error GoTo 0
    ws.Range(ws.Cells(titleRow, sumLabelCol), ws.Cells(titleRow, mergeEndCol)).Merge
    With ws.Cells(titleRow, sumLabelCol)
        .Value = "SCENARIO SUMMARY"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Cells(headerRow, sumLabelCol)
        .Value = ""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    Dim sn As Long
    For sn = 1 To N
        With ws.Cells(headerRow, sumLabelCol + sn)
            If sn = 1 Then
                .Value = "S1 (Pro-Rata)"
            Else
                .Value = "S" & sn
            End If
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = GetScenarioHeaderColor(sn)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next sn

    Dim cgRateRow As Long
    cgRateRow = headerRow + 1
    With ws.Cells(cgRateRow, sumLabelCol)
        .Value = "Eff. CG Rate"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    Dim defaultCG As Double
    defaultCG = GetSettingNum("DefaultCGRate", 0.371)

    With ws.Cells(cgRateRow, sumLabelCol + 1)
        .Value = defaultCG
        .Font.Color = RGB(0, 0, 255)
        .Font.Bold = True
        .NumberFormat = "0.0%"
        .Interior.Color = RGB(197, 217, 241)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With
    On Error Resume Next
    ws.Cells(cgRateRow, sumLabelCol + 1).Comment.Delete
    ws.Cells(cgRateRow, sumLabelCol + 1).AddComment "Combined federal + state CG rate." & vbLf & "Set default in CDS_Settings."
    On Error GoTo 0
    Dim taxRateCell As String
    taxRateCell = "$" & ColLetter(sumLabelCol + 1) & "$" & cgRateRow

    For sn = 2 To N
        With ws.Cells(cgRateRow, sumLabelCol + sn)
            .Formula = "=" & taxRateCell
            .NumberFormat = "0.0%"
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlCenter
        End With
    Next sn

    Dim r As Long
    Dim cashFMVCell As String
    cashFMVCell = ""
    For r = dataStart To dataEnd
        If UCase(Trim(ws.Cells(r, 1).Value)) = "CASH" Then
            If cashFMVCell = "" Then cashFMVCell = "E" & r Else cashFMVCell = cashFMVCell & "+E" & r
        End If
    Next r
    If cashFMVCell = "" Then cashFMVCell = "0"

    Dim monthlyDistCell As String
    monthlyDistCell = ""
    For r = totRow + 1 To totRow + 30
        If InStr(1, CStr(ws.Cells(r, 9).Value), "MONTHLY DISTRIBUTION", vbTextCompare) > 0 Then
            monthlyDistCell = "$J$" & r: Exit For
        End If
    Next r

    Dim metrics() As Variant
    ReDim metrics(0 To 9)
    metrics(0) = "Target Raise"
    metrics(1) = "Total Raised"
    metrics(2) = "Shortfall / Overage"
    metrics(3) = "Realized Gains"
    metrics(4) = "Estimated Tax"
    metrics(5) = "Net After Tax"
    metrics(6) = "Income Lost / Year"
    metrics(7) = "New Annual Income"
    metrics(8) = "% Portfolio Sold"
    metrics(9) = "New Cash Balance"

    Dim mi As Long
    Dim metricStartRow As Long
    metricStartRow = cgRateRow + 1

    For mi = 0 To UBound(metrics)
        r = metricStartRow + mi
        With ws.Cells(r, sumLabelCol)
            .Value = metrics(mi)
            .Font.Bold = True
            .Interior.Color = RGB(255, 242, 204)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        For sn = 1 To N
            FillSummaryMetric ws, r, sumLabelCol + sn, mi, sn, sumLabelCol, dataStart, dataEnd, totRow, taxRateCell, cashFMVCell, monthlyDistCell
        Next sn
    Next mi

    If monthlyDistCell <> "" Then
        Dim runwayRow As Long
        runwayRow = metricStartRow + UBound(metrics) + 1
        With ws.Cells(runwayRow, sumLabelCol)
            .Value = "Months of Runway"
            .Font.Bold = True
            .Interior.Color = RGB(255, 242, 204)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        Dim cashCellLetter As String
        For sn = 1 To N
            cashCellLetter = ColLetter(sumLabelCol + sn) & (metricStartRow + 9)
            With ws.Cells(runwayRow, sumLabelCol + sn)
                .Formula = "=IF(" & monthlyDistCell & "=0,0," & cashCellLetter & "/" & monthlyDistCell & ")"
                .NumberFormat = "0.0"
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With
        Next sn
    End If

    ws.Columns(ColLetter(sumLabelCol)).ColumnWidth = 22
    For sn = 1 To N
        ws.Columns(ColLetter(sumLabelCol + sn)).ColumnWidth = 16
    Next sn

    Dim targetCell As String
    targetCell = "$" & ColLetter(sumLabelCol + 1) & "$" & metricStartRow

    Dim s1RaiseCol As Long
    s1RaiseCol = ScenStartCol()
    For r = dataStart To dataEnd
        ws.Cells(r, s1RaiseCol).Formula = "=" & targetCell & "*(E" & r & "/E$" & totRow & ")"
        ws.Cells(r, s1RaiseCol).Font.Color = RGB(0, 0, 0)
        ws.Cells(r, s1RaiseCol).Font.Bold = False
        ws.Cells(r, s1RaiseCol).Font.Italic = True
        ws.Cells(r, s1RaiseCol).Interior.Pattern = xlNone
    Next r

    With ws.Cells(metricStartRow, sumLabelCol + 1)
        .Value = 0
        .Font.Color = RGB(0, 0, 255)
        .Font.Bold = True
        .Interior.Color = RGB(197, 217, 241)
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    On Error Resume Next
    ws.Cells(metricStartRow, sumLabelCol + 1).Comment.Delete
    ws.Cells(metricStartRow, sumLabelCol + 1).AddComment "Type total $ to raise. S1 holdings auto-distribute pro-rata."
    On Error GoTo 0
End Sub

Private Sub FillSummaryMetric(ws As Worksheet, r As Long, sumCellCol As Long, _
                              mi As Long, sn As Long, sumLabelCol As Long, _
                              dataStart As Long, dataEnd As Long, totRow As Long, _
                              taxRateCell As String, cashFMVCell As String, monthlyDistCell As String)

    Dim scenCol As Long
    scenCol = ScenStartCol() + (sn - 1) * ScenStride()

    Dim rcL As String, gcL As String, ilcL As String
    rcL = ColLetter(scenCol)
    gcL = ColLetter(scenCol + 1)
    ilcL = ColLetter(scenCol + 4)

    Dim sumCL As String
    sumCL = ColLetter(sumCellCol)

    With ws.Cells(r, sumCellCol)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight

        Select Case mi
            Case 0
                If sn = 1 Then
                    .Value = 0
                    .Font.Color = RGB(0, 0, 255)
                    .Font.Bold = True
                    .Interior.Color = RGB(197, 217, 241)
                Else
                    .Formula = "=" & rcL & totRow
                End If
                .NumberFormat = "$#,##0;($#,##0);""-"""
            Case 1
                .Formula = "=" & rcL & totRow
                .NumberFormat = "$#,##0;($#,##0);""-"""
            Case 2
                .Formula = "=" & sumCL & (r - 1) & "-" & sumCL & (r - 2)
                .NumberFormat = "$#,##0;($#,##0);""-"""
            Case 3
                .Formula = "=" & gcL & totRow
                .NumberFormat = "$#,##0;($#,##0);""-"""
            Case 4
                .Formula = "=" & sumCL & (r - 1) & "*" & taxRateCell
                .NumberFormat = "$#,##0;($#,##0);""-"""
            Case 5
                .Formula = "=" & sumCL & (r - 4) & "-" & sumCL & (r - 1)
                .NumberFormat = "$#,##0;($#,##0);""-"""
                .Font.Bold = True
            Case 6
                .Formula = "=" & ilcL & totRow
                .NumberFormat = "$#,##0;($#,##0);""-"""
            Case 7
                .Formula = "=I" & totRow & "-" & sumCL & (r - 1)
                .NumberFormat = "$#,##0;($#,##0);""-"""
                .Font.Bold = True
            Case 8
                .Formula = "=IF(E" & totRow & "=0,0," & rcL & totRow & "/E" & totRow & ")"
                .NumberFormat = "0.0%"
            Case 9
                .Formula = "=" & cashFMVCell & "+" & sumCL & (r - 8)
                .NumberFormat = "$#,##0;($#,##0);""-"""
                .Font.Bold = True
        End Select
    End With
End Sub

' ============================================================
' UTILITY HELPERS
' ============================================================

Private Function GetSheetState(ws As Worksheet, ByRef headerRow As Long, ByRef dataStart As Long, _
                                ByRef dataEnd As Long, ByRef totRow As Long) As Boolean
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then GetSheetState = False: Exit Function
    dataStart = headerRow + 1
    totRow = FindTotalRow(ws, dataStart)
    If totRow = 0 Then GetSheetState = False: Exit Function
    dataEnd = totRow - 1
    GetSheetState = True
End Function

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim i As Long
    FindHeaderRow = 0
    For i = 1 To 5
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRow = i: Exit Function
        End If
    Next i
End Function

Private Function FindTotalRow(ws As Worksheet, dataStart As Long) As Long
    Dim i As Long
    FindTotalRow = 0
    For i = dataStart To dataStart + 200
        If Trim(CStr(ws.Cells(i, 1).Value)) = "" And ws.Cells(i, 5).Value <> "" Then
            FindTotalRow = i: Exit Function
        End If
    Next i
End Function

Public Function GetScenarioCount(ws As Worksheet) As Long
    Dim i As Long, col As Long
    GetScenarioCount = 0
    For i = 0 To 30
        col = ScenStartCol() + i * ScenStride()
        If InStr(1, CStr(ws.Cells(2, col).Value), "Raise $", vbTextCompare) > 0 Then
            GetScenarioCount = i + 1
        Else
            Exit Function
        End If
    Next i
End Function

Private Sub DeleteSummary(ws As Worksheet)
    Dim col As Long
    For col = 1 To 200
        If InStr(1, CStr(ws.Cells(1, col).Value), "SCENARIO SUMMARY", vbTextCompare) > 0 Then
            On Error Resume Next
            ws.Range(ws.Cells(1, col), ws.Cells(40, col + 20)).UnMerge
            ws.Range(ws.Cells(1, col), ws.Cells(40, col + 20)).Comment.Delete
            On Error GoTo 0
            ws.Range(ws.Cells(1, col), ws.Cells(40, col + 20)).Clear
            Exit Sub
        End If
    Next col
End Sub

Private Sub ClearAllScenarios(ws As Worksheet)
    On Error Resume Next
    ws.Range(ws.Cells(1, ScenStartCol()), ws.Cells(60, 250)).UnMerge
    On Error GoTo 0
    ws.Range(ws.Cells(1, ScenStartCol()), ws.Cells(60, 250)).Clear
End Sub

Private Function GetScenarioColor(scenNum As Long) As Long
    Select Case ((scenNum - 1) Mod 6)
        Case 0: GetScenarioColor = RGB(89, 89, 89)
        Case 1: GetScenarioColor = RGB(84, 130, 53)
        Case 2: GetScenarioColor = RGB(197, 90, 17)
        Case 3: GetScenarioColor = RGB(47, 85, 151)
        Case 4: GetScenarioColor = RGB(112, 48, 160)
        Case 5: GetScenarioColor = RGB(192, 0, 0)
    End Select
End Function

Private Function GetScenarioHeaderColor(scenNum As Long) As Long
    Select Case ((scenNum - 1) Mod 6)
        Case 0: GetScenarioHeaderColor = RGB(217, 217, 217)
        Case 1: GetScenarioHeaderColor = RGB(198, 239, 206)
        Case 2: GetScenarioHeaderColor = RGB(248, 203, 173)
        Case 3: GetScenarioHeaderColor = RGB(189, 215, 238)
        Case 4: GetScenarioHeaderColor = RGB(204, 192, 218)
        Case 5: GetScenarioHeaderColor = RGB(252, 195, 185)
    End Select
End Function

Private Function ColLetter(colNum As Long) As String
    ColLetter = Split(Columns(colNum).Address(, False), ":")(0)
End Function
