Attribute VB_Name = "CDS_ScenarioUX"
Option Explicit

' ============================================================
' CDS Scenario UX Rules v1.0
'
' Purpose:
'   Enforces the intended scenario model:
'
'       S1 = Pro-Rata Baseline
'       S2+ = Manual sell scenarios
'
' What this does:
'   - Renames/styles S1 as "S1 Pro-Rata Baseline"
'   - Rebuilds S1 Raise $ formulas from the Summary Target Raise
'   - Makes S1 Raise $ cells visibly non-manual
'   - Locks S1 Raise $ cells
'   - Unlocks S2+ Raise $ cells
'   - Unlocks legitimate input cells:
'       * S1 Target Raise
'       * S1 Effective CG Rate
'       * Tax/manual input cells in column J
'       * Buy-plan ticker/amount fields
'   - Protects the worksheet so users cannot overwrite S1 formulas
'   - Uses UserInterfaceOnly:=True so macros can still edit the sheet
'
' Important:
'   Use the CDS Trade Assistant to run scenario/buy-plan steps so this
'   protection gets refreshed automatically.
' ============================================================

Private Const CDS_PROTECT_PASSWORD As String = ""
Private Const S1_TITLE_TEXT As String = "S1 Pro-Rata Baseline"
Private Const S1_HELP_TEXT As String = "S1 is the pro-rata baseline. Do not type in these cells. Enter the total raise amount in Scenario Summary > S1 Target Raise. Use Spawn Scenario for manual sells."

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub ApplyScenarioUXRules()
    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Click into the client workbook first. PERSONAL.XLSB is active.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ApplyScenarioUXRulesToSheet ws
    Exit Sub

ErrHandler:
    MsgBox "ApplyScenarioUXRules failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Public Sub PrepareCDSWorksheetForMacro(Optional ws As Worksheet)
    On Error Resume Next

    If ws Is Nothing Then Set ws = ActiveSheet

    If Not ws Is Nothing Then
        ws.Unprotect Password:=CDS_PROTECT_PASSWORD
    End If

    On Error GoTo 0
End Sub

Public Sub ProtectCDSWorksheet(Optional ws As Worksheet)
    On Error Resume Next

    If ws Is Nothing Then Set ws = ActiveSheet

    If Not ws Is Nothing Then
        ws.Protect Password:=CDS_PROTECT_PASSWORD, _
                   DrawingObjects:=False, _
                   Contents:=True, _
                   Scenarios:=False, _
                   UserInterfaceOnly:=True, _
                   AllowFormattingCells:=True, _
                   AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True, _
                   AllowSorting:=True, _
                   AllowFiltering:=True
    End If

    On Error GoTo 0
End Sub

Public Sub ResetS1ProRataBaseline()
    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Click into the client workbook first. PERSONAL.XLSB is active.", vbExclamation
        Exit Sub
    End If

    ApplyScenarioUXRulesToSheet ActiveSheet

    MsgBox "S1 reset to pro-rata baseline." & vbCrLf & vbCrLf & _
           "Enter the total raise amount in Scenario Summary > S1 Target Raise." & vbCrLf & _
           "Use Spawn Scenario for manual sells.", _
           vbInformation, "S1 Pro-Rata Baseline"

    Exit Sub

ErrHandler:
    MsgBox "ResetS1ProRataBaseline failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ============================================================
' CORE APPLY LOGIC
' ============================================================

Public Sub ApplyScenarioUXRulesToSheet(ws As Worksheet)
    On Error GoTo ErrHandler

    If ws Is Nothing Then Exit Sub
    If Not IsProcessedCDSReport(ws) Then Exit Sub

    Dim headerRow As Long
    Dim dataStart As Long
    Dim dataEnd As Long
    Dim totRow As Long

    If Not GetProcessedSheetState(ws, headerRow, dataStart, dataEnd, totRow) Then Exit Sub

    Dim scenCount As Long
    scenCount = CountScenariosUX(ws)

    If scenCount = 0 Then Exit Sub

    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    PrepareCDSWorksheetForMacro ws

    ' Default all cells locked. Then explicitly unlock only true inputs.
    ws.Cells.Locked = True

    ApplyS1Title ws, headerRow
    ApplyS1ProRataFormulas ws, dataStart, dataEnd, totRow
    StyleScenarioInputs ws, dataStart, dataEnd, scenCount
    UnlockScenarioSummaryInputs ws
    UnlockTaxInputs ws, totRow
    UnlockBuyPlanInputs ws, totRow, scenCount
    UnlockSellWorkbenchInputs ws, dataStart, dataEnd
    UnlockRoutingInputs ws

    Application.Calculate

    ProtectCDSWorksheet ws

Done:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Exit Sub

ErrHandler:
    MsgBox "ApplyScenarioUXRulesToSheet failed: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Done
End Sub

' ============================================================
' S1 TITLE + FORMULAS
' ============================================================

Private Sub ApplyS1Title(ws As Worksheet, headerRow As Long)
    On Error GoTo SafeExit

    Dim titleRow As Long
    titleRow = headerRow - 1

    If titleRow < 1 Then Exit Sub

    Dim s1StartCol As Long
    s1StartCol = ScenStartCol()

    Dim s1EndCol As Long
    s1EndCol = s1StartCol + 4

    On Error Resume Next
    ws.Range(ws.Cells(titleRow, s1StartCol), ws.Cells(titleRow, s1EndCol)).UnMerge
    On Error GoTo SafeExit

    ws.Range(ws.Cells(titleRow, s1StartCol), ws.Cells(titleRow, s1EndCol)).Merge

    Dim Q As String
    Q = Chr(34)

    Dim formulaStr As String

    ' Mirrors the existing title pattern, but with clearer S1 naming.
    formulaStr = "=" & Q & S1_TITLE_TEXT & " | " & Q & _
                 "&INDEX(C:C,$L$1)" & _
                 "&" & Q & " | FMV $" & Q & _
                 "&TEXT(INDEX(E:E,$L$1)/1000," & Q & "#,##0" & Q & ")" & _
                 "&" & Q & "K | CB $" & Q & _
                 "&TEXT(INDEX(H:H,$L$1)/1000," & Q & "#,##0" & Q & ")" & _
                 "&" & Q & "K | " & Q & _
                 "&TEXT(INDEX(G:G,$L$1)," & Q & "0%" & Q & ")"

    With ws.Cells(titleRow, s1StartCol)
        .Formula = formulaStr
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(89, 89, 89)
        .Locked = True
    End With

SafeExit:
End Sub

Private Sub ApplyS1ProRataFormulas(ws As Worksheet, dataStart As Long, dataEnd As Long, totRow As Long)
    Dim summaryCol As Long
    summaryCol = FindScenarioSummaryCol(ws)

    If summaryCol = 0 Then Exit Sub

    Dim targetRow As Long
    targetRow = FindLabelRowInColumn(ws, summaryCol, "Target Raise")

    If targetRow = 0 Then Exit Sub

    Dim targetCell As String
    targetCell = "$" & ColLetterUX(summaryCol + 1) & "$" & targetRow

    Dim s1RaiseCol As Long
    s1RaiseCol = ScenStartCol()

    Dim r As Long

    For r = dataStart To dataEnd
        With ws.Cells(r, s1RaiseCol)
            .Formula = "=" & targetCell & "*(E" & r & "/E$" & totRow & ")"
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Font.Color = RGB(80, 80, 80)
            .Font.Bold = False
            .Font.Italic = True
            .Interior.Color = RGB(242, 242, 242)
            .Locked = True

            On Error Resume Next
            .Comment.Delete
            .AddComment S1_HELP_TEXT
            .Comment.Visible = False
            On Error GoTo 0
        End With
    Next r

    ' Also make the S1 Raise $ header visually non-manual.
    With ws.Cells(2, s1RaiseCol)
        .Value = "Raise $"
        .Font.Bold = True
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
        .Locked = True

        On Error Resume Next
        .Comment.Delete
        .AddComment S1_HELP_TEXT
        .Comment.Visible = False
        On Error GoTo 0
    End With
End Sub

' ============================================================
' SCENARIO INPUT STYLING
' ============================================================

Private Sub StyleScenarioInputs(ws As Worksheet, dataStart As Long, dataEnd As Long, scenCount As Long)
    Dim sn As Long
    Dim scenCol As Long
    Dim r As Long

    For sn = 1 To scenCount
        scenCol = ScenStartCol() + (sn - 1) * ScenStride()

        If sn = 1 Then
            ' S1 stays locked/pro-rata.
            For r = dataStart To dataEnd
                With ws.Cells(r, scenCol)
                    .Font.Color = RGB(80, 80, 80)
                    .Font.Bold = False
                    .Font.Italic = True
                    .Interior.Color = RGB(242, 242, 242)
                    .Locked = True
                End With
            Next r

            With ws.Cells(2, scenCol)
                .Interior.Color = RGB(217, 217, 217)
                .Font.Color = RGB(80, 80, 80)
                .Locked = True
            End With

        Else
            ' S2+ are true manual raise scenarios.
            For r = dataStart To dataEnd
                With ws.Cells(r, scenCol)
                    .Font.Color = RGB(0, 0, 255)
                    .Font.Bold = True
                    .Font.Italic = False
                    .Interior.Color = RGB(197, 217, 241)
                    .NumberFormat = "$#,##0;($#,##0);""-"""
                    .Locked = False

                    On Error Resume Next
                    .Comment.Delete
                    .AddComment "Manual scenario. Enter the dollar amount to raise from this holding."
                    .Comment.Visible = False
                    On Error GoTo 0
                End With
            Next r

            With ws.Cells(2, scenCol)
                .Interior.Color = RGB(197, 217, 241)
                .Font.Color = RGB(0, 0, 0)
                .Locked = True
            End With
        End If
    Next sn
End Sub

' ============================================================
' UNLOCK LEGITIMATE INPUT AREAS
' ============================================================

Private Sub UnlockScenarioSummaryInputs(ws As Worksheet)
    Dim summaryCol As Long
    summaryCol = FindScenarioSummaryCol(ws)

    If summaryCol = 0 Then Exit Sub

    Dim cgRateRow As Long
    Dim targetRaiseRow As Long

    cgRateRow = FindLabelRowInColumn(ws, summaryCol, "Eff. CG Rate")
    targetRaiseRow = FindLabelRowInColumn(ws, summaryCol, "Target Raise")

    ' S1 Effective CG Rate input.
    If cgRateRow > 0 Then
        With ws.Cells(cgRateRow, summaryCol + 1)
            .Locked = False
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
            .Interior.Color = RGB(197, 217, 241)
        End With
    End If

    ' S1 Target Raise input. This is the only S1 sell input.
    If targetRaiseRow > 0 Then
        With ws.Cells(targetRaiseRow, summaryCol + 1)
            .Locked = False
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
            .Interior.Color = RGB(197, 217, 241)

            On Error Resume Next
            .Comment.Delete
            .AddComment "Enter the total dollar amount to raise for S1. S1 will automatically sell pro-rata across holdings."
            .Comment.Visible = False
            On Error GoTo 0
        End With
    End If
End Sub

Private Sub UnlockTaxInputs(ws As Worksheet, totRow As Long)
    Dim r As Long
    Dim labelText As String

    For r = totRow + 1 To totRow + 50
        labelText = UCase(Trim(CStr(ws.Cells(r, 9).Value)))

        If labelText = "" Then
            ' Keep scanning; there may be gaps.
        ElseIf InStr(1, labelText, "MONTHLY DISTRIBUTION", vbTextCompare) > 0 Or _
               InStr(1, labelText, "REALIZED GAINS", vbTextCompare) > 0 Or _
               InStr(1, labelText, "TAXABLE DIVIDENDS", vbTextCompare) > 0 Or _
               InStr(1, labelText, "NON-TAXABLE DIVIDENDS", vbTextCompare) > 0 Then

            With ws.Cells(r, 10)
                .Locked = False
                .Font.Color = RGB(0, 0, 255)
                .Font.Bold = True
                .Interior.Color = RGB(255, 242, 204)
            End With
        End If
    Next r
End Sub

Private Sub UnlockBuyPlanInputs(ws As Worksheet, totRow As Long, scenCount As Long)
    On Error GoTo SafeExit

    Dim sn As Long
    Dim scenCol As Long
    Dim grandRow As Long
    Dim inputStart As Long
    Dim inputEnd As Long
    Dim r As Long

    For sn = 1 To scenCount
        scenCol = ScenStartCol() + (sn - 1) * ScenStride()
        grandRow = FindScenarioPivotGrandTotalRowUX(ws, scenCol, totRow)

        If grandRow > 0 Then
    Dim buyPlanHeaderRow As Long
    buyPlanHeaderRow = FindBuyPlanHeaderRow_UX(ws, scenCol, grandRow)

    If buyPlanHeaderRow > 0 Then
        inputStart = buyPlanHeaderRow + 2
        inputEnd = inputStart + BuyPlanRows() - 1
                For r = inputStart To inputEnd
                    ' Ticker and Amount are manual inputs.
                    With ws.Cells(r, scenCol)
                        .Locked = False
                        .Font.Color = RGB(0, 0, 255)
                        .Font.Bold = True
                    End With

                    With ws.Cells(r, scenCol + 1)
                        .Locked = False
                        .Font.Color = RGB(0, 0, 255)
                        .Font.Bold = True
                    End With

                    ' Asset Class and Yield are formula/helper cells.
                    ws.Cells(r, scenCol + 2).Locked = True
                    ws.Cells(r, scenCol + 3).Locked = True
                Next r
            End If
        End If
    Next sn

SafeExit:
End Sub

' ============================================================
' SELL WORKBENCH AMOUNT-SPEC UNLOCK
'
' BuildSellWorkbench unlocks its own input cells when it runs, but any
' later macro (Spawn Scenario, Remove Scenario, Add Buy Plans) reruns
' ApplyScenarioUXRulesToSheet, which locks the ENTIRE sheet and then only
' re-unlocks the ranges this module knows about. Without this unlock,
' Sell Mode / Amt Type / Amount go read-only after the next macro run.
' Manual Sell $ is a formula column now, so it stays locked.
' ============================================================

Private Sub UnlockSellWorkbenchInputs(ws As Worksheet, dataStart As Long, dataEnd As Long)
    On Error GoTo SafeExit

    Dim modeCol As Long
    modeCol = FindSellWorkbenchModeCol_UX(ws)

    If modeCol = 0 Then Exit Sub

    ' Column order from BuildSellWorkbench: Sell Mode | Amt Type | Amount | Manual Sell $
    Dim specTypeCol As Long
    Dim specAmtCol As Long
    Dim manualCol As Long

    specTypeCol = modeCol + 1
    specAmtCol = modeCol + 2
    manualCol = modeCol + 3

    ws.Range(ws.Cells(dataStart, modeCol), ws.Cells(dataEnd, modeCol)).Locked = False
    ws.Range(ws.Cells(dataStart, specTypeCol), ws.Cells(dataEnd, specTypeCol)).Locked = False
    ws.Range(ws.Cells(dataStart, specAmtCol), ws.Cells(dataEnd, specAmtCol)).Locked = False
    ws.Range(ws.Cells(dataStart, manualCol), ws.Cells(dataEnd, manualCol)).Locked = True

SafeExit:
End Sub

' ============================================================
' PROCEEDS ROUTING UNLOCK
'
' CDS_Routing.bas writes Destination/Detail/Spec/Amount (rows 12-16)
' unlocked when it builds the block, but ApplyScenarioUXRulesToSheet
' locks the ENTIRE sheet on every run before re-unlocking known ranges.
' Without this, routing inputs go read-only after the next macro run
' (Spawn Scenario, Remove Scenario, Add Buy Plans). Routed $ and the
' status row stay locked - they are formulas.
' ============================================================

Private Sub UnlockRoutingInputs(ws As Worksheet)
    On Error GoTo SafeExit

    Dim c As Range
    Set c = FindRoutingTitleCell_UX(ws)

    If c Is Nothing Then Exit Sub

    Dim destCol As Long
    destCol = c.Column

    Dim dataStart As Long
    Dim dataEnd As Long
    dataStart = c.Row + 2
    dataEnd = dataStart + 4

    ws.Range(ws.Cells(dataStart, destCol), ws.Cells(dataEnd, destCol + 3)).Locked = False
    ws.Range(ws.Cells(dataStart, destCol + 4), ws.Cells(dataEnd, destCol + 4)).Locked = True

SafeExit:
End Sub

Private Function FindRoutingTitleCell_UX(ws As Worksheet) As Range
    On Error Resume Next
    Set FindRoutingTitleCell_UX = ws.Cells.Find(What:="PROCEEDS ROUTING", _
                                                LookIn:=xlValues, _
                                                LookAt:=xlWhole, _
                                                SearchOrder:=xlByRows, _
                                                SearchDirection:=xlNext, _
                                                MatchCase:=False)
    On Error GoTo 0
End Function

Private Function FindSellWorkbenchModeCol_UX(ws As Worksheet) As Long
    Dim c As Range

    On Error Resume Next
    Set c = ws.Cells.Find(What:="Sell Mode", _
                          LookIn:=xlValues, _
                          LookAt:=xlWhole, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          MatchCase:=False)
    On Error GoTo 0

    If Not c Is Nothing Then FindSellWorkbenchModeCol_UX = c.Column
End Function

' ============================================================
' DETECTION HELPERS
' ============================================================

Private Function GetProcessedSheetState(ws As Worksheet, ByRef headerRow As Long, ByRef dataStart As Long, _
                                        ByRef dataEnd As Long, ByRef totRow As Long) As Boolean
    headerRow = FindHeaderRowUX(ws)

    If headerRow = 0 Then Exit Function

    dataStart = headerRow + 1
    totRow = FindTotalRowUX(ws, dataStart)

    If totRow = 0 Then Exit Function

    dataEnd = totRow - 1

    GetProcessedSheetState = True
End Function

Private Function IsProcessedCDSReport(ws As Worksheet) As Boolean
    IsProcessedCDSReport = (FindHeaderRowUX(ws) > 0)
End Function

Private Function FindHeaderRowUX(ws As Worksheet) As Long
    Dim r As Long

    For r = 1 To 10
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRowUX = r
            Exit Function
        End If
    Next r
End Function

Private Function FindTotalRowUX(ws As Worksheet, dataStart As Long) As Long
    Dim r As Long

    For r = dataStart To dataStart + 500
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" And Trim(CStr(ws.Cells(r, 5).Value)) <> "" Then
            FindTotalRowUX = r
            Exit Function
        End If
    Next r
End Function

Private Function CountScenariosUX(ws As Worksheet) As Long
    On Error GoTo Done

    Dim i As Long
    Dim col As Long

    For i = 0 To 30
        col = ScenStartCol() + i * ScenStride()

        If InStr(1, CStr(ws.Cells(2, col).Value), "Raise $", vbTextCompare) > 0 Then
            CountScenariosUX = i + 1
        Else
            Exit Function
        End If
    Next i

Done:
End Function

Private Function FindScenarioSummaryCol(ws As Worksheet) As Long
    Dim c As Range

    On Error Resume Next
    Set c = ws.Cells.Find(What:="SCENARIO SUMMARY", _
                          LookIn:=xlValues, _
                          LookAt:=xlWhole, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          MatchCase:=False)
    On Error GoTo 0

    If Not c Is Nothing Then
        FindScenarioSummaryCol = c.Column
    End If
End Function

Private Function FindLabelRowInColumn(ws As Worksheet, colNum As Long, labelText As String) As Long
    Dim r As Long

    For r = 1 To 500
        If UCase(Trim(CStr(ws.Cells(r, colNum).Value))) = UCase(Trim(labelText)) Then
            FindLabelRowInColumn = r
            Exit Function
        End If
    Next r
End Function

Private Function FindScenarioPivotGrandTotalRowUX(ws As Worksheet, scenCol As Long, totRow As Long) As Long
    Dim r As Long

    For r = totRow + 2 To totRow + 150
        If UCase(Trim(CStr(ws.Cells(r, scenCol).Value))) = "GRAND TOTAL" Then
            FindScenarioPivotGrandTotalRowUX = r
            Exit Function
        End If
    Next r
End Function
Private Function FindBuyPlanHeaderRow_UX(ws As Worksheet, startCol As Long, pivotEnd As Long) As Long
    Dim r As Long

    FindBuyPlanHeaderRow_UX = 0

    For r = pivotEnd + 1 To pivotEnd + 120
        If UCase(Trim(CStr(ws.Cells(r, startCol).Value))) = "BUY PLAN" Then
            FindBuyPlanHeaderRow_UX = r
            Exit Function
        End If
    Next r
End Function
Private Function ColLetterUX(colNum As Long) As String
    ColLetterUX = Split(Columns(colNum).Address(, False), ":")(0)
End Function

