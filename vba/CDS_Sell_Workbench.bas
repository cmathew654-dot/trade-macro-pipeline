Attribute VB_Name = "CDS_Sell_Workbench"
Option Explicit

' ============================================================
' CDS Sell Workbench
'
' Builds a simple proposed sell plan on top of the existing
' scenario pipeline:
'   S1 = pro-rata baseline
'   S2 = proposed sell plan driven by easy row-level controls
'
' The final proposed sell amount is still written to S2 Raise $,
' so AddBuyPlans and GenerateTradeEmail keep using the existing
' scenario contract.
' ============================================================

Private Const WORKBENCH_TITLE As String = "SELL WORKBENCH"
Private Const PLAN_TARGET_LABEL As String = "Plan Target Raise"
Private Const MODE_HEADER As String = "Sell Mode"
Private Const SPEC_TYPE_HEADER As String = "Amt Type"
Private Const SPEC_AMT_HEADER As String = "Amount"
Private Const MANUAL_HEADER As String = "Manual Sell $"
Private Const PROPOSED_HEADER As String = "Proposed Sell $"
Private Const USED_HEADER As String = "Manual Used $"
Private Const STATUS_HEADER As String = "Plan Status"
Private Const WORKBENCH_WIDTH As Long = 12
Private Const PLAN_SCENARIO_NUM As Long = 2

Public Sub BuildSellWorkbench()
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation

    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB. Open the client's workbook first.", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim headerRow As Long
    Dim dataStart As Long
    Dim dataEnd As Long
    Dim totalRow As Long

    If Not GetSellSheetState(ws, headerRow, dataStart, dataEnd, totalRow) Then
        MsgBox "Run ProcessCDSHoldings first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    PrepareCDSWorksheetForMacro ws
    EnsureSellPlanScenario ws

    If GetScenarioCount(ws) < PLAN_SCENARIO_NUM Then
        MsgBox "The proposed sell plan could not create S2.", vbExclamation
        GoTo Done
    End If

    BuildSellWorkbenchOnSheet ws, headerRow, dataStart, dataEnd, totalRow

    Application.Calculate

Done:
    On Error Resume Next
    If Not ws Is Nothing Then ProtectCDSWorksheet ws
    On Error GoTo 0

    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Exit Sub

ErrHandler:
    MsgBox "BuildSellWorkbench failed: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Done
End Sub

Private Sub EnsureSellPlanScenario(ws As Worksheet)
    Dim scenarioCount As Long

    scenarioCount = GetScenarioCount(ws)

    If scenarioCount = 0 Then
        AddRaiseCashScenarios
        PrepareCDSWorksheetForMacro ws
        scenarioCount = GetScenarioCount(ws)
    End If

    If scenarioCount = 1 Then
        SpawnScenario
        PrepareCDSWorksheetForMacro ws
    End If
End Sub

Private Sub BuildSellWorkbenchOnSheet(ws As Worksheet, headerRow As Long, dataStart As Long, _
                                      dataEnd As Long, totalRow As Long)
    Dim oldTarget As Variant
    Dim oldModeCol As Long
    Dim oldManualCol As Long
    Dim oldSpecTypeCol As Long
    Dim oldSpecAmtCol As Long
    Dim oldModes() As Variant
    Dim oldManuals() As Variant
    Dim oldSpecTypes() As Variant
    Dim oldSpecAmts() As Variant
    Dim r As Long
    Dim legacySheet As Boolean

    oldTarget = ExistingPlanTarget(ws)
    oldModeCol = ExistingHeaderCol(ws, MODE_HEADER)
    oldManualCol = ExistingHeaderCol(ws, MANUAL_HEADER)
    oldSpecTypeCol = ExistingHeaderCol(ws, SPEC_TYPE_HEADER)
    oldSpecAmtCol = ExistingHeaderCol(ws, SPEC_AMT_HEADER)

    ' A sheet built before the amount-spec engine existed has a Manual Sell $
    ' header but no Amt Type/Amount headers. Seed specs from the old typed
    ' manual $ value so the advisor's plan survives the rebuild unchanged.
    legacySheet = (oldSpecTypeCol = 0 And oldSpecAmtCol = 0 And oldManualCol > 0)

    ReDim oldModes(dataStart To dataEnd)
    ReDim oldManuals(dataStart To dataEnd)
    ReDim oldSpecTypes(dataStart To dataEnd)
    ReDim oldSpecAmts(dataStart To dataEnd)

    For r = dataStart To dataEnd
        If oldModeCol > 0 Then oldModes(r) = ws.Cells(r, oldModeCol).Value
        If oldManualCol > 0 Then oldManuals(r) = ws.Cells(r, oldManualCol).Value
        If oldSpecTypeCol > 0 Then oldSpecTypes(r) = ws.Cells(r, oldSpecTypeCol).Value
        If oldSpecAmtCol > 0 Then oldSpecAmts(r) = ws.Cells(r, oldSpecAmtCol).Value

        If legacySheet And IsNumeric(oldManuals(r)) Then
            oldSpecTypes(r) = "$"
            oldSpecAmts(r) = oldManuals(r)
        End If
    Next r

    ClearExistingWorkbench ws, totalRow

    Dim summaryCol As Long
    summaryCol = FindScenarioSummaryColSell(ws)
    If summaryCol = 0 Then Err.Raise vbObjectError + 520, , "Scenario Summary not found."

    Dim scenarioCount As Long
    scenarioCount = GetScenarioCount(ws)

    Dim workCol As Long
    workCol = summaryCol + scenarioCount + 2

    Dim targetCol As Long
    Dim targetInputCol As Long
    Dim modeCol As Long
    Dim specTypeCol As Long
    Dim specAmtCol As Long
    Dim manualCol As Long
    Dim proposedCol As Long
    Dim usedCol As Long
    Dim statusCol As Long

    targetCol = workCol
    targetInputCol = workCol + 1
    modeCol = workCol + 3
    specTypeCol = workCol + 4
    specAmtCol = workCol + 5
    manualCol = workCol + 6
    proposedCol = workCol + 7
    usedCol = workCol + 8
    statusCol = workCol + 10

    ws.Range(ws.Cells(1, workCol), ws.Cells(totalRow + 90, workCol + WORKBENCH_WIDTH)).UnMerge
    ws.Range(ws.Cells(1, workCol), ws.Cells(totalRow + 90, workCol + WORKBENCH_WIDTH)).Clear
    ws.Range(ws.Columns(workCol), ws.Columns(workCol + WORKBENCH_WIDTH)).Hidden = False

    BuildWorkbenchHeader ws, headerRow, workCol, targetCol, targetInputCol, modeCol, _
                         specTypeCol, specAmtCol, manualCol, proposedCol, usedCol, statusCol, oldTarget

    BuildWorkbenchRows ws, dataStart, dataEnd, totalRow, targetInputCol, modeCol, _
                       specTypeCol, specAmtCol, manualCol, proposedCol, usedCol, _
                       oldModes, oldSpecTypes, oldSpecAmts

    ApplyProposedScenarioTitle ws, headerRow
    ApplyWorkbenchStatus ws, dataStart, dataEnd, targetInputCol, modeCol, manualCol, _
                         proposedCol, usedCol, statusCol
    ApplyWorkbenchFormatting ws, dataStart, dataEnd, workCol, targetInputCol, modeCol, _
                             specTypeCol, specAmtCol, manualCol, proposedCol, usedCol, statusCol
End Sub

Private Sub BuildWorkbenchHeader(ws As Worksheet, headerRow As Long, workCol As Long, _
                                 targetCol As Long, targetInputCol As Long, modeCol As Long, _
                                 specTypeCol As Long, specAmtCol As Long, manualCol As Long, _
                                 proposedCol As Long, usedCol As Long, _
                                 statusCol As Long, oldTarget As Variant)
    With ws.Range(ws.Cells(1, workCol), ws.Cells(1, workCol + WORKBENCH_WIDTH))
        .Merge
        .Value = WORKBENCH_TITLE
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(31, 78, 121)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    With ws.Cells(headerRow, targetCol)
        .Value = PLAN_TARGET_LABEL
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Cells(headerRow, targetInputCol)
        If IsNumeric(oldTarget) Then
            .Value = CDbl(oldTarget)
        Else
            .Value = 0
        End If
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 255)
        .Interior.Color = RGB(197, 217, 241)
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    SetRangeLockedSafe ws.Cells(headerRow, targetInputCol), False

    AddHeaderCell ws, headerRow, modeCol, MODE_HEADER
    AddHeaderCell ws, headerRow, specTypeCol, SPEC_TYPE_HEADER
    AddHeaderCell ws, headerRow, specAmtCol, SPEC_AMT_HEADER
    AddHeaderCell ws, headerRow, manualCol, MANUAL_HEADER
    AddHeaderCell ws, headerRow, proposedCol, PROPOSED_HEADER
    AddHeaderCell ws, headerRow, usedCol, USED_HEADER
    AddHeaderCell ws, headerRow, statusCol, STATUS_HEADER
End Sub

Private Sub AddHeaderCell(ws As Worksheet, rowNum As Long, colNum As Long, textValue As String)
    With ws.Cells(rowNum, colNum)
        .Value = textValue
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .Interior.Color = RGB(189, 215, 238)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub BuildWorkbenchRows(ws As Worksheet, dataStart As Long, dataEnd As Long, totalRow As Long, _
                               targetInputCol As Long, modeCol As Long, specTypeCol As Long, _
                               specAmtCol As Long, manualCol As Long, proposedCol As Long, _
                               usedCol As Long, oldModes() As Variant, oldSpecTypes() As Variant, _
                               oldSpecAmts() As Variant)
    Dim scenCol As Long
    scenCol = ScenStartCol() + (PLAN_SCENARIO_NUM - 1) * ScenStride()

    Dim targetCell As String
    Dim modeRange As String
    Dim usedRange As String
    Dim fmvRange As String
    Dim targetL As String
    Dim modeL As String
    Dim usedL As String
    Dim typeL As String
    Dim amtL As String

    targetL = ColLetterSell(targetInputCol)
    modeL = ColLetterSell(modeCol)
    usedL = ColLetterSell(usedCol)
    typeL = ColLetterSell(specTypeCol)
    amtL = ColLetterSell(specAmtCol)

    targetCell = "$" & targetL & "$2"
    modeRange = "$" & modeL & "$" & dataStart & ":$" & modeL & "$" & dataEnd
    usedRange = "$" & usedL & "$" & dataStart & ":$" & usedL & "$" & dataEnd
    fmvRange = "$E$" & dataStart & ":$E$" & dataEnd

    Dim poolFmvFormula As String
    Dim residualFormula As String
    poolFmvFormula = "SUMIFS(" & fmvRange & "," & modeRange & ",""Pool"")"
    residualFormula = "MAX(0," & targetCell & "-SUM(" & usedRange & "))"

    Dim r As Long
    Dim modeValue As String
    Dim specTypeValue As String

    For r = dataStart To dataEnd
        modeValue = NormalizeSellMode(oldModes(r), ws.Cells(r, 1).Value, ws.Cells(r, 3).Value)

        With ws.Cells(r, modeCol)
            .Value = modeValue
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        SetRangeLockedSafe ws.Cells(r, modeCol), False

        specTypeValue = Trim(CStr(oldSpecTypes(r)))
        If specTypeValue = "" Then specTypeValue = "$"

        With ws.Cells(r, specTypeCol)
            .Value = specTypeValue
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        SetRangeLockedSafe ws.Cells(r, specTypeCol), False

        With ws.Cells(r, specAmtCol)
            If IsNumeric(oldSpecAmts(r)) Then .Value = CDbl(oldSpecAmts(r)) Else .Value = 0
            .NumberFormat = "#,##0.00;(#,##0.00);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        SetRangeLockedSafe ws.Cells(r, specAmtCol), False

        With ws.Cells(r, manualCol)
            .Formula = "=IF(" & modeL & r & "<>""Manual"",0,IF(" & typeL & r & _
                       "=""ALL"",E" & r & ",IF(" & typeL & r & "=""Shares""," & _
                       amtL & r & "*IFERROR(E" & r & "/K" & r & ",0),IF(" & typeL & r & _
                       "=""% Pos""," & amtL & r & "/100*E" & r & ",IF(" & typeL & r & _
                       "=""% Acct""," & amtL & r & "/100*$E$" & totalRow & "," & amtL & r & ")))))"
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        SetRangeLockedSafe ws.Cells(r, manualCol), True

        With ws.Cells(r, usedCol)
            .Formula = "=IF(" & ColLetterSell(modeCol) & r & "=""Manual"",MIN(MAX(0," & _
                       ColLetterSell(manualCol) & r & "),E" & r & "),0)"
            .NumberFormat = "$#,##0;($#,##0);""-"""
        End With

        With ws.Cells(r, scenCol)
            .Formula = "=IF(" & ColLetterSell(modeCol) & r & "=""Manual""," & _
                       ColLetterSell(usedCol) & r & ",IF(" & ColLetterSell(modeCol) & r & _
                       "=""Pool"",IF(" & poolFmvFormula & "=0,0,MIN(E" & r & "," & _
                       residualFormula & "*E" & r & "/" & poolFmvFormula & ")),0))"
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .Font.Italic = False
            .Interior.Color = RGB(226, 239, 218)
        End With
        SetRangeLockedSafe ws.Cells(r, scenCol), True

        With ws.Cells(r, proposedCol)
            .Formula = "=" & ColLetterSell(scenCol) & r
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next r

    With ws.Cells(totalRow, modeCol)
        .Value = "Total"
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Cells(totalRow, manualCol)
        .Formula = "=SUM(" & ColLetterSell(manualCol) & dataStart & ":" & ColLetterSell(manualCol) & dataEnd & ")"
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Cells(totalRow, proposedCol)
        .Formula = "=SUM(" & ColLetterSell(proposedCol) & dataStart & ":" & ColLetterSell(proposedCol) & dataEnd & ")"
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub ApplyWorkbenchStatus(ws As Worksheet, dataStart As Long, dataEnd As Long, _
                                 targetInputCol As Long, modeCol As Long, manualCol As Long, _
                                 proposedCol As Long, usedCol As Long, statusCol As Long)
    Dim targetCell As String
    Dim proposedRange As String
    Dim usedRange As String
    Dim totalProposed As String
    Dim totalUsed As String
    Dim shortfallExpr As String
    Dim overageExpr As String
    Dim cashAvailCell As String

    targetCell = "$" & ColLetterSell(targetInputCol) & "$2"
    proposedRange = "$" & ColLetterSell(proposedCol) & "$" & dataStart & ":$" & ColLetterSell(proposedCol) & "$" & dataEnd
    usedRange = "$" & ColLetterSell(usedCol) & "$" & dataStart & ":$" & ColLetterSell(usedCol) & "$" & dataEnd
    totalProposed = "SUM(" & proposedRange & ")"
    totalUsed = "SUM(" & usedRange & ")"
    shortfallExpr = targetCell & "-" & totalProposed
    overageExpr = totalProposed & "-" & targetCell

    With ws.Cells(3, statusCol)
        .Formula = "=IF(" & targetCell & "<=0,""Enter one target raise amount.""," & _
                   "IF(" & totalUsed & ">" & targetCell & ",""Manual sells exceed target; pool is zero.""," & _
                   "IF(" & shortfallExpr & ">0.5,""SHORTFALL: raise is short $""&TEXT(" & shortfallExpr & ",""#,##0"")&"" vs target""," & _
                   "IF(" & overageExpr & ">0.5,""OVERAGE: proposed sells exceed target by $""&TEXT(" & overageExpr & ",""#,##0"")," & _
                   """OK: proposed sells match target.""))))"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Shortfall state overrides the default amber status styling with red/bold.
    On Error Resume Next
    ws.Cells(3, statusCol).FormatConditions.Delete
    On Error GoTo 0

    Dim shortfallFC As FormatCondition
    Set shortfallFC = ws.Cells(3, statusCol).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=LEFT($" & ColLetterSell(statusCol) & "$3,9)=""SHORTFALL""")
    With shortfallFC
        .Font.Bold = True
        .Font.Color = RGB(156, 0, 6)
        .Interior.Color = RGB(255, 199, 206)
    End With

    ws.Cells(4, statusCol).Value = "Manual Used"
    ws.Cells(4, statusCol + 1).Formula = "=" & totalUsed
    ws.Cells(5, statusCol).Value = "Pool Residual"
    ws.Cells(5, statusCol + 1).Formula = "=MAX(0," & targetCell & "-" & totalUsed & ")"
    ws.Cells(6, statusCol).Value = "Total Proposed"
    ws.Cells(6, statusCol + 1).Formula = "=" & totalProposed
    ws.Cells(7, statusCol).Value = "Cash/MM Avail"
    ws.Cells(7, statusCol + 1).Formula = "=SUMIF($A:$A,""CASH"",$E:$E)+SUMIF($A:$A,""SHORT"",$E:$E)"

    ws.Range(ws.Cells(4, statusCol), ws.Cells(7, statusCol + 1)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(4, statusCol), ws.Cells(7, statusCol)).Font.Bold = True
    ws.Range(ws.Cells(4, statusCol + 1), ws.Cells(7, statusCol + 1)).NumberFormat = "$#,##0;($#,##0);""-"""

    ' Row 8: non-blocking advisory when the target could be covered by cash/MM alone.
    cashAvailCell = "$" & ColLetterSell(statusCol + 1) & "$7"

    With ws.Cells(8, statusCol)
        .Formula = "=IF(AND(" & targetCell & "<=" & cashAvailCell & "," & targetCell & ">0),""Target <= available cash/MM - redemption may cover this without selling."","""")"
        .Font.Italic = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub ApplyWorkbenchFormatting(ws As Worksheet, dataStart As Long, dataEnd As Long, _
                                     workCol As Long, targetInputCol As Long, modeCol As Long, _
                                     specTypeCol As Long, specAmtCol As Long, manualCol As Long, _
                                     proposedCol As Long, usedCol As Long, statusCol As Long)
    Dim rngModes As Range
    Set rngModes = ws.Range(ws.Cells(dataStart, modeCol), ws.Cells(dataEnd, modeCol))

    On Error Resume Next
    rngModes.Validation.Delete
    rngModes.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                            Operator:=xlBetween, Formula1:="Pool,Manual,Exclude"
    ws.Range(ws.Cells(dataStart, specTypeCol), ws.Cells(dataEnd, specTypeCol)).Validation.Delete
    ws.Range(ws.Cells(dataStart, specTypeCol), ws.Cells(dataEnd, specTypeCol)).Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="$,Shares,% Pos,% Acct,ALL"
    ws.Range(ws.Cells(dataStart, specAmtCol), ws.Cells(dataEnd, specAmtCol)).Validation.Delete
    ws.Range(ws.Cells(dataStart, specAmtCol), ws.Cells(dataEnd, specAmtCol)).Validation.Add _
        Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
    On Error GoTo 0

    ws.Range(ws.Cells(dataStart, modeCol), ws.Cells(dataEnd, modeCol)).Interior.Color = RGB(226, 239, 218)
    ws.Range(ws.Cells(dataStart, specTypeCol), ws.Cells(dataEnd, specTypeCol)).Interior.Color = RGB(197, 217, 241)
    ws.Range(ws.Cells(dataStart, specAmtCol), ws.Cells(dataEnd, specAmtCol)).Interior.Color = RGB(197, 217, 241)
    ws.Range(ws.Cells(dataStart, manualCol), ws.Cells(dataEnd, manualCol)).Interior.Color = RGB(226, 239, 218)
    ws.Range(ws.Cells(dataStart, proposedCol), ws.Cells(dataEnd, proposedCol)).Interior.Color = RGB(226, 239, 218)

    ws.Columns(ColLetterSell(workCol)).ColumnWidth = 18
    ws.Columns(ColLetterSell(targetInputCol)).ColumnWidth = 14
    ws.Columns(ColLetterSell(workCol + 2)).ColumnWidth = 2
    ws.Columns(ColLetterSell(modeCol)).ColumnWidth = 12
    ws.Columns(ColLetterSell(specTypeCol)).ColumnWidth = 10
    ws.Columns(ColLetterSell(specAmtCol)).ColumnWidth = 12
    ws.Columns(ColLetterSell(manualCol)).ColumnWidth = 14
    ws.Columns(ColLetterSell(proposedCol)).ColumnWidth = 14
    ws.Columns(ColLetterSell(usedCol)).Hidden = True
    ws.Columns(ColLetterSell(statusCol)).ColumnWidth = 32
    ws.Columns(ColLetterSell(statusCol + 1)).ColumnWidth = 14

    SetRangeLockedSafe ws.Cells(2, targetInputCol), False
    SetRangeLockedSafe ws.Range(ws.Cells(dataStart, modeCol), ws.Cells(dataEnd, modeCol)), False
    SetRangeLockedSafe ws.Range(ws.Cells(dataStart, specTypeCol), ws.Cells(dataEnd, specTypeCol)), False
    SetRangeLockedSafe ws.Range(ws.Cells(dataStart, specAmtCol), ws.Cells(dataEnd, specAmtCol)), False
    SetRangeLockedSafe ws.Range(ws.Cells(dataStart, manualCol), ws.Cells(dataEnd, manualCol)), True
    SetRangeLockedSafe ws.Range(ws.Cells(dataStart, proposedCol), ws.Cells(dataEnd, proposedCol)), True

    On Error Resume Next
    ws.Cells(2, targetInputCol).Comment.Delete
    ws.Cells(2, targetInputCol).AddComment "Type the total cash amount to raise. Pool rows share the residual after Manual rows. Exclude rows sell zero."
    ws.Range(ws.Cells(dataStart, modeCol), ws.Cells(dataEnd, modeCol)).Comment.Delete
    ws.Cells(dataStart, modeCol).AddComment "Choose Pool, Manual, or Exclude."
    ws.Range(ws.Cells(dataStart, specTypeCol), ws.Cells(dataEnd, specTypeCol)).Comment.Delete
    ws.Cells(dataStart, specTypeCol).AddComment "Only used when Sell Mode is Manual. $ = dollar amount, Shares = share count, % Pos = % of this holding, % Acct = % of account, ALL = full position."
    ws.Range(ws.Cells(dataStart, specAmtCol), ws.Cells(dataEnd, specAmtCol)).Comment.Delete
    ws.Cells(dataStart, specAmtCol).AddComment "Amount to apply using Amt Type. Ignored when Amt Type is ALL."
    ws.Range(ws.Cells(dataStart, manualCol), ws.Cells(dataEnd, manualCol)).Comment.Delete
    ws.Cells(dataStart, manualCol).AddComment "Calculated from Amt Type/Amount when Sell Mode is Manual."
    On Error GoTo 0
End Sub

Private Sub ApplyProposedScenarioTitle(ws As Worksheet, headerRow As Long)
    Dim scenCol As Long
    Dim titleRow As Long

    scenCol = ScenStartCol() + (PLAN_SCENARIO_NUM - 1) * ScenStride()
    titleRow = headerRow - 1

    On Error Resume Next
    ws.Range(ws.Cells(titleRow, scenCol), ws.Cells(titleRow, scenCol + 4)).UnMerge
    On Error GoTo 0

    ws.Range(ws.Cells(titleRow, scenCol), ws.Cells(titleRow, scenCol + 4)).Merge

    With ws.Cells(titleRow, scenCol)
        .Value = "S2 Proposed Sell Plan"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(84, 130, 53)
    End With
    SetRangeLockedSafe ws.Cells(titleRow, scenCol), True
End Sub

Private Sub SetRangeLockedSafe(targetRange As Range, isLocked As Boolean)
    On Error Resume Next
    targetRange.Locked = isLocked
    On Error GoTo 0
End Sub

Private Function NormalizeSellMode(valueIn As Variant, assetClassValue As Variant, _
                                   Optional tickerValue As Variant = "") As String
    Dim s As String
    s = UCase(Trim(CStr(valueIn)))

    Select Case s
        Case "POOL"
            NormalizeSellMode = "Pool"
        Case "MANUAL"
            NormalizeSellMode = "Manual"
        Case "EXCLUDE"
            NormalizeSellMode = "Exclude"
        Case Else
            If UCase(Trim(CStr(assetClassValue))) = "CASH" Then
                NormalizeSellMode = "Exclude"
            ElseIf IsCUSIPSettings(CStr(tickerValue)) Then
                NormalizeSellMode = "Exclude"
            Else
                NormalizeSellMode = "Pool"
            End If
    End Select
End Function

Private Function ExistingPlanTarget(ws As Worksheet) As Variant
    Dim c As Range
    Set c = FindCellExactSell(ws, PLAN_TARGET_LABEL)

    If c Is Nothing Then
        ExistingPlanTarget = Empty
    Else
        ExistingPlanTarget = ws.Cells(c.Row, c.Column + 1).Value
    End If
End Function

Private Function ExistingHeaderCol(ws As Worksheet, headerText As String) As Long
    Dim c As Range
    Set c = FindCellExactSell(ws, headerText)

    If Not c Is Nothing Then ExistingHeaderCol = c.Column
End Function

Private Sub ClearExistingWorkbench(ws As Worksheet, totalRow As Long)
    Dim c As Range
    Set c = FindCellExactSell(ws, WORKBENCH_TITLE)

    If c Is Nothing Then Exit Sub

    Dim clearEndRow As Long
    clearEndRow = totalRow + 90
    If clearEndRow < 120 Then clearEndRow = 120

    On Error Resume Next
    ws.Range(ws.Cells(1, c.Column), ws.Cells(clearEndRow, c.Column + WORKBENCH_WIDTH)).UnMerge
    ws.Range(ws.Columns(c.Column), ws.Columns(c.Column + WORKBENCH_WIDTH)).Hidden = False
    On Error GoTo 0

    ws.Range(ws.Cells(1, c.Column), ws.Cells(clearEndRow, c.Column + WORKBENCH_WIDTH)).Clear
End Sub

Private Function GetSellSheetState(ws As Worksheet, ByRef headerRow As Long, ByRef dataStart As Long, _
                                   ByRef dataEnd As Long, ByRef totalRow As Long) As Boolean
    headerRow = FindHeaderRowSell(ws)
    If headerRow = 0 Then Exit Function

    dataStart = headerRow + 1
    totalRow = FindTotalRowSell(ws, dataStart)
    If totalRow = 0 Then Exit Function

    dataEnd = totalRow - 1
    GetSellSheetState = True
End Function

Private Function FindHeaderRowSell(ws As Worksheet) As Long
    Dim r As Long

    For r = 1 To 10
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRowSell = r
            Exit Function
        End If
    Next r
End Function

Private Function FindTotalRowSell(ws As Worksheet, dataStart As Long) As Long
    Dim r As Long

    For r = dataStart To dataStart + 500
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" And Trim(CStr(ws.Cells(r, 5).Value)) <> "" Then
            FindTotalRowSell = r
            Exit Function
        End If
    Next r
End Function

Private Function FindScenarioSummaryColSell(ws As Worksheet) As Long
    Dim c As Range

    On Error Resume Next
    Set c = ws.Cells.Find(What:="SCENARIO SUMMARY", _
                          LookIn:=xlValues, _
                          LookAt:=xlWhole, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          MatchCase:=False)
    On Error GoTo 0

    If Not c Is Nothing Then FindScenarioSummaryColSell = c.Column
End Function

Private Function FindCellExactSell(ws As Worksheet, textValue As String) As Range
    On Error Resume Next
    Set FindCellExactSell = ws.Cells.Find(What:=textValue, _
                                          LookIn:=xlValues, _
                                          LookAt:=xlWhole, _
                                          SearchOrder:=xlByRows, _
                                          SearchDirection:=xlNext, _
                                          MatchCase:=False)
    On Error GoTo 0
End Function

Private Function ColLetterSell(colNum As Long) As String
    ColLetterSell = Split(Columns(colNum).Address(, False), ":")(0)
End Function
