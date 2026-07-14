Attribute VB_Name = "CDS_Routing"
Option Explicit

' ============================================================
' CDS Proceeds Routing
'
' Lets the advisor declare up front where the plan scenario's
' sell proceeds go, instead of implicitly funding the whole buy
' plan and prompting for the residual destination at email time:
'   Destination: Buy Plan | Money Market | Transfer Out | Hold in Cash
'   Spec:        $ | % of Proceeds | Residual
'   Routed $:    $              = Amount
'                % of Proceeds  = Amount/100 * ProceedsTotal
'                Residual       = MAX(0, ProceedsTotal - SUM(other
'                                 rows where Spec <> "Residual"))
'
' Lives inside the Sell Workbench block, directly below its
' status rows (row 10 down), in the same statusCol as the
' workbench status block. ProceedsTotal is the workbench's own
' "Total Proposed" value cell (row 6, statusCol + 1).
'
' BuildSellWorkbench (CDS_Sell_Workbench.bas) captures any
' existing routing rows before it clears the workbench area, then
' calls WriteRoutingBlock so the block always exists after a
' rebuild - preserved values when present, defaults when not.
' BuildRoutingBlock is the standalone entry point (Alt+F8) for
' adding the block to a workbench built before this feature
' existed.
'
' AddBuyPlans (CDS_Buy_Plans.bas) reads the "Buy Plan" row's
' Routed $ as the S2 (plan scenario) funding basis whenever this
' block exists on the sheet.
' ============================================================

Private Const ROUTING_TITLE As String = "PROCEEDS ROUTING"
Private Const ROUTING_TITLE_ROW As Long = 10
Private Const ROUTING_HEADER_ROW As Long = 11
Private Const ROUTING_DATA_START_ROW As Long = 12
Private Const ROUTING_DATA_END_ROW As Long = 16
Private Const ROUTING_DATA_ROWS As Long = 5
Private Const ROUTING_STATUS_ROW As Long = 17
Private Const ROUTING_DEST_LIST As String = "Buy Plan,Money Market,Transfer Out,Hold in Cash"
Private Const ROUTING_SPEC_LIST As String = "$,% of Proceeds,Residual"

Public Sub BuildRoutingBlock()
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

    Dim modeCheckCell As Range
    Set modeCheckCell = FindCellExactRouting(ws, "Sell Mode")

    Dim statusCol As Long
    If Not modeCheckCell Is Nothing Then statusCol = FindPlanStatusColRouting(ws)

    If modeCheckCell Is Nothing Or statusCol = 0 Then
        MsgBox "Run Build Sell Workbench first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    PrepareCDSWorksheetForMacro ws

    Dim routingFound As Boolean
    Dim routDest() As Variant
    Dim routDetail() As Variant
    Dim routSpec() As Variant
    Dim routAmt() As Variant
    routingFound = CaptureRoutingRows(ws, routDest, routDetail, routSpec, routAmt)

    WriteRoutingBlock ws, statusCol, routingFound, routDest, routDetail, routSpec, routAmt

    Application.Calculate

Done:
    On Error Resume Next
    If Not ws Is Nothing Then ProtectCDSWorksheet ws
    On Error GoTo 0

    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Exit Sub

ErrHandler:
    MsgBox "BuildRoutingBlock failed: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Done
End Sub

' ============================================================
' PRESERVE ACROSS REBUILDS
'
' Call BEFORE the workbench area is cleared. Returns False (all
' output arrays dimensioned but unused) when no prior routing
' block exists on the sheet.
' ============================================================
Public Function CaptureRoutingRows(ws As Worksheet, ByRef destOut() As Variant, _
                                   ByRef detailOut() As Variant, ByRef specOut() As Variant, _
                                   ByRef amtOut() As Variant) As Boolean
    ReDim destOut(1 To ROUTING_DATA_ROWS)
    ReDim detailOut(1 To ROUTING_DATA_ROWS)
    ReDim specOut(1 To ROUTING_DATA_ROWS)
    ReDim amtOut(1 To ROUTING_DATA_ROWS)

    CaptureRoutingRows = False

    Dim c As Range
    Set c = FindCellExactRouting(ws, ROUTING_TITLE)
    If c Is Nothing Then Exit Function

    Dim destCol As Long
    destCol = c.Column

    Dim dataStartRow As Long
    dataStartRow = c.Row + 2

    Dim i As Long
    For i = 1 To ROUTING_DATA_ROWS
        destOut(i) = ws.Cells(dataStartRow + i - 1, destCol).Value
        detailOut(i) = ws.Cells(dataStartRow + i - 1, destCol + 1).Value
        specOut(i) = ws.Cells(dataStartRow + i - 1, destCol + 2).Value
        amtOut(i) = ws.Cells(dataStartRow + i - 1, destCol + 3).Value
    Next i

    CaptureRoutingRows = True
End Function

' ============================================================
' BUILD / REBUILD THE BLOCK AT A GIVEN statusCol
' ============================================================
Public Sub WriteRoutingBlock(ws As Worksheet, statusCol As Long, hadPrior As Boolean, _
                             destArr() As Variant, detailArr() As Variant, _
                             specArr() As Variant, amtArr() As Variant)
    Dim destCol As Long, detailCol As Long, specCol As Long, amtCol As Long, routedCol As Long
    destCol = statusCol
    detailCol = statusCol + 1
    specCol = statusCol + 2
    amtCol = statusCol + 3
    routedCol = statusCol + 4

    On Error Resume Next
    ws.Cells(ROUTING_STATUS_ROW, destCol).FormatConditions.Delete
    ws.Range(ws.Cells(ROUTING_TITLE_ROW, destCol), ws.Cells(ROUTING_STATUS_ROW, routedCol)).UnMerge
    On Error GoTo 0
    ws.Range(ws.Cells(ROUTING_TITLE_ROW, destCol), ws.Cells(ROUTING_STATUS_ROW, routedCol)).Clear

    With ws.Range(ws.Cells(ROUTING_TITLE_ROW, destCol), ws.Cells(ROUTING_TITLE_ROW, routedCol))
        .Merge
        .Value = ROUTING_TITLE
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(89, 89, 89)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    AddRoutingHeaderCell ws, ROUTING_HEADER_ROW, destCol, "Destination"
    AddRoutingHeaderCell ws, ROUTING_HEADER_ROW, detailCol, "Detail"
    AddRoutingHeaderCell ws, ROUTING_HEADER_ROW, specCol, "Spec"
    AddRoutingHeaderCell ws, ROUTING_HEADER_ROW, amtCol, "Amount"
    AddRoutingHeaderCell ws, ROUTING_HEADER_ROW, routedCol, "Routed $"

    ' ProceedsTotal = the workbench's own "Total Proposed" value cell.
    Dim proceedsCell As String
    proceedsCell = "$" & ColLetterRouting(statusCol + 1) & "$6"

    Dim specL As String, amtL As String, routedL As String
    specL = ColLetterRouting(specCol)
    amtL = ColLetterRouting(amtCol)
    routedL = ColLetterRouting(routedCol)

    Dim i As Long, r As Long
    Dim vDest As Variant, vDetail As Variant, vSpec As Variant, vAmt As Variant

    For i = 1 To ROUTING_DATA_ROWS
        r = ROUTING_DATA_START_ROW + i - 1

        If hadPrior Then
            vDest = destArr(i)
            vDetail = detailArr(i)
            vSpec = specArr(i)
            vAmt = amtArr(i)
        Else
            GetRoutingDefaultRow i, vDest, vDetail, vSpec, vAmt
        End If

        With ws.Cells(r, destCol)
            .Value = vDest
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Interior.Color = RGB(197, 217, 241)
        End With
        SetRangeLockedSafeRouting ws.Cells(r, destCol), False

        With ws.Cells(r, detailCol)
            .Value = vDetail
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Interior.Color = RGB(197, 217, 241)
        End With
        SetRangeLockedSafeRouting ws.Cells(r, detailCol), False

        With ws.Cells(r, specCol)
            .Value = vSpec
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Interior.Color = RGB(197, 217, 241)
        End With
        SetRangeLockedSafeRouting ws.Cells(r, specCol), False

        With ws.Cells(r, amtCol)
            If IsNumeric(vAmt) Then .Value = CDbl(vAmt) Else .Value = 0
            .NumberFormat = "#,##0.00;(#,##0.00);""-"""
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Interior.Color = RGB(197, 217, 241)
        End With
        SetRangeLockedSafeRouting ws.Cells(r, amtCol), False

        With ws.Cells(r, routedCol)
            .Formula = "=IF(" & specL & r & "=""$""," & amtL & r & _
                       ",IF(" & specL & r & "=""% of Proceeds""," & amtL & r & "/100*" & proceedsCell & _
                       ",IF(" & specL & r & "=""Residual"",MAX(0," & proceedsCell & "-(" & _
                       BuildOtherRowsSum(specL, routedL, ROUTING_DATA_START_ROW, ROUTING_DATA_END_ROW, r) & _
                       ")),0)))"
            .NumberFormat = "$#,##0;($#,##0);""-"""
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Interior.Color = RGB(226, 239, 218)
        End With
        SetRangeLockedSafeRouting ws.Cells(r, routedCol), True
    Next i

    On Error Resume Next
    ws.Range(ws.Cells(ROUTING_DATA_START_ROW, destCol), ws.Cells(ROUTING_DATA_END_ROW, destCol)).Validation.Delete
    ws.Range(ws.Cells(ROUTING_DATA_START_ROW, destCol), ws.Cells(ROUTING_DATA_END_ROW, destCol)).Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ROUTING_DEST_LIST
    ws.Range(ws.Cells(ROUTING_DATA_START_ROW, specCol), ws.Cells(ROUTING_DATA_END_ROW, specCol)).Validation.Delete
    ws.Range(ws.Cells(ROUTING_DATA_START_ROW, specCol), ws.Cells(ROUTING_DATA_END_ROW, specCol)).Validation.Add _
        Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ROUTING_SPEC_LIST
    ws.Range(ws.Cells(ROUTING_DATA_START_ROW, amtCol), ws.Cells(ROUTING_DATA_END_ROW, amtCol)).Validation.Delete
    ws.Range(ws.Cells(ROUTING_DATA_START_ROW, amtCol), ws.Cells(ROUTING_DATA_END_ROW, amtCol)).Validation.Add _
        Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
    On Error GoTo 0

    Dim destRange As String, specRange As String, routedRange As String
    destRange = "$" & ColLetterRouting(destCol) & "$" & ROUTING_DATA_START_ROW & ":$" & ColLetterRouting(destCol) & "$" & ROUTING_DATA_END_ROW
    specRange = "$" & specL & "$" & ROUTING_DATA_START_ROW & ":$" & specL & "$" & ROUTING_DATA_END_ROW
    routedRange = "$" & routedL & "$" & ROUTING_DATA_START_ROW & ":$" & routedL & "$" & ROUTING_DATA_END_ROW

    Dim residualCountExpr As String
    residualCountExpr = "COUNTIFS(" & destRange & ",""<>""," & specRange & ",""Residual"")"

    Dim diffExpr As String
    diffExpr = "SUM(" & routedRange & ")-" & proceedsCell

    ws.Range(ws.Cells(ROUTING_STATUS_ROW, destCol), ws.Cells(ROUTING_STATUS_ROW, routedCol)).Merge

    With ws.Cells(ROUTING_STATUS_ROW, destCol)
        .Formula = "=IF(" & residualCountExpr & "<>1,""Need exactly one Residual row""," & _
                   "IF(ABS(" & diffExpr & ")>0.5,IF(" & diffExpr & ">0,""Routing over-allocates by $""&TEXT(" & diffExpr & ",""#,##0""),""Routing under-allocates by $""&TEXT(-(" & diffExpr & "),""#,##0"")),""Routing OK""))"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    SetRangeLockedSafeRouting ws.Cells(ROUTING_STATUS_ROW, destCol), True

    Dim notOkFC As FormatCondition
    Set notOkFC = ws.Cells(ROUTING_STATUS_ROW, destCol).FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$" & ColLetterRouting(destCol) & "$" & ROUTING_STATUS_ROW & "<>""Routing OK""")
    With notOkFC
        .Font.Bold = True
        .Font.Color = RGB(156, 0, 6)
        .Interior.Color = RGB(255, 199, 206)
    End With

    ws.Columns(ColLetterRouting(detailCol)).ColumnWidth = 22
    ws.Columns(ColLetterRouting(specCol)).ColumnWidth = 14
    ws.Columns(ColLetterRouting(amtCol)).ColumnWidth = 12
    ws.Columns(ColLetterRouting(routedCol)).ColumnWidth = 14

    On Error Resume Next
    ws.Cells(ROUTING_DATA_START_ROW, destCol).Comment.Delete
    ws.Cells(ROUTING_DATA_START_ROW, destCol).AddComment "Choose where these proceeds go. Exactly one row should be Residual."
    ws.Cells(ROUTING_DATA_START_ROW, specCol).Comment.Delete
    ws.Cells(ROUTING_DATA_START_ROW, specCol).AddComment "$ = fixed dollar amount, % of Proceeds = percent of the total raised, Residual = whatever is left after the other rows."
    On Error GoTo 0
End Sub

Private Sub GetRoutingDefaultRow(rowIndex As Long, ByRef vDest As Variant, ByRef vDetail As Variant, _
                                 ByRef vSpec As Variant, ByRef vAmt As Variant)
    Select Case rowIndex
        Case 1
            vDest = "Buy Plan": vDetail = "": vSpec = "Residual": vAmt = 0
        Case 2
            vDest = "Money Market": vDetail = GetSetting("DefaultMM", "CJTXX"): vSpec = "$": vAmt = 0
        Case Else
            vDest = "": vDetail = "": vSpec = "": vAmt = 0
    End Select
End Sub

Private Function BuildOtherRowsSum(specL As String, routedL As String, dataStartRow As Long, _
                                   dataEndRow As Long, excludeRow As Long) As String
    Dim parts As String
    Dim j As Long

    parts = ""
    For j = dataStartRow To dataEndRow
        If j <> excludeRow Then
            If parts <> "" Then parts = parts & "+"
            parts = parts & "IF(" & specL & j & "<>""Residual""," & routedL & j & ",0)"
        End If
    Next j

    BuildOtherRowsSum = parts
End Function

Private Sub AddRoutingHeaderCell(ws As Worksheet, rowNum As Long, colNum As Long, textValue As String)
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

Private Sub SetRangeLockedSafeRouting(targetRange As Range, isLocked As Boolean)
    On Error Resume Next
    targetRange.Locked = isLocked
    On Error GoTo 0
End Sub

Private Function FindPlanStatusColRouting(ws As Worksheet) As Long
    Dim c As Range
    Set c = FindCellExactRouting(ws, "Plan Status")
    If Not c Is Nothing Then FindPlanStatusColRouting = c.Column
End Function

Private Function FindCellExactRouting(ws As Worksheet, textValue As String) As Range
    On Error Resume Next
    Set FindCellExactRouting = ws.Cells.Find(What:=textValue, _
                                             LookIn:=xlValues, _
                                             LookAt:=xlWhole, _
                                             SearchOrder:=xlByRows, _
                                             SearchDirection:=xlNext, _
                                             MatchCase:=False)
    On Error GoTo 0
End Function

Private Function ColLetterRouting(colNum As Long) As String
    ColLetterRouting = Split(Columns(colNum).Address(, False), ":")(0)
End Function
