Attribute VB_Name = "CDS_Buy_Plans"
Option Explicit

' ============================================================
' CDS Buy Plans v1.4
'
' v1.4 changes vs v1.3:
'   - Keeps original scenario-funded AddBuyPlans workflow intact.
'   - Adds AddCashOnlyBuyPlan as a decoupled cash-deployment workflow.
'   - Cash-only workflow creates/refreshes a separate worksheet named
'     "Cash Buy Plan" and does NOT touch S1/S2 scenario columns.
'   - Cash-only worksheet now distinguishes:
'       Current CASH Holdings
'       Scenario New Cash
'       Cash Available to Deploy
'
' Original v1.3 behavior:
'   - All settings via CDS_Settings (BuyPlanRows, ScenarioStartCol,
'     ScenarioStride)
'   - ClassifyTicker UDF delegates to ClassifyTickerWithFallback
'   - PERSONAL.XLSB safety guard
'
' Adds 4-column "BUY PLAN" input area below each scenario's
' allocation pivot. Augments pivot with Post-Rebalance % column.
' Notes row flags new asset classes introduced by buys.
' Preserves user inputs across reruns (e.g. after SpawnScenario).
'
' Run AFTER scenarios are spawned/dialed in.
' Alt+F8 > AddBuyPlans > Run
'
' Cash-only workflow:
' Alt+F8 > AddCashOnlyBuyPlan > Run
' ============================================================

' Constants kept as fallbacks; runtime uses settings values
Private Const BUY_PLAN_INPUT_ROWS_DEFAULT As Long = 10

Public Function BuyPlanRows() As Long
    BuyPlanRows = CLng(GetSettingNum("BuyPlanRows", BUY_PLAN_INPUT_ROWS_DEFAULT))
End Function

Sub AddBuyPlans()
    On Error GoTo ErrHandler

    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB.", vbCritical
        GoTo Done
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    PrepareCDSWorksheetForMacro ws

    Dim scenCount As Long
    scenCount = CountScenarios(ws)
    If scenCount = 0 Then
        MsgBox "No scenarios found. Run AddRaiseCashScenarios first.", vbExclamation
        GoTo Done
    End If

    Dim totRow As Long, dataStart As Long, dataEnd As Long
    totRow = FindTotalRow(ws)
    If totRow = 0 Then
        MsgBox "Could not find totals row.", vbExclamation
        GoTo Done
    End If
    dataStart = 3
    dataEnd = totRow - 1

    Dim i As Long, firstS1Input As Long
    firstS1Input = 0
    For i = 1 To scenCount
        Dim inputStart As Long
        inputStart = BuildBuyPlanForScenario(ws, i, totRow, dataStart, dataEnd)
        If i = 1 Then firstS1Input = inputStart
        AugmentPivotForScenario ws, i, totRow, dataStart, dataEnd, inputStart
    Next i

    Application.Calculate

    If firstS1Input > 0 Then
        ws.Activate
        ws.Cells(firstS1Input, ScenStartCol()).Select
    End If

    MsgBox scenCount & " scenario(s) updated." & vbCrLf & vbCrLf & _
           "Buy plan: type Ticker and Amount. Asset Class and Yield auto-fill." & vbCrLf & _
           "Pivot shows Post-Sell % and Post-Reb % side by side." & vbCrLf & _
           "Notes row flags any buys into classes not in the pivot.", _
           vbInformation, "Buy Plans Ready"

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
' BUY PLAN AREA
' ============================================================
Private Function BuildBuyPlanForScenario(ws As Worksheet, scenNum As Long, _
                                         totRow As Long, dataStart As Long, dataEnd As Long) As Long
    BuildBuyPlanForScenario = 0

    Dim startCol As Long
    startCol = ScenStartCol() + (scenNum - 1) * ScenStride()

    Dim pivotEnd As Long
    pivotEnd = FindPivotGrandTotalRow(ws, startCol, totRow)
    If pivotEnd = 0 Then Exit Function

    Dim rowsCount As Long
    rowsCount = BuyPlanRows()

    ' Snapshot user-entered BUY PLAN rows before clearing/rebuilding the block.
    Dim priorHeaderRow As Long
    Dim hadPriorPlan As Boolean
    Dim priorInputStart As Long
    Dim priorTickers() As String
    Dim priorAmounts() As Variant
    Dim k As Long

    priorHeaderRow = FindBuyPlanHeaderRow(ws, startCol, pivotEnd)
    hadPriorPlan = (priorHeaderRow > 0)
    priorInputStart = 0

    ReDim priorTickers(0 To rowsCount - 1)
    ReDim priorAmounts(0 To rowsCount - 1)

    If hadPriorPlan Then
        priorInputStart = priorHeaderRow + 2

        For k = 0 To rowsCount - 1
            priorTickers(k) = Trim(CStr(ws.Cells(priorInputStart + k, startCol).Value))
            priorAmounts(k) = ws.Cells(priorInputStart + k, startCol + 1).Value
        Next k
    End If

    ' ============================================================
    ' SELL CONTEXT BLOCK
    ' ============================================================

    Dim contextTitleRow As Long
    Dim contextHeaderRow As Long
    Dim contextFirstDataRow As Long
    Dim contextLastDataRow As Long
    Dim contextTotalRow As Long
    Dim contextIncomeRow As Long
    Dim contextAvailableRow As Long
    Dim contextEndRow As Long

    contextTitleRow = pivotEnd + 2
    contextHeaderRow = contextTitleRow + 1
    contextFirstDataRow = contextHeaderRow + 1

    Dim sellCount As Long
    Dim r As Long
    Dim raiseAmt As Double

    sellCount = 0

    For r = dataStart To dataEnd
        raiseAmt = 0
        If IsNumeric(ws.Cells(r, startCol).Value) Then raiseAmt = CDbl(ws.Cells(r, startCol).Value)
        If Abs(raiseAmt) >= 0.5 Then sellCount = sellCount + 1
    Next r

    If sellCount = 0 Then
        contextLastDataRow = contextFirstDataRow
    Else
        contextLastDataRow = contextFirstDataRow + sellCount - 1
    End If

    contextTotalRow = contextLastDataRow + 1
    contextIncomeRow = contextTotalRow + 1
    contextAvailableRow = contextIncomeRow + 1
    contextEndRow = contextAvailableRow

    ' Clear old context + buy plan area generously.
    On Error Resume Next
    ws.Range(ws.Cells(contextTitleRow, startCol), ws.Cells(contextTitleRow + 80, startCol + 3)).UnMerge
    On Error GoTo 0
    ws.Range(ws.Cells(contextTitleRow, startCol), ws.Cells(contextTitleRow + 80, startCol + 3)).Clear

    ' Title
    ws.Range(ws.Cells(contextTitleRow, startCol), ws.Cells(contextTitleRow, startCol + 3)).Merge
    With ws.Cells(contextTitleRow, startCol)
        .Value = "S" & scenNum & " SELL CONTEXT"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    ' Headers
    Dim contextHeaders As Variant
    contextHeaders = Array("Ticker", "Description", "Sell $", "% Sold")

    Dim hi As Long
    For hi = 0 To 3
        With ws.Cells(contextHeaderRow, startCol + hi)
            .Value = contextHeaders(hi)
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            If hi >= 2 Then
                .HorizontalAlignment = xlRight
            Else
                .HorizontalAlignment = xlLeft
            End If
        End With
    Next hi

    Dim outRow As Long
    outRow = contextFirstDataRow

    Dim fmv As Double
    Dim descText As String
    Dim tickerText As String

    If sellCount = 0 Then
        With ws.Range(ws.Cells(outRow, startCol), ws.Cells(outRow, startCol + 3))
            .Merge
            .Value = "No sells entered yet for this scenario."
            .Font.Italic = True
            .Font.Color = RGB(120, 120, 120)
            .Interior.Color = RGB(242, 242, 242)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Else
        For r = dataStart To dataEnd
            raiseAmt = 0
            If IsNumeric(ws.Cells(r, startCol).Value) Then raiseAmt = CDbl(ws.Cells(r, startCol).Value)

            If Abs(raiseAmt) >= 0.5 Then
                tickerText = Trim(CStr(ws.Cells(r, 3).Value))
                descText = Trim(CStr(ws.Cells(r, 2).Value))
                fmv = 0
                If IsNumeric(ws.Cells(r, 5).Value) Then fmv = CDbl(ws.Cells(r, 5).Value)

                With ws.Cells(outRow, startCol)
                    .Value = tickerText
                    .Font.Bold = True
                    .Interior.Color = RGB(242, 242, 242)
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With

                With ws.Cells(outRow, startCol + 1)
                    .Value = descText
                    .Interior.Color = RGB(242, 242, 242)
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With

                With ws.Cells(outRow, startCol + 2)
                    .Value = raiseAmt
                    .NumberFormat = "$#,##0;($#,##0);""-"""
                    .Interior.Color = RGB(242, 242, 242)
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .HorizontalAlignment = xlRight
                End With

                With ws.Cells(outRow, startCol + 3)
                    If fmv <> 0 Then
                        .Value = raiseAmt / fmv
                    Else
                        .Value = 0
                    End If
                    .NumberFormat = "0.0%"
                    .Interior.Color = RGB(242, 242, 242)
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .HorizontalAlignment = xlRight
                End With

                outRow = outRow + 1
            End If
        Next r
    End If

    ' Context totals
    With ws.Cells(contextTotalRow, startCol)
        .Value = "Sell Total"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(contextTotalRow, startCol + 1)
        .Value = ""
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(contextTotalRow, startCol + 2)
        .Formula = "=" & ColLetter(startCol) & totRow
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(contextTotalRow, startCol + 3)
        .Formula = "=IF(E" & totRow & "=0,0," & ColLetter(startCol) & totRow & "/E" & totRow & ")"
        .NumberFormat = "0.0%"
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight
    End With

    With ws.Cells(contextIncomeRow, startCol)
        .Value = "Income Lost"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(contextIncomeRow, startCol + 1)
        .Value = ""
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(contextIncomeRow, startCol + 2)
        .Formula = "=" & ColLetter(startCol + 4) & totRow
        .Font.Bold = True
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(contextIncomeRow, startCol + 3)
        .Value = ""
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Cells(contextAvailableRow, startCol)
        .Value = "Available to Buy"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 176, 80)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(contextAvailableRow, startCol + 1)
        .Value = ""
        .Interior.Color = RGB(0, 176, 80)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    With ws.Cells(contextAvailableRow, startCol + 2)
        .Formula = "=" & ColLetter(startCol) & totRow
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .NumberFormat = "$#,##0;($#,##0);""-"""
        .Interior.Color = RGB(0, 176, 80)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(contextAvailableRow, startCol + 3)
        .Value = ""
        .Interior.Color = RGB(0, 176, 80)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' ============================================================
    ' BUY PLAN AREA
    ' ============================================================

    Dim headerRow As Long, tickAmtRow As Long
    Dim inputStart As Long, inputEnd As Long
    Dim totalRow As Long, incomeRow As Long, diffRow As Long, notesRow As Long

    headerRow = contextEndRow + 2
    tickAmtRow = headerRow + 1
    inputStart = tickAmtRow + 1
    inputEnd = inputStart + rowsCount - 1
    totalRow = inputEnd + 1
    incomeRow = totalRow + 1
    diffRow = incomeRow + 1
    notesRow = diffRow + 1

    ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow, startCol + 3)).Merge
    With ws.Cells(headerRow, startCol)
        .Value = "BUY PLAN"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    Dim hdrs As Variant
    hdrs = Array("Ticker", "Amount", "Asset Class", "Yield")

    For hi = 0 To 3
        With ws.Cells(tickAmtRow, startCol + hi)
            .Value = hdrs(hi)
            .Font.Bold = True
            .Interior.Color = RGB(0, 176, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            If hi = 0 Then
                .HorizontalAlignment = xlLeft
            ElseIf hi = 2 Then
                .HorizontalAlignment = xlCenter
            Else
                .HorizontalAlignment = xlRight
            End If
        End With
    Next hi

    Dim tickerCL As String, amtCL As String, classCL As String, yieldCL As String
    Dim macroPrefix As String
    tickerCL = ColLetter(startCol)
    amtCL = ColLetter(startCol + 1)
    classCL = ColLetter(startCol + 2)
    yieldCL = ColLetter(startCol + 3)
    macroPrefix = MacroWorkbookFormulaPrefix()

    For r = inputStart To inputEnd
        With ws.Cells(r, startCol)
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
        End With

        With ws.Cells(r, startCol + 1)
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .NumberFormat = "$#,##0;($#,##0);"
            .HorizontalAlignment = xlRight
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
        End With

        With ws.Cells(r, startCol + 2)
            .Formula = "=IF(" & tickerCL & r & "="""","""", " & _
                       "IFERROR(INDEX($A:$A, MATCH(" & tickerCL & r & ", $C:$C, 0)), " & _
                       macroPrefix & "ClassifyTicker(" & tickerCL & r & ")))"
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Color = RGB(80, 80, 80)
            .HorizontalAlignment = xlCenter
        End With

        With ws.Cells(r, startCol + 3)
            .Formula = "=IF(" & tickerCL & r & "="""",0, " & _
                       "IFERROR(INDEX($J:$J, MATCH(" & tickerCL & r & ", $C:$C, 0)), 0))"
            .NumberFormat = "0.00%;-0.00%;"
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Color = RGB(80, 80, 80)
            .HorizontalAlignment = xlRight
        End With
    Next r

    With ws.Cells(totalRow, startCol)
        .Value = "Total Buys"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 176, 80)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Cells(totalRow, startCol + 1)
        .Formula = "=SUM(" & amtCL & inputStart & ":" & amtCL & inputEnd & ")"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 176, 80)
        .NumberFormat = "$#,##0;($#,##0)"
        .HorizontalAlignment = xlRight
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    Dim ec As Long
    For ec = 2 To 3
        With ws.Cells(totalRow, startCol + ec)
            .Interior.Color = RGB(0, 176, 80)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next ec

    With ws.Cells(incomeRow, startCol)
        .Value = "Income Gained"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Cells(incomeRow, startCol + 1)
        .Formula = "=SUMPRODUCT(" & amtCL & inputStart & ":" & amtCL & inputEnd & "," & _
                   yieldCL & inputStart & ":" & yieldCL & inputEnd & ")"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .NumberFormat = "$#,##0;($#,##0)"
        .HorizontalAlignment = xlRight
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    For ec = 2 To 3
        With ws.Cells(incomeRow, startCol + ec)
            .Interior.Color = RGB(255, 242, 204)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next ec

    Dim raiseCellAddr As String
    raiseCellAddr = ColLetter(startCol) & totRow

    With ws.Cells(diffRow, startCol)
        .Value = "Diff vs Raise"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Cells(diffRow, startCol + 1)
        .Formula = "=" & raiseCellAddr & "-" & amtCL & totalRow
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .NumberFormat = "$#,##0"" to MM"";[Red]($#,##0)"" OVER"";""Fully Allocated"""
        .HorizontalAlignment = xlRight
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    For ec = 2 To 3
        With ws.Cells(diffRow, startCol + ec)
            .Interior.Color = RGB(255, 242, 204)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next ec

    Dim pivotClassFirstRow As Long, pivotClassLastRow As Long
    pivotClassFirstRow = totRow + 4
    pivotClassLastRow = pivotEnd - 1

    Dim buyClassRangeStr As String, pivotClassRangeStr As String
    buyClassRangeStr = "$" & classCL & "$" & inputStart & ":$" & classCL & "$" & inputEnd
    pivotClassRangeStr = "$" & ColLetter(startCol) & "$" & pivotClassFirstRow & ":$" & ColLetter(startCol) & "$" & pivotClassLastRow

    ws.Range(ws.Cells(notesRow, startCol), ws.Cells(notesRow, startCol + 3)).Merge
    With ws.Cells(notesRow, startCol)
        .Formula = "=" & macroPrefix & "MissingClassesIn(" & buyClassRangeStr & "," & pivotClassRangeStr & ")"
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
        .Font.Size = 9
        .Interior.Pattern = xlNone
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlNone
    End With

    ' Restore preserved inputs
    If hadPriorPlan Then
        For k = 0 To rowsCount - 1
            If priorTickers(k) <> "" Then
                ws.Cells(inputStart + k, startCol).Value = priorTickers(k)
            End If

            If IsNumeric(priorAmounts(k)) Then
                If CDbl(priorAmounts(k)) > 0 Then
                    ws.Cells(inputStart + k, startCol + 1).Value = CDbl(priorAmounts(k))
                End If
            End If
        Next k
    End If

    ws.Columns(ColLetter(startCol)).ColumnWidth = 13
    ws.Columns(ColLetter(startCol + 1)).ColumnWidth = 28
    ws.Columns(ColLetter(startCol + 2)).ColumnWidth = 14
    ws.Columns(ColLetter(startCol + 3)).ColumnWidth = 10

    BuildBuyPlanForScenario = inputStart
End Function

' ============================================================
' PIVOT AUGMENTATION
' ============================================================
Private Sub AugmentPivotForScenario(ws As Worksheet, scenNum As Long, _
                                    totRow As Long, dataStart As Long, dataEnd As Long, _
                                    buyInputStart As Long)
    Dim startCol As Long
    startCol = ScenStartCol() + (scenNum - 1) * ScenStride()

    Dim pivotEnd As Long
    pivotEnd = FindPivotGrandTotalRow(ws, startCol, totRow)
    If pivotEnd = 0 Then Exit Sub

    Dim pivotTitleRow As Long, pivotHeaderRow As Long
    pivotTitleRow = totRow + 2
    pivotHeaderRow = pivotTitleRow + 1

    If ws.Cells(pivotTitleRow, startCol).MergeArea.Columns.Count <> 3 Then
        On Error Resume Next
        ws.Range(ws.Cells(pivotTitleRow, startCol), ws.Cells(pivotTitleRow, startCol + 2)).UnMerge
        On Error GoTo 0
        ws.Range(ws.Cells(pivotTitleRow, startCol), ws.Cells(pivotTitleRow, startCol + 2)).Merge
        With ws.Range(ws.Cells(pivotTitleRow, startCol), ws.Cells(pivotTitleRow, startCol + 2))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
            .Interior.Color = RGB(0, 0, 0)
        End With
    End If

    With ws.Cells(pivotHeaderRow, startCol + 2)
        .Value = "Post-Reb %"
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlRight
    End With

    Dim raiseCL As String, buyClassCL As String, buyAmtCL As String, postRebCL As String
    raiseCL = ColLetter(startCol)
    buyAmtCL = ColLetter(startCol + 1)
    buyClassCL = ColLetter(startCol + 2)
    postRebCL = ColLetter(startCol + 2)

    Dim rowsCount As Long
    rowsCount = BuyPlanRows()

    Dim buyInputEnd As Long, buyTotalRow As Long
    buyInputEnd = buyInputStart + rowsCount - 1
    buyTotalRow = buyInputEnd + 1

    Dim r As Long
    For r = pivotHeaderRow + 1 To pivotEnd - 1
        Dim classRef As String
        classRef = "$" & raiseCL & "$" & r

        Dim formulaStr As String
        formulaStr = "=( " & _
            "SUMIF($A$" & dataStart & ":$A$" & dataEnd & ", " & classRef & _
                ", $E$" & dataStart & ":$E$" & dataEnd & ") " & _
            "- SUMIF($A$" & dataStart & ":$A$" & dataEnd & ", " & classRef & _
                ", $" & raiseCL & "$" & dataStart & ":$" & raiseCL & "$" & dataEnd & ") " & _
            "+ SUMIF($" & buyClassCL & "$" & buyInputStart & ":$" & buyClassCL & "$" & buyInputEnd & ", " & classRef & _
                ", $" & buyAmtCL & "$" & buyInputStart & ":$" & buyAmtCL & "$" & buyInputEnd & ") " & _
            ") / ( $E$" & totRow & " - $" & raiseCL & "$" & totRow & _
            " + $" & buyAmtCL & "$" & buyTotalRow & " )"

        With ws.Cells(r, startCol + 2)
            .Formula = formulaStr
            .NumberFormat = "0%"
            .Interior.Color = RGB(221, 235, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlRight
        End With
    Next r

    With ws.Cells(pivotEnd, startCol + 2)
        .Formula = "=SUM(" & postRebCL & (pivotHeaderRow + 1) & ":" & _
                   postRebCL & (pivotEnd - 1) & ")"
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
' UDF: ClassifyTicker (delegates to centralized classifier)
' ============================================================
Public Function ClassifyTicker(ticker As Variant) As String
    On Error Resume Next
    If IsError(ticker) Or IsEmpty(ticker) Then
        ClassifyTicker = "": Exit Function
    End If
    ClassifyTicker = ClassifyTickerWithFallback(CStr(ticker))
End Function

' ============================================================
' UDF: MissingClassesIn
' ============================================================
Public Function MissingClassesIn(buyClassRange As Range, pivotClassRange As Range) As String
    On Error Resume Next

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim cell As Range
    For Each cell In pivotClassRange
        Dim p As String
        p = UCase(Trim(CStr(cell.Value)))
        If p <> "" Then seen(p) = True
    Next cell

    Dim found As Object: Set found = CreateObject("Scripting.Dictionary")
    found.CompareMode = vbTextCompare

    Dim result As String
    For Each cell In buyClassRange
        Dim b As String
        b = UCase(Trim(CStr(cell.Value)))
        If b <> "" And Not seen.Exists(b) And Not found.Exists(b) Then
            found(b) = True
            If result <> "" Then result = result & ", "
            result = result & b
        End If
    Next cell

    If result = "" Then
        MissingClassesIn = ""
    Else
        MissingClassesIn = "+New class(es) introduced: " & result
    End If
End Function

' ============================================================
' HELPERS
' ============================================================

Private Function FindPivotGrandTotalRow(ws As Worksheet, startCol As Long, totRow As Long) As Long
    Dim r As Long
    FindPivotGrandTotalRow = 0
    For r = totRow + 2 To totRow + 50
        If UCase(Trim(CStr(ws.Cells(r, startCol).Value))) = "GRAND TOTAL" Then
            FindPivotGrandTotalRow = r: Exit Function
        End If
    Next r
End Function

Private Function FindTotalRow(ws As Worksheet) As Long
    Dim i As Long
    FindTotalRow = 0
    For i = 3 To 200
        If Trim(CStr(ws.Cells(i, 1).Value)) = "" And ws.Cells(i, 5).Value <> "" Then
            FindTotalRow = i: Exit Function
        End If
    Next i
End Function

Private Function CountScenarios(ws As Worksheet) As Long
    Dim i As Long, col As Long
    CountScenarios = 0
    For i = 0 To 30
        col = ScenStartCol() + i * ScenStride()
        If InStr(1, CStr(ws.Cells(2, col).Value), "Raise $", vbTextCompare) > 0 Then
            CountScenarios = i + 1
        Else
            Exit Function
        End If
    Next i
End Function

Private Function FindBuyPlanHeaderRow(ws As Worksheet, startCol As Long, pivotEnd As Long) As Long
    Dim r As Long

    FindBuyPlanHeaderRow = 0

    For r = pivotEnd + 1 To pivotEnd + 120
        If UCase(Trim(CStr(ws.Cells(r, startCol).Value))) = "BUY PLAN" Then
            FindBuyPlanHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function ColLetter(colNum As Long) As String
    ColLetter = Split(Columns(colNum).Address(, False), ":")(0)
End Function

Private Function MacroWorkbookFormulaPrefix() As String
    MacroWorkbookFormulaPrefix = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
End Function

' ============================================================
' CASH-ONLY BUY PLAN - SAFE VERSION
'
' Purpose:
'   "We already have cash. What do we buy?"
'
' Behavior:
'   - Creates/refreshes a separate worksheet named "Cash Buy Plan"
'   - Does NOT touch S1 / S2 scenario columns
'   - Does NOT modify scenario summaries
'   - Does NOT augment existing pivots
'   - Does NOT create sells
'   - Shows Current CASH Holdings and Scenario New Cash separately
'   - Uses Scenario New Cash as deployable cash when found
'   - User enters Ticker + Amount
'   - Asset Class and Yield auto-fill using existing holdings or ClassifyTicker
' ============================================================

Public Sub AddCashOnlyBuyPlan()
    On Error GoTo ErrHandler

    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If ActiveWorkbook Is Nothing Then GoTo Done

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB. Click into the client workbook first.", vbCritical
        GoTo Done
    End If

    If ActiveSheet.Name = "Cash Buy Plan" Then
        MsgBox "Click back into the processed holdings sheet first, then run Cash-Only Buy Plan.", vbExclamation
        GoTo Done
    End If

    Dim srcWs As Worksheet
    Set srcWs = ActiveSheet

    Dim totRow As Long
    Dim dataStart As Long
    Dim dataEnd As Long

    totRow = FindTotalRow(srcWs)

    If totRow = 0 Then
        MsgBox "Could not find totals row. Run Process Holdings first, then run Cash-Only Buy Plan.", vbExclamation
        GoTo Done
    End If

    dataStart = 3
    dataEnd = totRow - 1

    Dim outWs As Worksheet
    Set outWs = CDSCashBuy_GetOrCreateSheet(ActiveWorkbook, "Cash Buy Plan")

    BuildCashBuyPlanSheet srcWs, outWs, dataStart, dataEnd, totRow

    outWs.Activate
    outWs.Range("B7").Select

    Application.Calculate

    MsgBox "Cash Buy Plan created on its own worksheet." & vbCrLf & vbCrLf & _
           "This version does not touch S1, scenarios, or raise-cash summaries." & vbCrLf & _
           "Enter the target buy amount, then enter ticker/amount rows.", _
           vbInformation, "Cash Buy Plan Ready"

    GoTo Done

ErrHandler:
    MsgBox "Cash-Only Buy Plan failed: " & Err.Number & " - " & Err.Description, vbExclamation

Done:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
End Sub

Private Sub BuildCashBuyPlanSheet(srcWs As Worksheet, outWs As Worksheet, _
                                  dataStart As Long, dataEnd As Long, totRow As Long)

    Dim srcRef As String
    srcRef = CDSCashBuy_SheetRef(srcWs)

    Dim rowsCount As Long
    rowsCount = BuyPlanRows()

    Dim inputStart As Long
    Dim inputEnd As Long
    Dim macroPrefix As String

    inputStart = 12
    inputEnd = inputStart + rowsCount - 1
    macroPrefix = MacroWorkbookFormulaPrefix()

    ' Clear output sheet.
    outWs.Cells.Clear

    outWs.Cells.Font.Name = "Arial"
    outWs.Cells.Font.Size = 10

    ' Column widths.
    outWs.Columns("A").ColumnWidth = 22
    outWs.Columns("B").ColumnWidth = 18
    outWs.Columns("C").ColumnWidth = 38
    outWs.Columns("D").ColumnWidth = 10
    outWs.Columns("E").ColumnWidth = 14
    outWs.Columns("F").ColumnWidth = 28
    outWs.Columns("G").ColumnWidth = 16
    outWs.Columns("H").ColumnWidth = 16
    outWs.Columns("I").ColumnWidth = 16
    outWs.Columns("J").ColumnWidth = 16
    outWs.Columns("K").ColumnWidth = 10

    ' ============================================================
    ' HEADER
    ' ============================================================

    With outWs.Range("A1:K1")
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    outWs.Range("A1").Value = "CASH-ONLY BUY PLAN"
    outWs.Range("B1").Value = srcWs.Cells(1, 2).Value
    outWs.Range("C1").Value = srcWs.Cells(1, 1).Value

    outWs.Range("A3").Value = "Source Sheet"
    outWs.Range("B3").Value = srcWs.Name

    Dim baseCash As Double
    Dim scenarioCash As Variant

    baseCash = CDSCashBuy_BaseCash(srcWs, dataStart, dataEnd)
    scenarioCash = CDSCashBuy_FindScenarioNewCashBalance(srcWs)

    outWs.Range("A4").Value = "Current CASH Holdings"
    outWs.Range("B4").Value = baseCash
    outWs.Range("C4").Value = "Original CASH from holdings table"

    outWs.Range("A5").Value = "Scenario New Cash"
    If IsEmpty(scenarioCash) Then
        outWs.Range("B5").Value = ""
        outWs.Range("C5").Value = "No scenario cash found"
    Else
        outWs.Range("B5").Value = CDbl(scenarioCash)
        outWs.Range("C5").Value = "Pulled from Scenario Summary > New Cash Balance"
    End If

    outWs.Range("A6").Value = "Cash Available to Deploy"
    If IsEmpty(scenarioCash) Then
        outWs.Range("B6").Value = baseCash
        outWs.Range("C6").Value = "Using current CASH holdings"
    Else
        outWs.Range("B6").Value = CDbl(scenarioCash)
        outWs.Range("C6").Value = "Using scenario new cash"
    End If

    outWs.Range("A7").Value = "Target Buy Amount"
    outWs.Range("B7").Value = ""

    outWs.Range("A8").Value = "Total Buys"
    outWs.Range("B8").Formula = "=SUM(B" & inputStart & ":B" & inputEnd & ")"

    outWs.Range("A9").Value = "Difference vs Target"
    outWs.Range("B9").Formula = "=B7-B8"

    outWs.Range("A10").Value = "Cash Remaining"
    outWs.Range("B10").Formula = "=B6-B8"

    outWs.Range("A4:A10").Font.Bold = True
    outWs.Range("B4:B10").NumberFormat = "$#,##0;[Red]($#,##0);""-"""

    With outWs.Range("B7")
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(0, 0, 255)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With outWs.Range("A4:C10")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With outWs.Range("A6:B6")
        .Font.Bold = True
        .Interior.Color = RGB(198, 239, 206)
    End With

    With outWs.Range("A10:B10")
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With

    ' ============================================================
    ' BUY INPUT TABLE
    ' ============================================================

    Dim hdrs As Variant
    hdrs = Array("Ticker", "Amount", "Asset Class", "Yield", "Income", "Notes")

    Dim c As Long
    For c = 0 To UBound(hdrs)
        With outWs.Cells(11, 1 + c)
            .Value = hdrs(c)
            .Font.Bold = True
            .Interior.Color = RGB(0, 176, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            If c = 0 Or c = 5 Then
                .HorizontalAlignment = xlLeft
            Else
                .HorizontalAlignment = xlRight
            End If
        End With
    Next c

    Dim r As Long

    For r = inputStart To inputEnd
        ' Ticker input
        With outWs.Cells(r, 1)
            .Interior.Color = RGB(221, 235, 247)
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlLeft
        End With

        ' Amount input
        With outWs.Cells(r, 2)
            .Interior.Color = RGB(221, 235, 247)
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
            .NumberFormat = "$#,##0;($#,##0);"
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlRight
        End With

        ' Asset class
        With outWs.Cells(r, 3)
            .Formula = "=IF(A" & r & "="""","""",IFERROR(INDEX(" & srcRef & "$A:$A,MATCH(A" & r & "," & srcRef & "$C:$C,0))," & macroPrefix & "ClassifyTicker(A" & r & ")))"
            .Interior.Color = RGB(242, 242, 242)
            .Font.Color = RGB(80, 80, 80)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlCenter
        End With

        ' Yield
        With outWs.Cells(r, 4)
            .Formula = "=IF(A" & r & "="""",0,IFERROR(INDEX(" & srcRef & "$J:$J,MATCH(A" & r & "," & srcRef & "$C:$C,0)),0))"
            .NumberFormat = "0.00%;-0.00%;"
            .Interior.Color = RGB(242, 242, 242)
            .Font.Color = RGB(80, 80, 80)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlRight
        End With

        ' Income
        With outWs.Cells(r, 5)
            .Formula = "=B" & r & "*D" & r
            .NumberFormat = "$#,##0;($#,##0);"
            .Interior.Color = RGB(242, 242, 242)
            .Font.Color = RGB(80, 80, 80)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlRight
        End With

        ' Notes
        With outWs.Cells(r, 6)
            .Interior.Color = RGB(242, 242, 242)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlLeft
        End With
    Next r

    ' Totals row under buy table
    Dim totalRow As Long
    totalRow = inputEnd + 1

    With outWs.Range("A" & totalRow & ":F" & totalRow)
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    outWs.Cells(totalRow, 1).Value = "Total"
    outWs.Cells(totalRow, 2).Formula = "=SUM(B" & inputStart & ":B" & inputEnd & ")"
    outWs.Cells(totalRow, 2).NumberFormat = "$#,##0;($#,##0)"
    outWs.Cells(totalRow, 5).Formula = "=SUM(E" & inputStart & ":E" & inputEnd & ")"
    outWs.Cells(totalRow, 5).NumberFormat = "$#,##0;($#,##0)"

    ' ============================================================
    ' POST-BUY ALLOCATION SUMMARY
    ' ============================================================

    BuildCashBuyAllocationSummary srcWs, outWs, dataStart, dataEnd, totRow, inputStart, inputEnd
End Sub

Private Sub BuildCashBuyAllocationSummary(srcWs As Worksheet, outWs As Worksheet, _
                                          dataStart As Long, dataEnd As Long, totRow As Long, _
                                          inputStart As Long, inputEnd As Long)

    Dim srcRef As String
    srcRef = CDSCashBuy_SheetRef(srcWs)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim r As Long
    Dim cls As String

    For r = dataStart To dataEnd
        cls = UCase(Trim(CStr(srcWs.Cells(r, 1).Value)))

        If cls <> "" Then
            If Not seen.Exists(cls) Then seen.Add cls, cls
        End If
    Next r

    Dim startRow As Long
    startRow = 11

    Dim hdrs As Variant
    hdrs = Array("Asset Class", "Current $", "Buy $", "Post-Buy $", "Post-Buy %")

    Dim c As Long
    For c = 0 To UBound(hdrs)
        With outWs.Cells(startRow, 7 + c)
            .Value = hdrs(c)
            .Font.Bold = True
            .Interior.Color = RGB(0, 176, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            If c = 0 Then
                .HorizontalAlignment = xlLeft
            Else
                .HorizontalAlignment = xlRight
            End If
        End With
    Next c

    Dim outRow As Long
    outRow = startRow + 1

    ' Preferred ordering.
    If seen.Exists("CASH") Then
        WriteCashBuyAllocRow outWs, srcRef, "CASH", outRow, dataStart, dataEnd, totRow, inputStart, inputEnd
        outRow = outRow + 1
    End If

    If seen.Exists("BOND") Then
        WriteCashBuyAllocRow outWs, srcRef, "BOND", outRow, dataStart, dataEnd, totRow, inputStart, inputEnd
        outRow = outRow + 1
    End If

    If seen.Exists("???") Then
        WriteCashBuyAllocRow outWs, srcRef, "???", outRow, dataStart, dataEnd, totRow, inputStart, inputEnd
        outRow = outRow + 1
    End If

    Dim key As Variant
    For Each key In seen.Keys
        If UCase(CStr(key)) <> "CASH" And UCase(CStr(key)) <> "BOND" And UCase(CStr(key)) <> "???" Then
            WriteCashBuyAllocRow outWs, srcRef, CStr(key), outRow, dataStart, dataEnd, totRow, inputStart, inputEnd
            outRow = outRow + 1
        End If
    Next key

    ' Total row
    With outWs.Range("G" & outRow & ":K" & outRow)
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    outWs.Cells(outRow, 7).Value = "TOTAL"
    outWs.Cells(outRow, 8).Formula = "=SUM(H" & (startRow + 1) & ":H" & (outRow - 1) & ")"
    outWs.Cells(outRow, 9).Formula = "=SUM(I" & (startRow + 1) & ":I" & (outRow - 1) & ")"
    outWs.Cells(outRow, 10).Formula = "=SUM(J" & (startRow + 1) & ":J" & (outRow - 1) & ")"
    outWs.Cells(outRow, 11).Formula = "=SUM(K" & (startRow + 1) & ":K" & (outRow - 1) & ")"

    outWs.Range("H" & (startRow + 1) & ":J" & outRow).NumberFormat = "$#,##0;($#,##0)"
    outWs.Range("K" & (startRow + 1) & ":K" & outRow).NumberFormat = "0%"

    outWs.Range("G" & startRow & ":K" & outRow).Borders.LineStyle = xlContinuous
    outWs.Range("G" & startRow & ":K" & outRow).Borders.Weight = xlThin
End Sub

Private Sub WriteCashBuyAllocRow(outWs As Worksheet, srcRef As String, assetClass As String, _
                                 outRow As Long, dataStart As Long, dataEnd As Long, totRow As Long, _
                                 inputStart As Long, inputEnd As Long)

    outWs.Cells(outRow, 7).Value = assetClass

    ' Current $
    outWs.Cells(outRow, 8).Formula = "=SUMIF(" & srcRef & "$A$" & dataStart & ":$A$" & dataEnd & ",G" & outRow & "," & srcRef & "$E$" & dataStart & ":$E$" & dataEnd & ")"

    ' Buy $
    outWs.Cells(outRow, 9).Formula = "=SUMIF($C$" & inputStart & ":$C$" & inputEnd & ",G" & outRow & ",$B$" & inputStart & ":$B$" & inputEnd & ")"

    ' Post-Buy $
    ' If CASH, reduce by total buys from B8. Otherwise add buys into that asset class.
    outWs.Cells(outRow, 10).Formula = "=H" & outRow & "+I" & outRow & "-IF(G" & outRow & "=" & """CASH""" & ",$B$8,0)"

    ' Post-Buy %
    outWs.Cells(outRow, 11).Formula = "=IF(" & srcRef & "$E$" & totRow & "=0,0,J" & outRow & "/" & srcRef & "$E$" & totRow & ")"

    With outWs.Range("G" & outRow & ":K" & outRow)
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    outWs.Cells(outRow, 7).HorizontalAlignment = xlLeft
    outWs.Range("H" & outRow & ":K" & outRow).HorizontalAlignment = xlRight
End Sub

Private Function CDSCashBuy_BaseCash(srcWs As Worksheet, dataStart As Long, dataEnd As Long) As Double
    Dim r As Long
    Dim cls As String

    CDSCashBuy_BaseCash = 0

    For r = dataStart To dataEnd
        cls = UCase(Trim(CStr(srcWs.Cells(r, 1).Value)))

        If cls = "CASH" Then
            If IsNumeric(srcWs.Cells(r, 5).Value) Then
                CDSCashBuy_BaseCash = CDSCashBuy_BaseCash + CDbl(srcWs.Cells(r, 5).Value)
            End If
        End If
    Next r
End Function

Private Function CDSCashBuy_FindScenarioNewCashBalance(srcWs As Worksheet) As Variant
    Dim cell As Range
    Dim offsetCol As Long

    CDSCashBuy_FindScenarioNewCashBalance = Empty

    For Each cell In srcWs.UsedRange.Cells
        If UCase(Trim(CStr(cell.Value))) = "NEW CASH BALANCE" Then
            For offsetCol = 1 To 5
                If IsNumeric(cell.Offset(0, offsetCol).Value) Then
                    CDSCashBuy_FindScenarioNewCashBalance = CDbl(cell.Offset(0, offsetCol).Value)
                    Exit Function
                End If
            Next offsetCol
        End If
    Next cell
End Function

Private Function CDSCashBuy_GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=ActiveSheet)
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set CDSCashBuy_GetOrCreateSheet = ws
End Function

Private Function CDSCashBuy_SheetRef(ws As Worksheet) As String
    CDSCashBuy_SheetRef = "'" & Replace(ws.Name, "'", "''") & "'!"
End Function

