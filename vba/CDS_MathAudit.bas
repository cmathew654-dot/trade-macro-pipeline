Attribute VB_Name = "CDS_MathAudit"
Option Explicit

' ============================================================
' CDS Formula & Logic Audit v1.0
'
' Purpose:
'   Audits the active processed CDS workbook for formula/math integrity.
'
' Checks:
'   - Main holdings totals
'   - Holding % and % G/L formulas
'   - Base allocation pivot values
'   - S1 pro-rata baseline UX/model
'   - Scenario row formulas
'   - Scenario totals
'   - Scenario Post-Sell pivots
'   - Scenario summary tax/math metrics
'   - Sell Context block above Buy Plans
'   - Buy-plan totals
'   - Buy-plan income gained
'   - Buy-plan Diff vs Raise
'   - Buy-plan Post-Reb % formulas
'   - New-class / underallocation / overbuy risks
'
' Run:
'   Alt+F8 > AuditActiveCDSMath
'
' Output:
'   Creates a new workbook with a CDS_MATH_AUDIT sheet.
' ============================================================

Private Const EPS_DOLLARS As Double = 0.05
Private Const EPS_PERCENT As Double = 0.000001
Private Const CDS_PROTECT_PASSWORD As String = ""

' W11: routing-aware funding + PROCEEDS ROUTING block layout. Mirrors the
' constants in CDS_Routing.bas / CDS_Buy_Plans.bas ResolveScenarioFundingCellAddr.
Private Const ROUTING_TITLE_MA As String = "PROCEEDS ROUTING"
Private Const ROUTING_BUY_PLAN_DEST_MA As String = "Buy Plan"
Private Const ROUTING_PLAN_SCENARIO_NUM As Long = 2
Private Const ROUTING_TITLE_ROW_MA As Long = 10
Private Const ROUTING_DATA_START_ROW_MA As Long = 12
Private Const ROUTING_DATA_END_ROW_MA As Long = 16
Private Const ROUTING_STATUS_ROW_MA As Long = 17

Private mRows As Collection
Private mPassCount As Long
Private mFailCount As Long
Private mWarnCount As Long
Private mInfoCount As Long

Public Sub AuditActiveCDSMath()
    Dim srcWb As Workbook
    Dim ws As Worksheet
    Dim wasProtected As Boolean

    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook.", vbExclamation
        Exit Sub
    End If

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Click into the processed client workbook first. PERSONAL.XLSB is active.", vbExclamation
        Exit Sub
    End If

    Set srcWb = ActiveWorkbook
    Set ws = ActiveSheet

    InitAudit

    AuditLog "INFO", "Audit started", "Workbook=" & srcWb.Name & "; Sheet=" & ws.Name

    wasProtected = ws.ProtectContents

    On Error Resume Next
    ws.Unprotect Password:=CDS_PROTECT_PASSWORD
    On Error GoTo ErrHandler

    If Not IsProcessedCDSReport(ws) Then
        AuditLog "FAIL", "Processed CDS report detected", "Could not find ASSET CLASS header."
        WriteAuditResults srcWb.Name, ws.Name
        MsgBox "This sheet does not look like a processed CDS report.", vbExclamation
        GoTo CleanExit
    End If

    Application.CalculateFull

    Dim headerRow As Long
    Dim dataStart As Long
    Dim dataEnd As Long
    Dim totRow As Long

    headerRow = FindHeaderRow(ws)
    dataStart = headerRow + 1
    totRow = FindTotalRow(ws, dataStart)
    dataEnd = totRow - 1

    AuditAssertTrue "Header row found", headerRow > 0, "headerRow=" & headerRow
    AuditAssertTrue "Total row found", totRow > 0, "totRow=" & totRow

    If headerRow = 0 Or totRow = 0 Then
        WriteAuditResults srcWb.Name, ws.Name
        GoTo CleanExit
    End If

    AuditMainHoldings ws, dataStart, dataEnd, totRow
    AuditBaseAllocationPivot ws, dataStart, dataEnd
    AuditSellSpecs ws, dataStart, dataEnd, totRow
    AuditShortfallCoherence ws
    AuditRouting ws

    Dim scenCount As Long
    scenCount = CountScenariosLocal(ws)

    If scenCount = 0 Then
        AuditLog "INFO", "Scenario audit skipped", "No scenarios found."
    Else
        AuditS1ProRataModel ws, dataStart, dataEnd, totRow

        Dim sn As Long
        For sn = 1 To scenCount
            AuditScenarioBlock ws, sn, dataStart, dataEnd, totRow
            AuditSellContextForScenario ws, sn, dataStart, dataEnd, totRow
            AuditBuyPlanForScenario ws, sn, dataStart, dataEnd, totRow
        Next sn

        AuditScenarioSummary ws, scenCount, dataStart, dataEnd, totRow
    End If

    AuditWashFlag ws, dataStart, dataEnd, totRow, scenCount

    WriteAuditResults srcWb.Name, ws.Name

    srcWb.Activate
    ws.Activate

    MsgBox "CDS math audit complete." & vbCrLf & vbCrLf & _
           "PASS: " & mPassCount & vbCrLf & _
           "WARN: " & mWarnCount & vbCrLf & _
           "FAIL: " & mFailCount & vbCrLf & vbCrLf & _
           "Review the CDS_MATH_AUDIT results workbook.", _
           IIf(mFailCount = 0, vbInformation, vbExclamation), _
           "CDS Math Audit"

CleanExit:
    If wasProtected Then
        On Error Resume Next
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
        On Error GoTo 0
    End If

    Exit Sub

ErrHandler:
    Dim wbName As String
    Dim wsName As String

    If srcWb Is Nothing Then
        wbName = "(unknown)"
    Else
        wbName = srcWb.Name
    End If

    If ws Is Nothing Then
        wsName = "(unknown)"
    Else
        wsName = ws.Name
    End If

    AuditLog "FAIL", "Fatal audit error", "Error " & Err.Number & ": " & Err.Description
    WriteAuditResults wbName, wsName
    MsgBox "Audit failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ============================================================
' MAIN HOLDINGS AUDIT
' ============================================================

Private Sub AuditMainHoldings(ws As Worksheet, dataStart As Long, dataEnd As Long, totRow As Long)
    AuditLog "INFO", "Main holdings audit", "Rows " & dataStart & ":" & dataEnd

    Dim r As Long
    Dim totalFMV As Double
    Dim totalGL As Double
    Dim totalBasis As Double
    Dim totalIncome As Double
    Dim unknownCount As Long

    For r = dataStart To dataEnd
        totalFMV = totalFMV + Num(ws.Cells(r, 5).Value)
        totalGL = totalGL + Num(ws.Cells(r, 6).Value)
        totalBasis = totalBasis + Num(ws.Cells(r, 8).Value)
        totalIncome = totalIncome + Num(ws.Cells(r, 9).Value)

        If Trim(CStr(ws.Cells(r, 1).Value)) = "???" Then
            unknownCount = unknownCount + 1
        End If
    Next r

    If unknownCount > 0 Then
        AuditLog "FAIL", "Unknown asset classes remain", unknownCount & " row(s) still show ???."
    Else
        AuditLog "PASS", "Unknown asset classes remain", "None."
    End If

    AuditNear "Main total FMV equals sum of holdings", ws.Cells(totRow, 5).Value, totalFMV, EPS_DOLLARS
    AuditNear "Main total G/L equals sum of holdings", ws.Cells(totRow, 6).Value, totalGL, EPS_DOLLARS
    AuditNear "Main total cost basis equals sum of holdings", ws.Cells(totRow, 8).Value, totalBasis, EPS_DOLLARS
    AuditNear "Main total income equals sum of holdings", ws.Cells(totRow, 9).Value, totalIncome, EPS_DOLLARS

    If totalFMV <> 0 Then
        AuditNear "Main total % equals 100%", ws.Cells(totRow, 4).Value, 1#, EPS_PERCENT
        AuditNear "Main total % G/L equals total G/L / total FMV", ws.Cells(totRow, 7).Value, totalGL / totalFMV, EPS_PERCENT
        AuditNear "Main total yield equals total income / total FMV", ws.Cells(totRow, 10).Value, totalIncome / totalFMV, EPS_PERCENT
    Else
        AuditNear "Main total % with zero FMV", ws.Cells(totRow, 4).Value, 0#, EPS_PERCENT
    End If

    For r = dataStart To dataEnd
        Dim fmv As Double
        Dim gl As Double

        fmv = Num(ws.Cells(r, 5).Value)
        gl = Num(ws.Cells(r, 6).Value)

        If totalFMV <> 0 Then
            AuditNear "Row " & r & " holding %", ws.Cells(r, 4).Value, fmv / totalFMV, EPS_PERCENT
        Else
            AuditNear "Row " & r & " holding % with zero total FMV", ws.Cells(r, 4).Value, 0#, EPS_PERCENT
        End If

        If fmv <> 0 Then
            AuditNear "Row " & r & " % G/L", ws.Cells(r, 7).Value, gl / fmv, EPS_PERCENT
        Else
            AuditNear "Row " & r & " % G/L with zero FMV", ws.Cells(r, 7).Value, 0#, EPS_PERCENT
        End If

        If Trim(CStr(ws.Cells(r, 1).Value)) = "" Then
            AuditLog "WARN", "Blank asset class", "Row " & r & " has blank asset class."
        End If
    Next r

    AuditDuplicateTickers ws, dataStart, dataEnd
End Sub

Private Sub AuditDuplicateTickers(ws As Worksheet, dataStart As Long, dataEnd As Long)
    Dim d As Object
    Set d = NewDict()

    Dim r As Long
    Dim tk As String

    For r = dataStart To dataEnd
        tk = UCase(Trim(CStr(ws.Cells(r, 3).Value)))

        If tk <> "" Then
            If d.Exists(tk) Then
                d(tk) = d(tk) + 1
            Else
                d.Add tk, 1
            End If
        End If
    Next r

    Dim key As Variant

    For Each key In d.Keys
        If d(key) > 1 Then
            AuditLog "WARN", "Duplicate ticker in holdings", key & " appears " & d(key) & " times. Buy-plan class/yield lookup uses first match."
        End If
    Next key
End Sub

' ============================================================
' BASE ALLOCATION PIVOT AUDIT
' ============================================================

Private Sub AuditBaseAllocationPivot(ws As Worksheet, dataStart As Long, dataEnd As Long)
    AuditLog "INFO", "Base allocation pivot audit", "Pivot=AllocPivot"

    On Error GoTo ErrHandler

    Dim pvt As PivotTable
    Set pvt = ws.PivotTables("AllocPivot")

    pvt.PivotCache.Refresh
    pvt.RefreshTable

    Dim expected As Object
    Set expected = NewDict()

    Dim r As Long
    Dim cls As String

    For r = dataStart To dataEnd
        cls = UCase(Trim(CStr(ws.Cells(r, 1).Value)))

        If cls <> "" Then
            DictAdd expected, cls, Num(ws.Cells(r, 4).Value)
        End If
    Next r

    Dim tbl As Range
    Set tbl = pvt.TableRange1

    Dim seen As Object
    Set seen = NewDict()

    For r = 2 To tbl.Rows.Count - 1
        cls = UCase(Trim(CStr(tbl.Cells(r, 1).Value)))

        If cls <> "" Then
            seen(cls) = True
            AuditNear "AllocPivot class " & cls, tbl.Cells(r, 2).Value, DictVal(expected, cls), EPS_PERCENT
        End If
    Next r

    Dim key As Variant

    For Each key In expected.Keys
        If Not seen.Exists(CStr(key)) Then
            AuditLog "FAIL", "AllocPivot missing class", CStr(key)
        End If
    Next key

    AuditNear "AllocPivot grand total", tbl.Cells(tbl.Rows.Count, tbl.Columns.Count).Value, 1#, EPS_PERCENT
    Exit Sub

ErrHandler:
    AuditLog "FAIL", "AllocPivot audit failed", "Error " & Err.Number & ": " & Err.Description
End Sub

' ============================================================
' S1 PRO-RATA MODEL AUDIT
' ============================================================

Private Sub AuditS1ProRataModel(ws As Worksheet, dataStart As Long, dataEnd As Long, totRow As Long)
    AuditLog "INFO", "S1 pro-rata model audit", "S1 should be protected/formula-driven baseline."

    Dim s1Col As Long
    s1Col = ScenStartCol()

    Dim summaryCol As Long
    summaryCol = FindScenarioSummaryCol(ws)

    If summaryCol = 0 Then
        AuditLog "FAIL", "S1 summary column found", "SCENARIO SUMMARY not found."
        Exit Sub
    End If

    Dim targetRow As Long
    targetRow = FindRowByLabelInColumn(ws, summaryCol, "Target Raise")

    If targetRow = 0 Then
        AuditLog "FAIL", "S1 Target Raise row found", "Target Raise row not found."
        Exit Sub
    End If

    Dim targetRaise As Double
    targetRaise = Num(ws.Cells(targetRow, summaryCol + 1).Value)

    If ws.Cells(targetRow, summaryCol + 1).Locked Then
        AuditLog "FAIL", "S1 Target Raise input unlocked", "Scenario Summary S1 Target Raise cell is locked."
    Else
        AuditLog "PASS", "S1 Target Raise input unlocked", "Scenario Summary S1 Target Raise can be edited."
    End If

    Dim r As Long

    For r = dataStart To dataEnd
        If Not ws.Cells(r, s1Col).HasFormula Then
            AuditLog "FAIL", "S1 Raise $ formula protected", "Row " & r & " is not formula-driven."
        Else
            AuditLog "PASS", "S1 Raise $ formula protected", "Row " & r & " has formula."
        End If

        If Not ws.Cells(r, s1Col).Locked Then
            AuditLog "FAIL", "S1 Raise $ cell locked", "Row " & r & " is unlocked."
        Else
            AuditLog "PASS", "S1 Raise $ cell locked", "Row " & r & " is locked."
        End If

        If Num(ws.Cells(totRow, 5).Value) <> 0 Then
            AuditNear "S1 row " & r & " pro-rata raise amount", _
                      ws.Cells(r, s1Col).Value, _
                      targetRaise * (Num(ws.Cells(r, 5).Value) / Num(ws.Cells(totRow, 5).Value)), _
                      EPS_DOLLARS
        End If
    Next r
End Sub

' ============================================================
' SCENARIO AUDIT
' ============================================================

Private Sub AuditScenarioBlock(ws As Worksheet, scenNum As Long, dataStart As Long, dataEnd As Long, totRow As Long)
    Dim scenCol As Long
    scenCol = ScenStartCol() + (scenNum - 1) * ScenStride()

    AuditLog "INFO", "Scenario S" & scenNum & " audit", "StartCol=" & scenCol

    If InStr(1, CStr(ws.Cells(2, scenCol).Value), "Raise $", vbTextCompare) = 0 Then
        AuditLog "FAIL", "S" & scenNum & " header", "Expected Raise $ in row 2 col " & scenCol
        Exit Sub
    End If

    Dim r As Long
    Dim totalRaise As Double
    Dim totalGain As Double
    Dim totalNewFMV As Double
    Dim totalIncomeLost As Double

    For r = dataStart To dataEnd
        Dim fmv As Double
        Dim gl As Double
        Dim income As Double
        Dim raiseAmt As Double
        Dim expectedGain As Double
        Dim expectedNewFMV As Double
        Dim expectedIncomeLost As Double

        fmv = Num(ws.Cells(r, 5).Value)
        gl = Num(ws.Cells(r, 6).Value)
        income = Num(ws.Cells(r, 9).Value)
        raiseAmt = Num(ws.Cells(r, scenCol).Value)

        If raiseAmt < -EPS_DOLLARS Then
            AuditLog "FAIL", "S" & scenNum & " negative Raise $", "Row " & r & " has " & FormatCurrency(raiseAmt)
        End If

        If raiseAmt > fmv + EPS_DOLLARS Then
            AuditLog "FAIL", "S" & scenNum & " Raise $ exceeds FMV", _
                     "Row " & r & " ticker " & ws.Cells(r, 3).Value & _
                     " raise=" & FormatCurrency(raiseAmt) & _
                     "; FMV=" & FormatCurrency(fmv)
        End If

        If fmv = 0 And Abs(raiseAmt) > EPS_DOLLARS Then
            AuditLog "FAIL", "S" & scenNum & " Raise $ against zero-FMV holding", _
                     "Row " & r & " ticker " & ws.Cells(r, 3).Value
        End If

        If fmv = 0 Then
            expectedGain = 0
            expectedIncomeLost = 0
        Else
            expectedGain = raiseAmt * (gl / fmv)
            expectedIncomeLost = (raiseAmt / fmv) * income
        End If

        expectedNewFMV = fmv - raiseAmt

        totalRaise = totalRaise + raiseAmt
        totalGain = totalGain + expectedGain
        totalNewFMV = totalNewFMV + expectedNewFMV
        totalIncomeLost = totalIncomeLost + expectedIncomeLost
    Next r

    For r = dataStart To dataEnd
        fmv = Num(ws.Cells(r, 5).Value)
        gl = Num(ws.Cells(r, 6).Value)
        income = Num(ws.Cells(r, 9).Value)
        raiseAmt = Num(ws.Cells(r, scenCol).Value)

        If fmv = 0 Then
            expectedGain = 0
            expectedIncomeLost = 0
        Else
            expectedGain = raiseAmt * (gl / fmv)
            expectedIncomeLost = (raiseAmt / fmv) * income
        End If

        expectedNewFMV = fmv - raiseAmt

        AuditNear "S" & scenNum & " row " & r & " realized gains", ws.Cells(r, scenCol + 1).Value, expectedGain, EPS_DOLLARS
        AuditNear "S" & scenNum & " row " & r & " new FMV", ws.Cells(r, scenCol + 2).Value, expectedNewFMV, EPS_DOLLARS

        If totalNewFMV <> 0 Then
            AuditNear "S" & scenNum & " row " & r & " new %", ws.Cells(r, scenCol + 3).Value, expectedNewFMV / totalNewFMV, EPS_PERCENT
        Else
            AuditNear "S" & scenNum & " row " & r & " new % with zero new total", ws.Cells(r, scenCol + 3).Value, 0#, EPS_PERCENT
        End If

        AuditNear "S" & scenNum & " row " & r & " income lost", ws.Cells(r, scenCol + 4).Value, expectedIncomeLost, EPS_DOLLARS
    Next r

    AuditNear "S" & scenNum & " total raised", ws.Cells(totRow, scenCol).Value, totalRaise, EPS_DOLLARS
    AuditNear "S" & scenNum & " total realized gains", ws.Cells(totRow, scenCol + 1).Value, totalGain, EPS_DOLLARS
    AuditNear "S" & scenNum & " total new FMV", ws.Cells(totRow, scenCol + 2).Value, totalNewFMV, EPS_DOLLARS
    AuditNear "S" & scenNum & " total income lost", ws.Cells(totRow, scenCol + 4).Value, totalIncomeLost, EPS_DOLLARS

    If totalNewFMV <> 0 Then
        AuditNear "S" & scenNum & " total new %", ws.Cells(totRow, scenCol + 3).Value, 1#, EPS_PERCENT
    End If

    AuditScenarioPivot ws, scenNum, dataStart, dataEnd, totRow, scenCol
End Sub

Private Sub AuditScenarioPivot(ws As Worksheet, scenNum As Long, dataStart As Long, dataEnd As Long, totRow As Long, scenCol As Long)
    Dim grandRow As Long
    grandRow = FindScenarioPivotGrandTotalRow(ws, scenCol, totRow)

    If grandRow = 0 Then
        AuditLog "FAIL", "S" & scenNum & " Post-Sell pivot found", "No Grand Total row found."
        Exit Sub
    End If

    Dim expected As Object
    Set expected = NewDict()

    Dim r As Long
    Dim cls As String

    For r = dataStart To dataEnd
        cls = UCase(Trim(CStr(ws.Cells(r, 1).Value)))

        If cls <> "" And cls <> "???" Then
            DictAdd expected, cls, Num(ws.Cells(r, scenCol + 3).Value)
        End If
    Next r

    Dim seen As Object
    Set seen = NewDict()

    For r = totRow + 4 To grandRow - 1
        cls = UCase(Trim(CStr(ws.Cells(r, scenCol).Value)))

        If cls <> "" Then
            seen(cls) = True
            AuditNear "S" & scenNum & " Post-Sell pivot class " & cls, ws.Cells(r, scenCol + 1).Value, DictVal(expected, cls), EPS_PERCENT
        End If
    Next r

    Dim key As Variant

    For Each key In expected.Keys
        If Not seen.Exists(CStr(key)) Then
            AuditLog "FAIL", "S" & scenNum & " Post-Sell pivot missing class", CStr(key)
        End If
    Next key

    AuditNear "S" & scenNum & " Post-Sell pivot grand total", ws.Cells(grandRow, scenCol + 1).Value, 1#, EPS_PERCENT
End Sub

' ============================================================
' SCENARIO SUMMARY AUDIT
' ============================================================

Private Sub AuditScenarioSummary(ws As Worksheet, scenCount As Long, dataStart As Long, dataEnd As Long, totRow As Long)
    AuditLog "INFO", "Scenario summary audit", "Scenario count=" & scenCount

    Dim titleCell As Range
    Set titleCell = FindCellExact(ws, "SCENARIO SUMMARY")

    If titleCell Is Nothing Then
        AuditLog "FAIL", "Scenario summary found", "SCENARIO SUMMARY not found."
        Exit Sub
    End If

    Dim summaryCol As Long
    summaryCol = titleCell.Column

    Dim cgRateRow As Long
    Dim targetRow As Long
    Dim totalRaisedRow As Long
    Dim shortfallRow As Long
    Dim gainsRow As Long
    Dim taxRow As Long
    Dim netRow As Long
    Dim incomeLostRow As Long
    Dim newIncomeRow As Long
    Dim pctSoldRow As Long
    Dim cashRow As Long
    Dim runwayRow As Long

    cgRateRow = FindRowByLabelInColumn(ws, summaryCol, "Eff. CG Rate")
    targetRow = FindRowByLabelInColumn(ws, summaryCol, "Target Raise")
    totalRaisedRow = FindRowByLabelInColumn(ws, summaryCol, "Total Raised")
    shortfallRow = FindRowByLabelInColumn(ws, summaryCol, "Shortfall / Overage")
    gainsRow = FindRowByLabelInColumn(ws, summaryCol, "Realized Gains")
    taxRow = FindRowByLabelInColumn(ws, summaryCol, "Estimated Tax")
    netRow = FindRowByLabelInColumn(ws, summaryCol, "Net After Tax")
    incomeLostRow = FindRowByLabelInColumn(ws, summaryCol, "Income Lost / Year")
    newIncomeRow = FindRowByLabelInColumn(ws, summaryCol, "New Annual Income")
    pctSoldRow = FindRowByLabelInColumn(ws, summaryCol, "% Portfolio Sold")
    cashRow = FindRowByLabelInColumn(ws, summaryCol, "New Cash Balance")
    runwayRow = FindRowByLabelInColumn(ws, summaryCol, "Months of Runway")

    If cgRateRow = 0 Then AuditLog "FAIL", "Summary Eff. CG Rate row found", ""
    If targetRow = 0 Then AuditLog "FAIL", "Summary Target Raise row found", ""
    If totalRaisedRow = 0 Then AuditLog "FAIL", "Summary Total Raised row found", ""
    If gainsRow = 0 Then AuditLog "FAIL", "Summary Realized Gains row found", ""
    If taxRow = 0 Then AuditLog "FAIL", "Summary Estimated Tax row found", ""

    If cgRateRow = 0 Or targetRow = 0 Or totalRaisedRow = 0 Or gainsRow = 0 Or taxRow = 0 Then Exit Sub

    Dim cashFMV As Double
    cashFMV = SumClassFMV(ws, dataStart, dataEnd, "CASH")

    Dim totalFMV As Double
    totalFMV = Num(ws.Cells(totRow, 5).Value)

    Dim totalIncome As Double
    totalIncome = Num(ws.Cells(totRow, 9).Value)

    Dim monthlyDist As Double
    monthlyDist = FindMonthlyDistribution(ws, totRow)

    Dim sn As Long

    For sn = 1 To scenCount
        Dim scenCol As Long
        Dim sumCol As Long

        scenCol = ScenStartCol() + (sn - 1) * ScenStride()
        sumCol = summaryCol + sn

        Dim targetRaise As Double
        Dim totalRaised As Double
        Dim realizedGains As Double
        Dim taxRate As Double
        Dim estTax As Double
        Dim incomeLost As Double
        Dim newCash As Double

        targetRaise = Num(ws.Cells(targetRow, sumCol).Value)
        totalRaised = Num(ws.Cells(totRow, scenCol).Value)
        realizedGains = Num(ws.Cells(totRow, scenCol + 1).Value)
        taxRate = Num(ws.Cells(cgRateRow, sumCol).Value)
        estTax = realizedGains * taxRate
        incomeLost = Num(ws.Cells(totRow, scenCol + 4).Value)
        newCash = cashFMV + totalRaised

        AuditNear "S" & sn & " summary Total Raised", ws.Cells(totalRaisedRow, sumCol).Value, totalRaised, EPS_DOLLARS
        AuditNear "S" & sn & " summary Shortfall / Overage", ws.Cells(shortfallRow, sumCol).Value, totalRaised - targetRaise, EPS_DOLLARS
        AuditNear "S" & sn & " summary Realized Gains", ws.Cells(gainsRow, sumCol).Value, realizedGains, EPS_DOLLARS
        AuditNear "S" & sn & " summary Estimated Tax", ws.Cells(taxRow, sumCol).Value, estTax, EPS_DOLLARS
        AuditNear "S" & sn & " summary Net After Tax", ws.Cells(netRow, sumCol).Value, totalRaised - estTax, EPS_DOLLARS
        AuditNear "S" & sn & " summary Income Lost / Year", ws.Cells(incomeLostRow, sumCol).Value, incomeLost, EPS_DOLLARS
        AuditNear "S" & sn & " summary New Annual Income", ws.Cells(newIncomeRow, sumCol).Value, totalIncome - incomeLost, EPS_DOLLARS

        If totalFMV <> 0 Then
            AuditNear "S" & sn & " summary % Portfolio Sold", ws.Cells(pctSoldRow, sumCol).Value, totalRaised / totalFMV, EPS_PERCENT
        End If

        AuditNear "S" & sn & " summary New Cash Balance", ws.Cells(cashRow, sumCol).Value, newCash, EPS_DOLLARS

        If runwayRow > 0 And monthlyDist <> 0 Then
            AuditNear "S" & sn & " summary Months of Runway", ws.Cells(runwayRow, sumCol).Value, newCash / monthlyDist, 0.0001
        End If

        If realizedGains < -EPS_DOLLARS And estTax < -EPS_DOLLARS Then
            AuditLog "WARN", "S" & sn & " negative estimated tax", _
                     "Realized losses produce negative tax under current formula. Confirm whether this is intended."
        End If
    Next sn
End Sub

' ============================================================
' SELL CONTEXT AUDIT
' ============================================================

Private Sub AuditSellContextForScenario(ws As Worksheet, scenNum As Long, dataStart As Long, dataEnd As Long, totRow As Long)
    Dim scenCol As Long
    scenCol = ScenStartCol() + (scenNum - 1) * ScenStride()

    Dim grandRow As Long
    grandRow = FindScenarioPivotGrandTotalRow(ws, scenCol, totRow)

    If grandRow = 0 Then Exit Sub

    Dim buyPlanHeaderRow As Long
    buyPlanHeaderRow = FindBuyPlanHeaderRow_Audit(ws, scenCol, grandRow)

    If buyPlanHeaderRow = 0 Then
        AuditLog "INFO", "S" & scenNum & " sell context audit skipped", "No BUY PLAN block found."
        Exit Sub
    End If

    Dim contextTitleRow As Long
    contextTitleRow = FindSellContextTitleRow_Audit(ws, scenCol, grandRow, buyPlanHeaderRow)

    If contextTitleRow = 0 Then
        AuditLog "FAIL", "S" & scenNum & " Sell Context block found", "BUY PLAN exists but SELL CONTEXT block was not found."
        Exit Sub
    End If

    AuditLog "PASS", "S" & scenNum & " Sell Context block found", "Row=" & contextTitleRow

    Dim sellTotalRow As Long
    Dim incomeLostRow As Long
    Dim availableRow As Long

    sellTotalRow = FindLabelRowBetween(ws, scenCol, "Sell Total", contextTitleRow, buyPlanHeaderRow)
    incomeLostRow = FindLabelRowBetween(ws, scenCol, "Income Lost", contextTitleRow, buyPlanHeaderRow)
    availableRow = FindLabelRowBetween(ws, scenCol, "Available to Buy", contextTitleRow, buyPlanHeaderRow)

    If sellTotalRow = 0 Then
        AuditLog "FAIL", "S" & scenNum & " Sell Context Sell Total row", "Not found."
    Else
        AuditNear "S" & scenNum & " Sell Context Sell Total", ws.Cells(sellTotalRow, scenCol + 2).Value, Num(ws.Cells(totRow, scenCol).Value), EPS_DOLLARS
    End If

    If incomeLostRow = 0 Then
        AuditLog "FAIL", "S" & scenNum & " Sell Context Income Lost row", "Not found."
    Else
        AuditNear "S" & scenNum & " Sell Context Income Lost", ws.Cells(incomeLostRow, scenCol + 2).Value, Num(ws.Cells(totRow, scenCol + 4).Value), EPS_DOLLARS
    End If

    If availableRow = 0 Then
        AuditLog "FAIL", "S" & scenNum & " Sell Context Available to Buy row", "Not found."
    Else
        AuditNear "S" & scenNum & " Sell Context Available to Buy", ws.Cells(availableRow, scenCol + 2).Value, _
                  ResolveScenarioFundingValueForAudit(ws, scenNum, scenCol, totRow), EPS_DOLLARS
    End If
End Sub

' ============================================================
' BUY PLAN AUDIT
' ============================================================

Private Sub AuditBuyPlanForScenario(ws As Worksheet, scenNum As Long, dataStart As Long, dataEnd As Long, totRow As Long)
    Dim scenCol As Long
    scenCol = ScenStartCol() + (scenNum - 1) * ScenStride()

    Dim grandRow As Long
    grandRow = FindScenarioPivotGrandTotalRow(ws, scenCol, totRow)

    If grandRow = 0 Then Exit Sub

    Dim buyPlanHeaderRow As Long
    buyPlanHeaderRow = FindBuyPlanHeaderRow_Audit(ws, scenCol, grandRow)

    If buyPlanHeaderRow = 0 Then
        AuditLog "INFO", "S" & scenNum & " buy-plan audit skipped", "No BUY PLAN block found."
        Exit Sub
    End If

    AuditLog "INFO", "S" & scenNum & " buy-plan audit", "StartCol=" & scenCol & "; BuyPlanHeaderRow=" & buyPlanHeaderRow

    Dim rowsCount As Long
    rowsCount = BuyPlanRows()

    Dim inputStart As Long
    Dim inputEnd As Long
    Dim totalRow As Long
    Dim incomeRow As Long
    Dim diffRow As Long
    Dim notesRow As Long

    inputStart = buyPlanHeaderRow + 2
    inputEnd = inputStart + rowsCount - 1
    totalRow = inputEnd + 1
    incomeRow = totalRow + 1
    diffRow = incomeRow + 1
    notesRow = diffRow + 1

    Dim totalBuys As Double
    Dim incomeGained As Double

    Dim buyClassAmounts As Object
    Set buyClassAmounts = NewDict()

    Dim pivotClasses As Object
    Set pivotClasses = ScenarioPivotClassDict(ws, scenCol, totRow, grandRow)

    Dim r As Long

    For r = inputStart To inputEnd
        Dim ticker As String
        Dim amount As Double
        Dim actualClass As String
        Dim expectedClass As String
        Dim actualYield As Double
        Dim expectedYield As Double

        ticker = UCase(Trim(CStr(ws.Cells(r, scenCol).Value)))
        amount = Num(ws.Cells(r, scenCol + 1).Value)
        actualClass = UCase(Trim(CStr(ws.Cells(r, scenCol + 2).Value)))
        actualYield = Num(ws.Cells(r, scenCol + 3).Value)

        If ticker = "" And Abs(amount) > EPS_DOLLARS Then
            AuditLog "FAIL", "S" & scenNum & " buy amount with blank ticker", "Row " & r & "; amount=" & FormatCurrency(amount)
        End If

        If ticker <> "" And Abs(amount) <= EPS_DOLLARS Then
            AuditLog "WARN", "S" & scenNum & " buy ticker with no amount", "Row " & r & "; ticker=" & ticker
        End If

        If ticker <> "" And amount > EPS_DOLLARS Then
            expectedClass = ExpectedBuyClass(ws, ticker, dataStart, dataEnd)
            expectedYield = ExpectedBuyYield(ws, ticker, dataStart, dataEnd)

            If expectedClass = "" Then
                AuditLog "FAIL", "S" & scenNum & " buy ticker unclassified", "Ticker=" & ticker
            Else
                AuditText "S" & scenNum & " buy class " & ticker, actualClass, expectedClass
            End If

            AuditNear "S" & scenNum & " buy yield " & ticker, actualYield, expectedYield, EPS_PERCENT

            If expectedYield = 0 And FindTickerRow(ws, ticker, dataStart, dataEnd) = 0 Then
                AuditLog "WARN", "S" & scenNum & " new buy ticker has 0% yield", _
                         ticker & " is not currently held, so current logic assigns 0 yield. Income gained may be understated."
            End If

            totalBuys = totalBuys + amount
            incomeGained = incomeGained + amount * actualYield

            If actualClass <> "" Then
                DictAdd buyClassAmounts, actualClass, amount

                If Not pivotClasses.Exists(actualClass) Then
                    AuditLog "FAIL", "S" & scenNum & " Post-Reb pivot missing new buy class", _
                             "Class " & actualClass & " is bought but not present in scenario pivot. Pivot will not fully update."
                End If
            End If
        End If
    Next r

    AuditNear "S" & scenNum & " buy-plan total buys", ws.Cells(totalRow, scenCol + 1).Value, totalBuys, EPS_DOLLARS
    AuditNear "S" & scenNum & " buy-plan income gained", ws.Cells(incomeRow, scenCol + 1).Value, incomeGained, EPS_DOLLARS

    ' Funding basis: scenario's own Raise $ total, unless this is the plan
    ' scenario (S2) and a PROCEEDS ROUTING block exists, in which case the
    ' sheet's "Diff vs Raise" / "Available to Buy" cells are wired to the
    ' routing block's "Buy Plan" row Routed $ instead (CDS_Buy_Plans.bas
    ' ResolveScenarioFundingCellAddr). Mirror that resolution here.
    Dim fundingValue As Double
    fundingValue = ResolveScenarioFundingValueForAudit(ws, scenNum, scenCol, totRow)

    AuditNear "S" & scenNum & " buy-plan diff vs raise", ws.Cells(diffRow, scenCol + 1).Value, fundingValue - totalBuys, EPS_DOLLARS

    If totalBuys > fundingValue + EPS_DOLLARS Then
        AuditLog "FAIL", "S" & scenNum & " buy plan overallocated", _
                 "Buys exceed funding basis by " & FormatCurrency(totalBuys - fundingValue)
    End If

    If totalBuys < fundingValue - EPS_DOLLARS Then
        AuditLog "WARN", "S" & scenNum & " buy plan underallocated", _
                 "Buys are short of funding basis by " & FormatCurrency(fundingValue - totalBuys) & _
                 ". Confirm residual proceeds are intentionally going to money market/cash."
    End If

    AuditPostRebalancePivot ws, scenNum, dataStart, dataEnd, totRow, scenCol, grandRow, buyClassAmounts, totalBuys

    If Trim(CStr(ws.Cells(notesRow, scenCol).Value)) <> "" Then
        AuditLog "INFO", "S" & scenNum & " buy-plan notes", CStr(ws.Cells(notesRow, scenCol).Value)
    End If
End Sub

Private Sub AuditPostRebalancePivot(ws As Worksheet, scenNum As Long, dataStart As Long, dataEnd As Long, _
                                    totRow As Long, scenCol As Long, grandRow As Long, _
                                    buyClassAmounts As Object, totalBuys As Double)

    If UCase(Trim(CStr(ws.Cells(totRow + 3, scenCol + 2).Value))) <> "POST-REB %" Then
        AuditLog "INFO", "S" & scenNum & " Post-Reb audit skipped", "Post-Reb % column not present."
        Exit Sub
    End If

    Dim totalFMV As Double
    Dim totalRaised As Double
    Dim denominator As Double

    totalFMV = Num(ws.Cells(totRow, 5).Value)
    totalRaised = Num(ws.Cells(totRow, scenCol).Value)
    denominator = totalFMV - totalRaised + totalBuys

    If denominator = 0 Then
        AuditLog "FAIL", "S" & scenNum & " Post-Reb denominator", "Denominator is zero."
        Exit Sub
    End If

    Dim r As Long
    Dim cls As String
    Dim expected As Double

    For r = totRow + 4 To grandRow - 1
        cls = UCase(Trim(CStr(ws.Cells(r, scenCol).Value)))

        If cls <> "" Then
            expected = (SumClassFMV(ws, dataStart, dataEnd, cls) _
                        - SumClassScenarioRaise(ws, dataStart, dataEnd, cls, scenCol) _
                        + DictVal(buyClassAmounts, cls)) / denominator

            AuditNear "S" & scenNum & " Post-Reb class " & cls, ws.Cells(r, scenCol + 2).Value, expected, EPS_PERCENT
        End If
    Next r

    Dim expectedGrand As Double
    expectedGrand = 0

    For r = totRow + 4 To grandRow - 1
        expectedGrand = expectedGrand + Num(ws.Cells(r, scenCol + 2).Value)
    Next r

    AuditNear "S" & scenNum & " Post-Reb grand total equals visible class sum", ws.Cells(grandRow, scenCol + 2).Value, expectedGrand, EPS_PERCENT

    If Abs(totalBuys - totalRaised) <= EPS_DOLLARS Then
        AuditNear "S" & scenNum & " Post-Reb grand total equals 100% when fully allocated", ws.Cells(grandRow, scenCol + 2).Value, 1#, EPS_PERCENT
    Else
        AuditLog "WARN", "S" & scenNum & " Post-Reb not full-portfolio allocation", _
                 "Total buys differ from total raised. Residual proceeds need explicit review."
    End If
End Sub

' ============================================================
' SELL SPEC AUDIT (W1/W11)
'
' Recomputes each row's Manual Sell $ from (mode, Amt Type, Amount, FMV,
' Qty, report total FMV) per the workbench's own formula semantics
' (CDS_Sell_Workbench.bas BuildWorkbenchRows manualCol formula) and flags
' any row where the sheet value drifts from that recomputation.
' ============================================================
Private Sub AuditSellSpecs(ws As Worksheet, dataStart As Long, dataEnd As Long, totRow As Long)
    Dim modeCol As Long
    modeCol = FindWorkbenchModeCol(ws)

    If modeCol = 0 Then
        AuditLog "INFO", "Sell spec audit skipped", "No Sell Workbench found."
        Exit Sub
    End If

    Dim specTypeCol As Long
    Dim specAmtCol As Long
    Dim manualCol As Long

    specTypeCol = modeCol + 1
    specAmtCol = modeCol + 2
    manualCol = modeCol + 3

    Dim totalFMV As Double
    totalFMV = Num(ws.Cells(totRow, 5).Value)

    Dim r As Long
    Dim mismatchCount As Long
    Dim firstMismatchDetail As String

    For r = dataStart To dataEnd
        Dim modeVal As String
        Dim typeVal As String
        Dim amtVal As Double
        Dim fmv As Double
        Dim qty As Double
        Dim expectedManual As Double
        Dim actualManual As Double

        modeVal = Trim(CStr(ws.Cells(r, modeCol).Value))
        typeVal = Trim(CStr(ws.Cells(r, specTypeCol).Value))
        amtVal = Num(ws.Cells(r, specAmtCol).Value)
        fmv = Num(ws.Cells(r, 5).Value)
        qty = Num(ws.Cells(r, 11).Value)

        If modeVal <> "Manual" Then
            expectedManual = 0
        ElseIf typeVal = "ALL" Then
            expectedManual = fmv
        ElseIf typeVal = "Shares" Then
            If qty <> 0 Then
                expectedManual = amtVal * (fmv / qty)
            Else
                expectedManual = 0
            End If
        ElseIf typeVal = "% Pos" Then
            expectedManual = amtVal / 100 * fmv
        ElseIf typeVal = "% Acct" Then
            expectedManual = amtVal / 100 * totalFMV
        Else
            expectedManual = amtVal
        End If

        actualManual = Num(ws.Cells(r, manualCol).Value)

        If Abs(actualManual - expectedManual) > EPS_DOLLARS Then
            mismatchCount = mismatchCount + 1
            If firstMismatchDetail = "" Then
                firstMismatchDetail = "Row " & r & " ticker " & ws.Cells(r, 3).Value & _
                    ": Actual=" & FormatCurrency(actualManual) & "; Expected=" & FormatCurrency(expectedManual)
            End If
        End If
    Next r

    If mismatchCount = 0 Then
        AuditLog "PASS", "Sell spec Manual Sell $ matches mode/type/amount", "Rows " & dataStart & ":" & dataEnd & " all match."
    Else
        AuditLog "FAIL", "Sell spec Manual Sell $ matches mode/type/amount", _
                 mismatchCount & " row(s) mismatch. First: " & firstMismatchDetail
    End If
End Sub

' ============================================================
' SHORTFALL COHERENCE AUDIT (W4/W11)
'
' The workbench status cell (row 3) shows a "SHORTFALL: ..." message via a
' sheet formula/format condition when Target - TotalProposed > 0.5
' (CDS_Sell_Workbench.bas ApplyWorkbenchStatus). Confirm the displayed
' state agrees with an independent recomputation from the same two cells.
' ============================================================
Private Sub AuditShortfallCoherence(ws As Worksheet)
    Dim modeCol As Long
    modeCol = FindWorkbenchModeCol(ws)

    If modeCol = 0 Then
        AuditLog "INFO", "Shortfall coherence check skipped", "No Sell Workbench found."
        Exit Sub
    End If

    Dim targetInputCol As Long
    Dim statusCol As Long
    targetInputCol = modeCol - 2
    statusCol = modeCol + 7

    Dim targetVal As Double
    Dim totalProposed As Double
    Dim statusText As String

    targetVal = Num(ws.Cells(2, targetInputCol).Value)
    totalProposed = Num(ws.Cells(6, statusCol + 1).Value)
    statusText = CStr(ws.Cells(3, statusCol).Value)

    Dim statusShowsShortfall As Boolean
    Dim conditionShortfall As Boolean

    statusShowsShortfall = (Left(statusText, 9) = "SHORTFALL")
    conditionShortfall = (targetVal - totalProposed > 0.5)

    If statusShowsShortfall = conditionShortfall Then
        AuditLog "PASS", "Shortfall status agrees with target vs proposed", _
                 "Status shortfall=" & statusShowsShortfall & "; Target-Proposed=" & FormatCurrency(targetVal - totalProposed)
    Else
        AuditLog "WARN", "Shortfall status agrees with target vs proposed", _
                 "Status shortfall=" & statusShowsShortfall & " but Target-Proposed=" & FormatCurrency(targetVal - totalProposed)
    End If
End Sub

' ============================================================
' PROCEEDS ROUTING AUDIT (W2/W11)
'
' Validates the PROCEEDS ROUTING block (CDS_Routing.bas WriteRoutingBlock),
' which lives at the same statusCol as the workbench status block starting
' at ROUTING_TITLE_ROW_MA. Skips entirely (INFO) when no workbench or no
' routing block is present on the sheet.
' ============================================================
Private Sub AuditRouting(ws As Worksheet)
    Dim modeCol As Long
    modeCol = FindWorkbenchModeCol(ws)

    If modeCol = 0 Then
        AuditLog "INFO", "Routing audit skipped", "No Sell Workbench found."
        Exit Sub
    End If

    Dim statusCol As Long
    statusCol = modeCol + 7

    If UCase(Trim(CStr(ws.Cells(ROUTING_TITLE_ROW_MA, statusCol).Value))) <> UCase(ROUTING_TITLE_MA) Then
        AuditLog "INFO", "Routing audit skipped", "No PROCEEDS ROUTING block found."
        Exit Sub
    End If

    Dim destCol As Long, specCol As Long, routedCol As Long
    destCol = statusCol
    specCol = statusCol + 2
    routedCol = statusCol + 4

    Dim proceedsTotal As Double
    proceedsTotal = Num(ws.Cells(6, statusCol + 1).Value)

    Dim r As Long
    Dim residualCount As Long
    Dim sumRouted As Double
    Dim negativeCount As Long
    Dim negativeDetail As String

    For r = ROUTING_DATA_START_ROW_MA To ROUTING_DATA_END_ROW_MA
        Dim destVal As String
        Dim specVal As String
        Dim routedVal As Double

        destVal = Trim(CStr(ws.Cells(r, destCol).Value))
        specVal = Trim(CStr(ws.Cells(r, specCol).Value))
        routedVal = Num(ws.Cells(r, routedCol).Value)

        If destVal <> "" And specVal = "Residual" Then
            residualCount = residualCount + 1
        End If

        sumRouted = sumRouted + routedVal

        If routedVal < -EPS_DOLLARS Then
            negativeCount = negativeCount + 1
            If negativeDetail = "" Then negativeDetail = "Row " & r & ": " & FormatCurrency(routedVal)
        End If
    Next r

    If residualCount = 1 Then
        AuditLog "PASS", "Routing has exactly one Residual row", "Count=1."
    Else
        AuditLog "FAIL", "Routing has exactly one Residual row", _
                 "Found " & residualCount & " Residual row(s) among rows with a Destination."
    End If

    AuditNear "Routing sum of Routed $ matches Total Proposed", sumRouted, proceedsTotal, 0.5

    If negativeCount = 0 Then
        AuditLog "PASS", "Routing Routed $ non-negative", "None negative."
    Else
        AuditLog "FAIL", "Routing Routed $ non-negative", negativeCount & " row(s) negative. First: " & negativeDetail
    End If

    Dim diff As Double
    diff = sumRouted - proceedsTotal

    Dim expectedStatus As String

    If residualCount <> 1 Then
        expectedStatus = "Need exactly one Residual row"
    ElseIf Abs(diff) > 0.5 Then
        If diff > 0 Then
            expectedStatus = "Routing over-allocates by $" & Format(diff, "#,##0")
        Else
            expectedStatus = "Routing under-allocates by $" & Format(-diff, "#,##0")
        End If
    Else
        expectedStatus = "Routing OK"
    End If

    Dim actualStatus As String
    actualStatus = Trim(CStr(ws.Cells(ROUTING_STATUS_ROW_MA, destCol).Value))

    If actualStatus = expectedStatus Then
        AuditLog "PASS", "Routing status cell agrees with computed condition", "Status=" & actualStatus
    Else
        AuditLog "WARN", "Routing status cell agrees with computed condition", _
                 "Cell=" & actualStatus & "; Computed=" & expectedStatus
    End If
End Sub

' ============================================================
' WASH-SALE FLAG AUDIT (W8)
'
' Warning-level only: for every report row with Proposed Sell $ > 0 and a
' loss (G/L < 0), checks every scenario's BUY PLAN ticker cells for the
' same ticker. On a match, logs a WARN and colors that buy-plan ticker
' cell orange so it stands out on the sheet. Recoloring the same matched
' cells the same color on every run keeps this idempotent.
' ============================================================
Private Sub AuditWashFlag(ws As Worksheet, dataStart As Long, dataEnd As Long, totRow As Long, scenCount As Long)
    Dim modeCol As Long
    modeCol = FindWorkbenchModeCol(ws)

    If modeCol = 0 Then
        AuditLog "INFO", "Wash-sale flag audit skipped", "No Sell Workbench found."
        Exit Sub
    End If

    If scenCount = 0 Then
        AuditLog "INFO", "Wash-sale flag audit skipped", "No scenarios/buy plans found."
        Exit Sub
    End If

    Dim proposedCol As Long
    proposedCol = modeCol + 4

    Dim r As Long
    Dim washCount As Long

    For r = dataStart To dataEnd
        Dim proposedAmt As Double
        Dim gl As Double
        Dim ticker As String

        proposedAmt = Num(ws.Cells(r, proposedCol).Value)
        gl = Num(ws.Cells(r, 6).Value)
        ticker = UCase(Trim(CStr(ws.Cells(r, 3).Value)))

        If proposedAmt > EPS_DOLLARS And gl < -EPS_DOLLARS And ticker <> "" Then
            Dim sn As Long

            For sn = 1 To scenCount
                Dim scenCol As Long
                Dim grandRow As Long
                Dim buyPlanHeaderRow As Long

                scenCol = ScenStartCol() + (sn - 1) * ScenStride()
                grandRow = FindScenarioPivotGrandTotalRow(ws, scenCol, totRow)

                If grandRow > 0 Then
                    buyPlanHeaderRow = FindBuyPlanHeaderRow_Audit(ws, scenCol, grandRow)

                    If buyPlanHeaderRow > 0 Then
                        Dim inputStart As Long
                        Dim inputEnd As Long
                        Dim br As Long

                        inputStart = buyPlanHeaderRow + 2
                        inputEnd = inputStart + BuyPlanRows() - 1

                        For br = inputStart To inputEnd
                            If UCase(Trim(CStr(ws.Cells(br, scenCol).Value))) = ticker Then
                                AuditLog "WARN", "Wash-sale risk", _
                                         "Wash-sale risk: loss sale of " & ticker & " reappears in S" & sn & " buy plan"
                                ws.Cells(br, scenCol).Interior.Color = RGB(255, 192, 96)
                                washCount = washCount + 1
                            End If
                        Next br
                    End If
                End If
            Next sn
        End If
    Next r

    If washCount = 0 Then
        AuditLog "PASS", "Wash-sale flag check", "No loss-sale tickers reappear in any buy plan."
    End If
End Sub

' ============================================================
' ROUTING-AWARE FUNDING RESOLUTION (W2/W11)
'
' Mirrors CDS_Buy_Plans.bas's ResolveScenarioFundingCellAddr (Private to
' that module, so reimplemented here by header text lookup rather than
' shared): when auditing the plan scenario (S2) and a PROCEEDS ROUTING
' block exists, the funding basis for "Available to Buy" / "Diff vs
' Raise" is the routing block's "Buy Plan" row Routed $ value instead of
' the scenario's own Raise $ total.
' ============================================================
Private Function ResolveScenarioFundingValueForAudit(ws As Worksheet, scenNum As Long, scenCol As Long, totRow As Long) As Double
    ResolveScenarioFundingValueForAudit = Num(ws.Cells(totRow, scenCol).Value)

    If scenNum <> ROUTING_PLAN_SCENARIO_NUM Then Exit Function

    Dim c As Range
    Set c = FindCellExact(ws, ROUTING_TITLE_MA)
    If c Is Nothing Then Exit Function

    Dim dataStart As Long, dataEnd As Long, r As Long
    dataStart = c.Row + 2
    dataEnd = dataStart + 4

    For r = dataStart To dataEnd
        If UCase(Trim(CStr(ws.Cells(r, c.Column).Value))) = UCase(ROUTING_BUY_PLAN_DEST_MA) Then
            ResolveScenarioFundingValueForAudit = Num(ws.Cells(r, c.Column + 4).Value)
            Exit Function
        End If
    Next r
End Function

' ============================================================
' WORKBENCH ANCHOR RESOLUTION (W1/W2/W8/W11)
'
' Sell Workbench columns sit at fixed offsets from workCol
' (BuildSellWorkbenchOnSheet, CDS_Sell_Workbench.bas):
'   targetInputCol=workCol+1, modeCol=workCol+3, specTypeCol=workCol+4,
'   specAmtCol=workCol+5, manualCol=workCol+6, proposedCol=workCol+7,
'   usedCol=workCol+8, statusCol=workCol+10. workCol itself moves with
'   scenario count, so audits anchor on the "Sell Mode" header text
'   (unique on the sheet) and use offsets relative to it instead of any
'   hardcoded column number.
' ============================================================
Private Function FindWorkbenchModeCol(ws As Worksheet) As Long
    Dim c As Range
    Set c = FindCellExact(ws, "Sell Mode")
    If Not c Is Nothing Then FindWorkbenchModeCol = c.Column
End Function

' ============================================================
' LOOKUP HELPERS
' ============================================================

Private Function ExpectedBuyClass(ws As Worksheet, ticker As String, dataStart As Long, dataEnd As Long) As String
    Dim rowFound As Long
    rowFound = FindTickerRow(ws, ticker, dataStart, dataEnd)

    If rowFound > 0 Then
        ExpectedBuyClass = UCase(Trim(CStr(ws.Cells(rowFound, 1).Value)))
    Else
        ExpectedBuyClass = UCase(Trim(CStr(ClassifyTickerWithFallback(ticker))))
    End If
End Function

Private Function ExpectedBuyYield(ws As Worksheet, ticker As String, dataStart As Long, dataEnd As Long) As Double
    Dim rowFound As Long
    rowFound = FindTickerRow(ws, ticker, dataStart, dataEnd)

    If rowFound > 0 Then
        ExpectedBuyYield = Num(ws.Cells(rowFound, 10).Value)
    Else
        ExpectedBuyYield = 0
    End If
End Function

Private Function SumClassFMV(ws As Worksheet, dataStart As Long, dataEnd As Long, assetClass As String) As Double
    Dim r As Long

    For r = dataStart To dataEnd
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = UCase(assetClass) Then
            SumClassFMV = SumClassFMV + Num(ws.Cells(r, 5).Value)
        End If
    Next r
End Function

Private Function SumClassScenarioRaise(ws As Worksheet, dataStart As Long, dataEnd As Long, assetClass As String, scenCol As Long) As Double
    Dim r As Long

    For r = dataStart To dataEnd
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = UCase(assetClass) Then
            SumClassScenarioRaise = SumClassScenarioRaise + Num(ws.Cells(r, scenCol).Value)
        End If
    Next r
End Function

Private Function FindMonthlyDistribution(ws As Worksheet, totRow As Long) As Double
    Dim r As Long

    For r = totRow + 1 To totRow + 50
        If InStr(1, CStr(ws.Cells(r, 9).Value), "MONTHLY DISTRIBUTION", vbTextCompare) > 0 Then
            FindMonthlyDistribution = Num(ws.Cells(r, 10).Value)
            Exit Function
        End If
    Next r
End Function

Private Function ScenarioPivotClassDict(ws As Worksheet, scenCol As Long, totRow As Long, grandRow As Long) As Object
    Dim d As Object
    Set d = NewDict()

    Dim r As Long
    Dim cls As String

    For r = totRow + 4 To grandRow - 1
        cls = UCase(Trim(CStr(ws.Cells(r, scenCol).Value)))

        If cls <> "" Then
            d(cls) = True
        End If
    Next r

    Set ScenarioPivotClassDict = d
End Function

Private Function FindTickerRow(ws As Worksheet, ticker As String, startRow As Long, endRow As Long) As Long
    Dim r As Long
    Dim tk As String

    tk = UCase(Trim(ticker))

    For r = startRow To endRow
        If UCase(Trim(CStr(ws.Cells(r, 3).Value))) = tk Then
            FindTickerRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long

    For r = 1 To 10
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindTotalRow(ws As Worksheet, dataStart As Long) As Long
    Dim r As Long

    For r = dataStart To dataStart + 500
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" And Trim(CStr(ws.Cells(r, 5).Value)) <> "" Then
            FindTotalRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindScenarioPivotGrandTotalRow(ws As Worksheet, scenCol As Long, totRow As Long) As Long
    Dim r As Long

    For r = totRow + 2 To totRow + 150
        If UCase(Trim(CStr(ws.Cells(r, scenCol).Value))) = "GRAND TOTAL" Then
            FindScenarioPivotGrandTotalRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindBuyPlanHeaderRow_Audit(ws As Worksheet, startCol As Long, pivotEnd As Long) As Long
    Dim r As Long

    FindBuyPlanHeaderRow_Audit = 0

    For r = pivotEnd + 1 To pivotEnd + 400
        If UCase(Trim(CStr(ws.Cells(r, startCol).Value))) = "BUY PLAN" Then
            FindBuyPlanHeaderRow_Audit = r
            Exit Function
        End If
    Next r
End Function

Private Function FindSellContextTitleRow_Audit(ws As Worksheet, startCol As Long, pivotEnd As Long, buyPlanHeaderRow As Long) As Long
    Dim r As Long

    FindSellContextTitleRow_Audit = 0

    For r = pivotEnd + 1 To buyPlanHeaderRow - 1
        If InStr(1, UCase(Trim(CStr(ws.Cells(r, startCol).Value))), "SELL CONTEXT", vbTextCompare) > 0 Then
            FindSellContextTitleRow_Audit = r
            Exit Function
        End If
    Next r
End Function

Private Function FindLabelRowBetween(ws As Worksheet, colNum As Long, labelText As String, startRow As Long, endRow As Long) As Long
    Dim r As Long

    For r = startRow To endRow
        If UCase(Trim(CStr(ws.Cells(r, colNum).Value))) = UCase(Trim(labelText)) Then
            FindLabelRowBetween = r
            Exit Function
        End If
    Next r
End Function

Private Function CountScenariosLocal(ws As Worksheet) As Long
    On Error GoTo Done

    Dim i As Long
    Dim col As Long

    For i = 0 To 30
        col = ScenStartCol() + i * ScenStride()

        If InStr(1, CStr(ws.Cells(2, col).Value), "Raise $", vbTextCompare) > 0 Then
            CountScenariosLocal = i + 1
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
                          After:=ws.Cells(1, 1), _
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

Private Function FindCellExact(ws As Worksheet, searchText As String) As Range
    On Error Resume Next

    Set FindCellExact = ws.Cells.Find(What:=searchText, _
                                      After:=ws.Cells(1, 1), _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlNext, _
                                      MatchCase:=False)

    On Error GoTo 0
End Function

Private Function FindRowByLabelInColumn(ws As Worksheet, colNum As Long, labelText As String) As Long
    Dim r As Long

    For r = 1 To 500
        If UCase(Trim(CStr(ws.Cells(r, colNum).Value))) = UCase(Trim(labelText)) Then
            FindRowByLabelInColumn = r
            Exit Function
        End If
    Next r
End Function

Private Function IsProcessedCDSReport(ws As Worksheet) As Boolean
    IsProcessedCDSReport = (FindHeaderRow(ws) > 0)
End Function

' ============================================================
' DICTIONARY HELPERS
' ============================================================

Private Function NewDict() As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Set NewDict = d
End Function

Private Sub DictAdd(d As Object, key As String, amount As Double)
    key = UCase(Trim(key))

    If key = "" Then key = "(BLANK)"

    If d.Exists(key) Then
        d(key) = CDbl(d(key)) + amount
    Else
        d.Add key, amount
    End If
End Sub

Private Function DictVal(d As Object, key As String) As Double
    key = UCase(Trim(key))

    If key = "" Then key = "(BLANK)"

    If d.Exists(key) Then
        DictVal = CDbl(d(key))
    Else
        DictVal = 0
    End If
End Function

' ============================================================
' ASSERTIONS + RESULTS
' ============================================================

Private Sub InitAudit()
    Set mRows = New Collection

    mPassCount = 0
    mFailCount = 0
    mWarnCount = 0
    mInfoCount = 0
End Sub

Private Sub AuditNear(checkName As String, actualValue As Variant, expectedValue As Double, tolerance As Double)
    If IsError(actualValue) Then
        AuditLog "FAIL", checkName, "Actual is Excel error. Expected=" & expectedValue
        Exit Sub
    End If

    If Not IsNumeric(actualValue) Then
        AuditLog "FAIL", checkName, "Actual is not numeric: " & CStr(actualValue) & "; Expected=" & expectedValue
        Exit Sub
    End If

    Dim actualDouble As Double
    actualDouble = CDbl(actualValue)

    If Abs(actualDouble - expectedValue) <= tolerance Then
        AuditLog "PASS", checkName, "Actual=" & actualDouble & "; Expected=" & expectedValue
    Else
        AuditLog "FAIL", checkName, "Actual=" & actualDouble & "; Expected=" & expectedValue & "; Tolerance=" & tolerance
    End If
End Sub

Private Sub AuditText(checkName As String, actualValue As Variant, expectedValue As String)
    Dim actualText As String
    Dim expectedText As String

    actualText = UCase(Trim(CStr(actualValue)))
    expectedText = UCase(Trim(expectedValue))

    If actualText = expectedText Then
        AuditLog "PASS", checkName, "Actual=" & actualText
    Else
        AuditLog "FAIL", checkName, "Expected=" & expectedText & "; Actual=" & actualText
    End If
End Sub

Private Sub AuditAssertTrue(checkName As String, condition As Boolean, detail As String)
    If condition Then
        AuditLog "PASS", checkName, detail
    Else
        AuditLog "FAIL", checkName, detail
    End If
End Sub

Private Sub AuditLog(statusText As String, checkName As String, detail As String)
    Dim rec(0 To 3) As Variant

    rec(0) = statusText
    rec(1) = checkName
    rec(2) = detail
    rec(3) = Now

    mRows.Add rec

    Select Case UCase(statusText)
        Case "PASS"
            mPassCount = mPassCount + 1
        Case "FAIL"
            mFailCount = mFailCount + 1
        Case "WARN"
            mWarnCount = mWarnCount + 1
        Case Else
            mInfoCount = mInfoCount + 1
    End Select

    Debug.Print statusText & " | " & checkName & " | " & detail
End Sub

Private Sub WriteAuditResults(sourceWorkbookName As String, sourceSheetName As String)
    Dim outWb As Workbook
    Dim outWs As Worksheet

    Set outWb = Workbooks.Add(xlWBATWorksheet)
    Set outWs = outWb.Worksheets(1)

    outWs.Name = "CDS_MATH_AUDIT"

    outWs.Range("A1").Value = "Source Workbook"
    outWs.Range("B1").Value = sourceWorkbookName
    outWs.Range("A2").Value = "Source Sheet"
    outWs.Range("B2").Value = sourceSheetName
    outWs.Range("A3").Value = "Run Time"
    outWs.Range("B3").Value = Now

    outWs.Range("A5:D5").Value = Array("Status", "Check", "Detail", "Timestamp")

    With outWs.Range("A5:D5")
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim i As Long
    Dim rec As Variant

    For i = 1 To mRows.Count
        rec = mRows(i)

        outWs.Cells(i + 5, 1).Value = rec(0)
        outWs.Cells(i + 5, 2).Value = rec(1)
        outWs.Cells(i + 5, 3).Value = rec(2)
        outWs.Cells(i + 5, 4).Value = rec(3)

        Select Case UCase(CStr(rec(0)))
            Case "PASS"
                outWs.Cells(i + 5, 1).Interior.Color = RGB(198, 239, 206)
            Case "FAIL"
                outWs.Cells(i + 5, 1).Interior.Color = RGB(255, 199, 206)
            Case "WARN"
                outWs.Cells(i + 5, 1).Interior.Color = RGB(255, 242, 204)
            Case Else
                outWs.Cells(i + 5, 1).Interior.Color = RGB(221, 235, 247)
        End Select
    Next i

    outWs.Columns("A:D").AutoFit
    outWs.Activate
End Sub

Private Function Num(v As Variant) As Double
    If IsError(v) Then
        Num = 0
    ElseIf IsNumeric(v) Then
        Num = CDbl(v)
    Else
        Num = 0
    End If
End Function

